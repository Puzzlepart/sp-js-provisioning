import assert from 'node:assert/strict'
import { createHash } from 'crypto'
import * as http from 'http'
import { after, before, describe, it } from 'node:test'
import { apply, IApplyResult } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import { redact } from '../../debug/lib/redact'
import { loadFixture } from '../fixtures/fixtures'
import {
  applyTwice,
  assertClean,
  recording,
  skipReason,
  withScratchWeb
} from './helpers'

const HANDLERS = ['Files']
const LIBRARY = 'SiteAssets'
const SOURCES = 'SPJS Itest kilder'

/** Text with characters outside ASCII and a CRLF: `response.text()` and a text blob changed them. */
const TEXT = Buffer.from('SPJS Itest: æøå ÆØÅ – «sitat»\r\nlinje 2\n', 'utf8')
/** Bytes that are not text: a PNG signature, zeros and bytes above 0x7F. */
const BINARY = Buffer.from([
  0x89,
  0x50,
  0x4e,
  0x47,
  0x0d,
  0x0a,
  0x1a,
  0x0a,
  0x00,
  0x00,
  0xff,
  0xfe,
  0x80,
  0x7f,
  0xc3,
  0x28
])

/** A local file server: another origin than SharePoint's, answering without credentials. */
interface IFileServer {
  url: string
  files: Record<string, Buffer>
  requests: string[]
  close(): Promise<void>
}

/**
 * Starts a file server on a free port on 127.0.0.1. The handler runs in this process, so it is
 * reachable from a run against SharePoint Online; a path it does not have answers 404.
 *
 * @param files - The files, by path
 */
function startFileServer(files: Record<string, Buffer>): Promise<IFileServer> {
  const requests: string[] = []
  const server = http.createServer((request, response) => {
    requests.push(`${request.method} ${request.url}`)
    const body = files[request.url || '']
    if (!body) {
      response.writeHead(404)
      response.end()
      return
    }
    response.writeHead(200, { 'Access-Control-Allow-Origin': '*' })
    response.end(body)
  })
  return new Promise((resolve, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () =>
      resolve({
        url: `http://127.0.0.1:${server.address().port}`,
        files,
        requests,
        close: () => new Promise((done) => server.close(() => done()))
      })
    )
  })
}

/** A file as the suite reads it back: a hash of its bytes, its size and its item's title. */
interface IFileState {
  sha256: string
  size: number
  title: string
}

const hash = (bytes: Buffer | ArrayBuffer) =>
  createHash('sha256')
    .update(Buffer.from(bytes as ArrayBuffer))
    .digest('hex')

/**
 * The lines a run logged at a level, redacted (they reach JUnit).
 *
 * @param run - The run
 */
const errorsOf = (run: IApplyResult) =>
  run.errors.map((entry) => redact(entry.message))

/**
 * The Files handler on a scratch web (docs/specs/jsom-exit.md, J7): files uploaded through REST
 * from a source on the web's origin (read through the web, signed in) and from another origin (a
 * local server, read without credentials), byte for byte; a list item property; `Overwrite: false`;
 * an unchanged re-apply that uploads nothing; a failing source, logged, and the next file still
 * provisioned; classic web parts refused with an error.
 */
describe('Files', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('files')
  let server: IFileServer

  before(async () => {
    server = await startFileServer({
      '/spjs-itest-bilde.png': BINARY,
      '/spjs-itest-tekst.txt': TEXT
    })
    // The same-origin source: a file in the scratch web, put there outside the engine.
    const { conn } = scratch()
    const root = conn.webInfo.ServerRelativeUrl
    await conn.web.folders.addUsingPath(`${root}/${LIBRARY}/${SOURCES}`)
    await conn.web
      .getFolderByServerRelativePath(`${root}/${LIBRARY}/${SOURCES}`)
      .files.addUsingPath('spjs-itest-kilde.txt', TEXT, { Overwrite: true })
  })

  after(async () => {
    if (server) await server.close()
  })

  /** The values the fixture's `{parameter:}` tokens stand for. */
  const parameters = (conn: ISiteConnection) => ({
    SPJSItestFilesWeb: conn.webInfo.Url,
    SPJSItestFilesServer: server.url
  })

  /**
   * Reads back the fixture's target files.
   *
   * @param conn - The scratch web
   * @param names - The file names in the library
   */
  const reader = (names: string[]) => async (
    conn: ISiteConnection
  ): Promise<Record<string, IFileState>> => {
    const state: Record<string, IFileState> = {}
    for (const name of names) {
      const file = conn.web.getFileByServerRelativePath(
        `${conn.webInfo.ServerRelativeUrl}/${LIBRARY}/${name}`
      )
      const bytes = await file.getBuffer()
      const item = await file.listItemAllFields.select('Title')<{
        Title: string | null
      }>()
      state[name] = {
        sha256: hash(bytes),
        size: bytes.byteLength,
        title: item.Title || ''
      }
    }
    return state
  }
  const NAMES = [
    'SPJS Itest tekst.txt',
    'SPJS Itest bilde.png',
    'SPJS Itest beholdt.txt'
  ]

  it('uploads files from the web and from another origin byte for byte, sets a property, and an unchanged second apply uploads nothing', async () => {
    const { conn } = scratch()
    const requests: string[] = []
    const marks: number[] = []
    const { afterFirst } = await applyTwice(
      recording(conn, requests),
      () => loadFixture('files'),
      async (target) => {
        marks.push(requests.length)
        return reader(NAMES)(target)
      },
      { handlers: HANDLERS, parameters: parameters(conn) }
    )

    assert.deepEqual(afterFirst, {
      'SPJS Itest tekst.txt': {
        sha256: hash(TEXT),
        size: TEXT.length,
        title: 'SPJS Itest tekstfil'
      },
      'SPJS Itest bilde.png': {
        sha256: hash(BINARY),
        size: BINARY.length,
        title: ''
      },
      'SPJS Itest beholdt.txt': {
        sha256: hash(TEXT),
        size: TEXT.length,
        title: ''
      }
    })
    // The other origin was asked without credentials (the server saw the requests).
    assert.ok(server.requests.includes('GET /spjs-itest-bilde.png'))

    const writes = (slice: string[]) =>
      slice.filter((request) => /^(POST|MERGE) /.test(request))
    assert.equal(
      writes(requests.slice(0, marks[0])).length,
      4,
      'the first run: three uploads and one property'
    )
    assert.deepEqual(
      writes(requests.slice(marks[0], marks[1])),
      [],
      'the second run wrote'
    )
  })

  it('uploads a changed source again, and leaves a file with Overwrite false as it is', async () => {
    const { conn } = scratch()
    assertClean(
      await apply(conn, loadFixture('files'), {
        handlers: HANDLERS,
        parameters: parameters(conn)
      }),
      'the first run'
    )
    const changed = Buffer.concat([BINARY, Buffer.from([0x01, 0x02])])
    const changedText = Buffer.from('SPJS Itest: endret\r\n', 'utf8')
    server.files['/spjs-itest-bilde.png'] = changed
    server.files['/spjs-itest-tekst.txt'] = changedText
    try {
      assertClean(
        await apply(conn, loadFixture('files'), {
          handlers: HANDLERS,
          parameters: parameters(conn)
        }),
        'the run with changed sources'
      )
      const state = await reader(NAMES)(conn)
      assert.equal(state['SPJS Itest bilde.png'].sha256, hash(changed))
      assert.equal(
        state['SPJS Itest beholdt.txt'].sha256,
        hash(TEXT),
        'Overwrite false replaced the file'
      )
    } finally {
      server.files['/spjs-itest-bilde.png'] = BINARY
      server.files['/spjs-itest-tekst.txt'] = TEXT
    }
  })

  it('logs a file whose source fails and one with classic web parts, and still provisions the next', async () => {
    const { conn } = scratch()
    const run = await apply(
      conn,
      {
        Files: [
          {
            Folder: LIBRARY,
            Src: '{parameter:SPJSItestFilesServer}/finnes-ikke.txt',
            Url: 'SPJS Itest mangler.txt'
          },
          {
            Folder: LIBRARY,
            Src: '{parameter:SPJSItestFilesServer}/spjs-itest-tekst.txt',
            Url: 'SPJS Itest med webdeler.txt',
            WebParts: [
              {
                Title: 'SPJS Itest',
                Zone: 'Main',
                Order: 1,
                Contents: { Xml: '<webParts />' }
              }
            ]
          }
        ]
      },
      { handlers: HANDLERS, parameters: parameters(conn) }
    )

    assert.equal(run.error, undefined)
    const errors = errorsOf(run)
    assert.equal(errors.length, 2, errors.join('\n'))
    assert.match(
      errors[0],
      /Failed to provision file SiteAssets\/SPJS Itest mangler\.txt: Failed to read http:\/\/127\.0\.0\.1:\d+\/finnes-ikke\.txt: 404/
    )
    assert.match(
      errors[1],
      /File SiteAssets\/SPJS Itest med webdeler\.txt: classic web parts \(WebParts, RemoveExistingWebParts\) are no longer supported/
    )
    const state = await reader(['SPJS Itest med webdeler.txt'])(conn)
    assert.equal(state['SPJS Itest med webdeler.txt'].sha256, hash(TEXT))
  })
})
