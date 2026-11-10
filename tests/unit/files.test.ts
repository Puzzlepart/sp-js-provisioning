import assert from 'node:assert/strict'
import { afterEach, beforeEach, describe, it } from 'node:test'
import { LogLevel } from '@pnp/logging'
import { Files } from '../../src/handlers/files'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { IFileObject } from '../../src/schema'
import { captureLogs, levelOf } from '../helpers/logs'
import { createWireWeb, IWireRequest } from '../helpers/transport'

const WEB = 'https://contoso.sharepoint.com/sites/demo'
const TARGET =
  "_api/web/getFileByServerRelativePath(decodedUrl='/sites/demo/SiteAssets/a.png')"
const UPLOAD =
  "_api/web/getFolderByServerRelativePath(decodedUrl='/sites/demo/SiteAssets')/files/AddUsingPath(decodedurl='a.png',Overwrite=true)"

/** Bytes that are not text: a PNG signature, a zero byte and bytes above 0x7F. */
const BINARY = new Uint8Array([
  0x89,
  0x50,
  0x4e,
  0x47,
  0x0d,
  0x0a,
  0x1a,
  0x0a,
  0x00,
  0xff,
  0xfe,
  0x80
])

const notFound = () => Object.assign(new Error('Not Found'), { status: 404 })

/** An SPFx context stand-in for the URL tokens (`{site}`, `{sitecollection}`). */
const spfxContext = {
  pageContext: {
    site: { absoluteUrl: WEB, serverRelativeUrl: '/sites/demo' },
    web: { absoluteUrl: WEB }
  }
}

/**
 * A web with a SiteAssets library: the file there (`stored`, or none), sources on its origin, and
 * the uploads and item updates it is sent.
 *
 * @param options - What the web has, and what fails
 */
function createSite(
  options: {
    stored?: Uint8Array
    sources?: Record<string, Uint8Array>
    item?: Record<string, any>
    uploadFails?: boolean
  } = {}
) {
  const uploads: Uint8Array[] = []
  const { web, requests } = createWireWeb((request: IWireRequest) => {
    if (request.path === '_api/web')
      return { ServerRelativeUrl: '/sites/demo', Url: WEB }
    if (request.path === `${TARGET}/$value`) {
      if (!options.stored) throw notFound()
      return options.stored.slice().buffer
    }
    if (options.sources && options.sources[request.path])
      return options.sources[request.path].slice().buffer
    if (request.path === UPLOAD) {
      if (options.uploadFails)
        throw Object.assign(new Error('Ingen tilgang'), { status: 403 })
      return { ServerRelativeUrl: '/sites/demo/SiteAssets/a.png' }
    }
    if (request.path === `${TARGET}/listItemAllFields`)
      return { ID: 7, ParentList: { Id: 'list-1' }, ...(options.item || {}) }
    if (request.path.startsWith('SiteAssets/missing')) throw notFound()
    return undefined
  })
  // The upload's body is the bytes: kept to compare.
  web.on.pre(async (url: string, init: any, result: any) => {
    if (url.includes('AddUsingPath') && init && init.body)
      uploads.push(new Uint8Array(init.body as ArrayBuffer))
    return [url, init, result]
  })
  return { web, requests, uploads }
}

/**
 * Runs the handler and returns what it logged.
 *
 * @param site - The fake web
 * @param files - The template's files
 */
async function provision(
  site: ReturnType<typeof createSite>,
  files: Array<Partial<IFileObject>>
) {
  const { entries } = await captureLogs(() =>
    new Files({ spfxContext }).ProvisionObjects(
      site.web,
      files as IFileObject[],
      new ProvisioningContext()
    )
  )
  return {
    entries,
    errors: entries
      .filter((entry) => entry.level === LogLevel.Error)
      .map((entry) => entry.message.replace(/^.*?\): \(\w+\): /, ''))
  }
}

const file = (members: Partial<IFileObject> = {}): Partial<IFileObject> => ({
  Folder: 'SiteAssets',
  Src: '{site}/SiteAssets/source/a.png',
  Url: 'a.png',
  ...members
})

const lines = (requests: IWireRequest[]) =>
  requests.map((request) => `${request.method} ${request.path}`)

describe('Files', () => {
  const realFetch = globalThis.fetch
  let fetched: string[]
  beforeEach(() => {
    fetched = []
    globalThis.fetch = (async (url: string, init: any) => {
      fetched.push(
        `${
          init && init.credentials ? init.credentials : 'no credentials'
        } ${url}`
      )
      if (url.includes('missing'))
        return new Response('Not found', {
          status: 404,
          statusText: 'Not Found'
        })
      return new Response(BINARY.slice().buffer, { status: 200 })
    }) as any
  })
  afterEach(() => {
    globalThis.fetch = realFetch
  })

  it("reads a source on the web's origin through the web, as bytes, and uploads them unchanged", async () => {
    const site = createSite({ sources: { 'SiteAssets/source/a.png': BINARY } })

    const { errors } = await provision(site, [file()])

    assert.deepEqual(errors, [])
    assert.deepEqual(lines(site.requests), [
      'GET _api/web',
      `GET ${TARGET}/$value`,
      'GET SiteAssets/source/a.png',
      `POST ${UPLOAD}`
    ])
    assert.deepEqual(site.uploads, [BINARY])
    assert.deepEqual(fetched, [])
  })

  it('reads a source on another origin with fetch and no credentials', async () => {
    const site = createSite()

    await provision(site, [
      file({ Src: 'https://raw.githubusercontent.com/org/repo/main/a.png' })
    ])

    assert.deepEqual(fetched, [
      'no credentials https://raw.githubusercontent.com/org/repo/main/a.png'
    ])
    assert.deepEqual(site.uploads, [BINARY])
  })

  it('does not upload a file whose contents are already the source', async () => {
    const site = createSite({
      stored: BINARY,
      sources: { 'SiteAssets/source/a.png': BINARY }
    })

    const { entries } = await provision(site, [file()])

    assert.deepEqual(site.uploads, [])
    assert.ok(!lines(site.requests).includes(`POST ${UPLOAD}`))
    assert.equal(
      levelOf(entries, /File SiteAssets\/a\.png is unchanged\./),
      LogLevel.Info
    )
  })

  it('uploads a file whose contents differ, and leaves one with Overwrite false as it is', async () => {
    const changed = createSite({
      stored: new Uint8Array([1, 2, 3]),
      sources: { 'SiteAssets/source/a.png': BINARY }
    })
    await provision(changed, [file()])
    assert.deepEqual(changed.uploads, [BINARY])

    const kept = createSite({
      stored: new Uint8Array([1, 2, 3]),
      sources: { 'SiteAssets/source/a.png': BINARY }
    })
    const { entries } = await provision(kept, [
      file({ Overwrite: false, Properties: { Title: 'Ny' } })
    ])
    assert.deepEqual(lines(kept.requests), [
      'GET _api/web',
      `GET ${TARGET}/$value`
    ])
    assert.equal(
      levelOf(
        entries,
        /File SiteAssets\/a\.png exists and is left as it is \(Overwrite is false\)\./
      ),
      LogLevel.Info
    )
  })

  it('sets only the properties whose value differs', async () => {
    const site = createSite({
      stored: BINARY,
      sources: { 'SiteAssets/source/a.png': BINARY },
      item: { Title: 'Logo', GtOrder: 2 }
    })

    await provision(site, [file({ Properties: { Title: 'Logo', GtOrder: 3 } })])

    const reads = site.requests.filter((request) =>
      request.path.endsWith('listItemAllFields')
    )
    assert.deepEqual(
      reads.map((request) => [request.query.$select, request.query.$expand]),
      [['ID,ParentList/Id,Title,GtOrder', 'ParentList']]
    )
    const merges = site.requests.filter((request) => request.method === 'MERGE')
    assert.deepEqual(
      merges.map((request) => [request.path, request.body]),
      [["_api/web/lists('list-1')/items(7)", { GtOrder: 3 }]]
    )

    const same = createSite({
      stored: BINARY,
      sources: { 'SiteAssets/source/a.png': BINARY },
      item: { Title: 'Logo', GtOrder: 3 }
    })
    await provision(same, [file({ Properties: { Title: 'Logo', GtOrder: 3 } })])
    assert.deepEqual(
      same.requests.filter((request) => request.method !== 'GET'),
      []
    )
  })

  it('logs a file that fails, naming it, and provisions the next', async () => {
    const site = createSite({ sources: { 'SiteAssets/source/a.png': BINARY } })

    const { errors } = await provision(site, [
      file({ Src: '{site}/SiteAssets/missing.png', Url: 'b.png' }),
      file({ Src: 'https://example.org/missing.png', Url: 'c.png' }),
      file()
    ])

    assert.equal(errors.length, 2, errors.join('\n'))
    assert.match(errors[0], /^Failed to provision file SiteAssets\/b\.png: /)
    assert.equal(
      errors[1],
      'Failed to provision file SiteAssets/c.png: Failed to read https://example.org/missing.png: 404 Not Found'
    )
    assert.deepEqual(site.uploads, [BINARY])
  })

  it('logs a refused upload as an error, and does not throw', async () => {
    const site = createSite({
      sources: { 'SiteAssets/source/a.png': BINARY },
      uploadFails: true
    })

    const { errors } = await provision(site, [file()])

    assert.deepEqual(errors, [
      'Failed to provision file SiteAssets/a.png: Ingen tilgang'
    ])
  })

  it('runs under an SPFx context, and refuses classic web parts with an error while the file is provisioned', async () => {
    const site = createSite({ sources: { 'SiteAssets/source/a.png': BINARY } })

    const { errors } = await provision(site, [
      file({
        RemoveExistingWebParts: true,
        WebParts: [
          {
            Title: 'Gammel',
            Zone: 'Main',
            Order: 1,
            Contents: { Xml: '<webParts />' }
          }
        ]
      })
    ])

    assert.deepEqual(errors, [
      'File SiteAssets/a.png: classic web parts (WebParts, RemoveExistingWebParts) are no longer supported and are left out; the file itself is provisioned.'
    ])
    assert.deepEqual(site.uploads, [BINARY])
  })
})
