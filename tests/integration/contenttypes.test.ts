import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { apply, IApplyResult } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import { describeError, redact } from '../../debug/lib/redact'
import {
  IContentTypeState,
  readContentType
} from '../../debug/lib/readback/site'
import { loadFixture } from '../fixtures/fixtures'
import {
  applyTwice,
  assertClean,
  recording,
  skipReason,
  withScratchWeb
} from './helpers'

/** The handlers the fixture needs: its columns, its content types, the list bound to one. */
const HANDLERS = ['SiteFields', 'ContentTypes', 'Lists']

const PARENT = '0x0100ED0C3F94FCE968FD1B421031A3141A32'
const CHILD = `${PARENT}01`
const DOCUMENT = '0x010100A0D1F01C823B33452A95A801640901B9'
const LIST = 'SPJS Itest Ct Liste'
const CSOM = 'POST /_vti_bin/client.svc/ProcessQuery'

/** What the suite reads back: the fixture's content types and the columns of its list. */
interface IState {
  parent: IContentTypeState
  child: IContentTypeState
  document: IContentTypeState
  /** Internal names of the list's columns that the fixture's site columns gave it, sorted. */
  listFields: string[]
}

/**
 * The lines a run logged at a level, redacted (they reach JUnit).
 *
 * @param run - The run
 * @param level - `errors` or `warnings`
 */
const lines = (run: IApplyResult, level: 'errors' | 'warnings') =>
  run[level].map((entry) => redact(entry.message))

const names = (ct: IContentTypeState) => ct.fieldLinks.map((link) => link.Name)

/**
 * A field link's flags, as `Name` with `(required)` and `(hidden)`.
 *
 * @param ct - The content type
 * @param name - The link
 */
function flags(ct: IContentTypeState, name: string): string {
  const link = ct.fieldLinks.find((candidate) => candidate.Name === name)
  if (!link) return `${name} (missing)`
  return `${name}${link.Required ? ' (required)' : ''}${
    link.Hidden ? ' (hidden)' : ''
  }`
}

/**
 * The ContentTypes handler on a scratch web (docs/specs/jsom-exit.md): content types with the ids
 * the template gives them (a child of Item, a child of that one, a child of Document), created
 * through CSOM over HTTP, their field links with `Required`, `Hidden` and an order of their own,
 * and only what differs written on a later run. `Update(true)` pushes a change to the content
 * types and lists that inherit from one, which the child content type and the bound list show.
 */
describe('ContentTypes', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('contenttypes')

  /**
   * Reads back the fixture's content types and the list's columns.
   *
   * @param conn - The scratch web
   */
  async function read(conn: ISiteConnection): Promise<IState> {
    const [parent, child, document] = await Promise.all(
      [PARENT, CHILD, DOCUMENT].map((id) => readContentType(conn, id))
    )
    const fields = await conn.web.lists
      .getByTitle(LIST)
      .fields.select('InternalName')
      .filter("startswith(InternalName,'SPJSItestCt')")<
      Array<{ InternalName: string }>
    >()
    return {
      parent,
      child,
      document,
      listFields: fields.map((field) => field.InternalName).sort()
    }
  }

  it('creates content types with the ids of the template, their links, flags and order, and an unchanged second apply writes nothing', async (t) => {
    const { conn } = scratch()
    const requests: string[] = []
    const marks: number[] = []
    const { afterFirst, second } = await applyTwice(
      recording(conn, requests),
      () => loadFixture('contenttypes'),
      // Read through the plain connection: only the engine's requests are recorded.
      () => {
        marks.push(requests.length)
        return read(conn)
      },
      { handlers: HANDLERS }
    )
    const fixture = loadFixture('contenttypes').ContentTypes
    const { parent, child, document } = afterFirst

    // The ids are the template's, on this web (REST makes a child of Item with an id of its own).
    assert.deepEqual(
      [parent, child, document].map((ct) => [
        ct.Id,
        ct.Name,
        ct.Group,
        ct.onWeb
      ]),
      // The read-back has ids in upper case.
      fixture.map((ct: any) => [ct.ID.toUpperCase(), ct.Name, ct.Group, true])
    )
    assert.equal(parent.Description, fixture[0].Description)
    assert.equal(document.Description, fixture[2].Description)

    // The parent: its four links first, in the template's order, with the flags it gives.
    assert.deepEqual(names(parent).slice(0, 4), [
      'SPJSItestCtNumber',
      'Title',
      'SPJSItestCtText',
      'SPJSItestCtChoice'
    ])
    assert.deepEqual(
      names(parent)
        .slice(0, 4)
        .map((name) => flags(parent, name)),
      [
        'SPJSItestCtNumber (required)',
        'Title (required) (hidden)',
        'SPJSItestCtText',
        'SPJSItestCtChoice'
      ]
    )
    t.diagnostic(`parent links: ${names(parent).join(', ')}`)

    // The child (`<parent id>01`) has its parent's links as they were when it was made, and its
    // own, named without an id.
    for (const name of names(parent))
      assert.ok(names(child).includes(name), `the child lacks ${name}`)
    assert.ok(names(child).includes('SPJSItestCtNote'))
    assert.equal(
      flags(child, 'SPJSItestCtNumber'),
      'SPJSItestCtNumber (required)'
    )
    t.diagnostic(`child links: ${names(child).join(', ')}`)

    // A child of Document, its column named by id only.
    assert.ok(names(document).includes('FileLeafRef'), names(document).join())
    assert.ok(names(document).includes('SPJSItestCtDate'))

    // The list bound to the parent got its columns.
    assert.deepEqual(afterFirst.listFields, [
      'SPJSItestCtChoice',
      'SPJSItestCtNumber',
      'SPJSItestCtText'
    ])

    // The second run read, compared and sent no CSOM request: an `Update(true)` pushes a content
    // type to everything that inherits from it, changed or not.
    const secondRun = requests.slice(marks[0], marks[1])
    assert.ok(secondRun.length > 0, 'the second run was not recorded')
    assert.deepEqual(
      secondRun.filter((request) => request.startsWith(CSOM)),
      [],
      'the second run sent CSOM requests'
    )
    for (const ct of fixture)
      assert.ok(
        second.entries.some((entry) =>
          entry.message.endsWith(
            `Content type [${ct.Name}] (${ct.ID}) is unchanged.`
          )
        ),
        `${ct.Name} is not logged as unchanged`
      )
  })

  it('applies a new name, a new first link and a changed Required (B variant), the child and the list follow, and back', async (t) => {
    const { conn } = scratch()
    assertClean(
      await apply(conn, loadFixture('contenttypes'), { handlers: HANDLERS }),
      'the A run'
    )
    const a = await read(conn)

    const requests: string[] = []
    const marks: number[] = []
    const { afterFirst: b } = await applyTwice(
      recording(conn, requests),
      () => loadFixture('contenttypes.b'),
      () => {
        marks.push(requests.length)
        return read(conn)
      },
      { handlers: HANDLERS }
    )
    const fixture = loadFixture('contenttypes.b').ContentTypes[0]

    assert.equal(b.parent.Id, a.parent.Id)
    assert.deepEqual(
      [b.parent.Name, b.parent.Description],
      [fixture.Name, fixture.Description]
    )
    assert.deepEqual(names(b.parent).slice(0, 5), [
      'SPJSItestCtDate',
      'SPJSItestCtNumber',
      'Title',
      'SPJSItestCtText',
      'SPJSItestCtChoice'
    ])
    assert.equal(
      flags(a.parent, 'SPJSItestCtNumber'),
      'SPJSItestCtNumber (required)'
    )
    assert.equal(flags(b.parent, 'SPJSItestCtNumber'), 'SPJSItestCtNumber')
    // One request changed the parent; the others were left alone.
    assert.equal(
      requests.slice(0, marks[0]).filter((request) => request.startsWith(CSOM))
        .length,
      1,
      'B changed more than the parent'
    )
    assert.deepEqual(
      requests
        .slice(marks[0], marks[1])
        .filter((request) => request.startsWith(CSOM)),
      [],
      'B again sent CSOM requests'
    )
    assert.deepEqual(b.document, a.document)

    // Update(true): the child and the list's copy follow their parent.
    assert.ok(names(b.child).includes('SPJSItestCtDate'), names(b.child).join())
    assert.equal(flags(b.child, 'SPJSItestCtNumber'), 'SPJSItestCtNumber')
    assert.ok(names(b.child).includes('SPJSItestCtNote'))
    assert.deepEqual(b.listFields, [
      'SPJSItestCtChoice',
      'SPJSItestCtDate',
      'SPJSItestCtNumber',
      'SPJSItestCtText'
    ])
    t.diagnostic(`child links after B: ${names(b.child).join(', ')}`)

    // A again: the name, the description and Required come back; a link the template does not
    // name is never removed, and goes after the ones it names.
    assertClean(
      await apply(conn, loadFixture('contenttypes'), { handlers: HANDLERS }),
      'the A run after B'
    )
    const back = await read(conn)
    assert.deepEqual(
      [back.parent.Name, back.parent.Description],
      [a.parent.Name, a.parent.Description]
    )
    assert.deepEqual(names(back.parent).slice(0, 5), [
      'SPJSItestCtNumber',
      'Title',
      'SPJSItestCtText',
      'SPJSItestCtChoice',
      'SPJSItestCtDate'
    ])
    assert.equal(
      flags(back.parent, 'SPJSItestCtNumber'),
      'SPJSItestCtNumber (required)'
    )
  })

  it('logs a field ref whose column the web does not have, and links the others', async () => {
    const { conn } = scratch()
    const id = '0x01009D4E1A7C52B84F60A3C1D8E7F6059B2A'
    const template = () => ({
      ContentTypes: [
        {
          ID: id,
          Name: 'SPJS Itest Mangler kolonne',
          Group: 'SPJS Itest',
          FieldRefs: [
            { Name: 'SPJSItestCtText' },
            { Name: 'SPJSItestCtFinnesIkke' },
            { Name: 'SPJSItestCtNumber' }
          ]
        }
      ]
    })
    assertClean(
      await apply(conn, loadFixture('contenttypes'), {
        handlers: ['SiteFields']
      }),
      'the columns'
    )

    const run = await apply(conn, template(), { handlers: ['ContentTypes'] })

    assert.equal(run.error, undefined, run.error && describeError(run.error))
    assert.equal(run.errors.length, 1, lines(run, 'errors').join('\n'))
    assert.match(
      lines(run, 'errors')[0],
      /Field ref SPJSItestCtFinnesIkke of content type \[SPJS Itest Mangler kolonne\] \(0x01009D4E1A7C52B84F60A3C1D8E7F6059B2A\) is left out: the web has no site column with that name\./
    )
    const stored = await readContentType(conn, id)
    assert.deepEqual(names(stored).slice(0, 2), [
      'SPJSItestCtText',
      'SPJSItestCtNumber'
    ])
  })

  it('stops with what SharePoint said when a content type cannot be created', async (t) => {
    const { conn } = scratch()
    assertClean(
      await apply(conn, loadFixture('contenttypes'), { handlers: HANDLERS }),
      'the A run'
    )
    const taken = (await readContentType(conn, PARENT)).Name
    const id = '0x01003B7F2C9E6A1D4E85B0C4F7A2D9E8136C'

    // A name the web already has, under another id: SharePoint refuses the duplicate.
    const run = await apply(
      conn,
      { ContentTypes: [{ ID: id, Name: taken, Group: 'SPJS Itest' }] },
      { handlers: ['ContentTypes'] }
    )

    assert.ok(run.error, 'the run did not stop')
    assert.equal(run.error.handler, 'ContentTypes')
    const message = redact(describeError(run.error))
    t.diagnostic(message)
    assert.match(
      message,
      /^Failed to create content type \[SPJS Itest Prosjekt[^\]]*\] \(0x01003B7F2C9E6A1D4E85B0C4F7A2D9E8136C\): .+ \(Microsoft\.SharePoint\.\w+\)/
    )
    assert.equal(await readContentType(conn, id), null)
  })
})
