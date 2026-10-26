import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import { describeError } from '../../debug/lib/redact'
import { ISiteConnection } from '../../debug/lib/connect'
import {
  fieldSchema,
  IFieldSchema,
  ISiteFieldsState,
  ISiteFieldState,
  readSiteFields
} from '../../debug/lib/readback/sitefields'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

/** The list the fixture's Lookup column points at, created before the first run. */
const SOURCE = 'SPJS Itest Sitefields Source'

/** The children of `<Field>` the read-back keeps. */
const PARTS: Array<Exclude<keyof IFieldSchema, 'attributes'>> = [
  'Default',
  'Formula',
  'Choices',
  'FieldRefs'
]

/** A read-back without the `SchemaXml` versions, which every update bumps. */
type IStable = Omit<ISiteFieldsState, 'versions'>

/**
 * The internal names of a fixture's columns, in template order.
 *
 * @param fixture - The template
 */
function namesOf(fixture: Record<string, any>): string[] {
  return (fixture.SiteFields as string[]).map(
    (xml) => fieldSchema(xml).attributes.Name
  )
}

/**
 * The SiteFields handler on a scratch web: site columns of eleven types from field XML, `{sitename}`
 * in a description and `{listid:}` in a Lookup column's `List`, created on the first run and
 * updated in place (same ids, `SchemaXml` version bumped) on every later run.
 */
describe('SiteFields', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('sitefields')
  let sourceId: string

  before(async () => {
    const { conn } = scratch()
    // `{listid:}` resolves only for lists that exist when the run starts (matrix).
    await conn.web.lists.add(
      SOURCE,
      'Lookup source for the SiteFields suite',
      100,
      false
    )
    sourceId = (
      await conn.web.lists.getByTitle(SOURCE).select('Id')<{ Id: string }>()
    ).Id.toLowerCase()
  })

  /**
   * The fixture's field XML with the values the engine's tokens stand for on the scratch web.
   *
   * @param conn - The scratch web
   * @param xml - A column of the fixture
   */
  const resolved = (conn: ISiteConnection, xml: string) =>
    xml
      .replace(/{sitename}/g, conn.webInfo.Title)
      .replace(`{listid:${SOURCE}}`, sourceId)

  /**
   * Asserts that every column of a fixture reads back as declared: id, type, title, group,
   * required, hidden, description, every attribute of the XML and its `Default`, `Formula`,
   * `CHOICES` and `FieldRefs`; and that the group holds those columns and no others.
   *
   * @param conn - The scratch web
   * @param state - The read-back
   * @param fixture - The template
   */
  const assertDeclared = (
    conn: ISiteConnection,
    state: IStable,
    fixture: Record<string, any>
  ) => {
    for (const xml of fixture.SiteFields as string[]) {
      const want = fieldSchema(resolved(conn, xml))
      const name = want.attributes.Name
      const field = state.fields.find((f) => f.InternalName === name)
      assert.ok(field, `${name} was not read back`)
      assert.deepEqual(
        {
          Id: field.Id,
          TypeAsString: field.TypeAsString,
          Title: field.Title,
          Group: field.Group,
          Required: field.Required,
          Hidden: field.Hidden,
          Description: field.Description
        },
        {
          Id: want.attributes.ID,
          TypeAsString: want.attributes.Type,
          Title: want.attributes.DisplayName,
          Group: want.attributes.Group,
          Required: want.attributes.Required === 'TRUE',
          Hidden: want.attributes.Hidden === 'TRUE',
          Description: want.attributes.Description ?? ''
        },
        `${name}: properties`
      )
      // SharePoint adds attributes (SourceID, ColName, ...): the declared ones must be there as
      // declared.
      const declared: Record<string, string> = {}
      for (const attribute of Object.keys(want.attributes))
        declared[attribute] = field.schema.attributes[attribute]
      assert.deepEqual(
        declared,
        want.attributes,
        `${name}: SchemaXml attributes`
      )
      for (const part of PARTS)
        if (want[part] !== undefined)
          assert.deepEqual(field.schema[part], want[part], `${name}: ${part}`)
    }
    assert.deepEqual(
      state.inGroup,
      namesOf(fixture).sort(),
      'the group holds exactly the template columns (no duplicates)'
    )
  }

  /**
   * Reads back the fixture's columns, keeping each run's `SchemaXml` versions apart.
   *
   * @param names - The columns
   * @param versions - Collects the versions of each read
   */
  const reader = (
    names: string[],
    versions: Array<Record<string, number>>
  ) => async (conn: ISiteConnection): Promise<IStable> => {
    const { versions: read, ...state } = await readSiteFields(conn, names)
    versions.push(read)
    return state
  }

  const byName = (state: IStable, name: string): ISiteFieldState =>
    state.fields.find((field) => field.InternalName === name)

  it('creates eleven column types from field XML as declared, and updates them in place on a second apply', async () => {
    const { conn } = scratch()
    const names = namesOf(loadFixture('sitefields'))
    const versions: Array<Record<string, number>> = []
    const { afterFirst } = await applyTwice(
      conn,
      () => loadFixture('sitefields'),
      reader(names, versions)
    )
    assertDeclared(conn, afterFirst, loadFixture('sitefields'))

    // The tokens: `{sitename}` in a description, `{listid:}` in a Lookup column's list.
    assert.equal(
      byName(afterFirst, 'SPJSItestSfText').Description,
      `Text column on ${conn.webInfo.Title}`
    )
    assert.equal(
      byName(afterFirst, 'SPJSItestSfLookup').schema.attributes.List,
      sourceId
    )
    // The second run took the update path for every column: same ids (applyTwice compared the
    // read-backs), and SharePoint bumped each column's SchemaXml version.
    for (const name of names)
      assert.ok(
        versions[1][name] > versions[0][name],
        `${name}: version ${versions[0][name]} -> ${versions[1][name]}`
      )
  })

  it('applies a changed Title and Required to the same column (B variant), and back', async () => {
    const { conn } = scratch()
    const names = namesOf(loadFixture('sitefields'))
    const versions: Array<Record<string, number>> = []
    const read = reader(names, versions)
    assertClean(await apply(conn, loadFixture('sitefields')), 'the A run')
    const a = await read(conn)

    const { afterFirst: b } = await applyTwice(
      conn,
      () => loadFixture('sitefields.b'),
      read
    )
    assertDeclared(conn, b, loadFixture('sitefields.b'))
    const text = 'SPJSItestSfText'
    assert.equal(byName(b, text).Id, byName(a, text).Id)
    assert.deepEqual(
      [byName(a, text).Title, byName(a, text).Required],
      ['SPJS Itest Text', false]
    )
    assert.deepEqual(
      [byName(b, text).Title, byName(b, text).Required],
      ['SPJS Itest Text Renamed', true]
    )
    assert.deepEqual(
      b.fields.filter((field) => field.InternalName !== text),
      a.fields.filter((field) => field.InternalName !== text),
      'B changes only the Text column'
    )

    assertClean(
      await apply(conn, loadFixture('sitefields')),
      'the A run after B'
    )
    assert.deepEqual(await read(conn), a, 'A again restores the Text column')
  })

  it(
    'resolves {listid:} in a lookup to a list from the same template on the first run',
    {
      todo:
        'SiteFields (sort 0) runs before Lists (3), and {listid:} resolves only for lists that ' +
        'exist when the run starts (matrix): SharePoint stores the literal token as the ' +
        "lookup's List, nothing is logged, and only a second run points it at the list"
    },
    async (t) => {
      const { conn } = scratch()
      const later = 'SPJS Itest Sitefields Later'
      const template = () => ({
        SiteFields: [
          `<Field ID="{69b1c001-3ffe-414a-8774-507a9da1d89e}" Name="SPJSItestSfLookupLater" StaticName="SPJSItestSfLookupLater" DisplayName="SPJS Itest Lookup Later" Type="Lookup" List="{listid:${later}}" ShowField="Title" Group="SPJS Itest" />`
        ],
        Lists: [
          {
            Title: later,
            Description: 'Created in the same template as a lookup to it',
            Template: 100,
            ContentTypesEnabled: false
          }
        ]
      })
      const lookupList = () =>
        readSiteFields(conn, ['SPJSItestSfLookupLater'])
          .then((state) => state.fields[0].schema.attributes.List)
          .catch((error) => `not read back (${error.status})`)
      const run = await apply(conn, template())
      const listId = await conn.web.lists
        .getByTitle(later)
        .select('Id')<{ Id: string }>()
        .then((list) => list.Id.toLowerCase())
        .catch(() => undefined)
      const first = await lookupList()
      // What a second run makes of it, for the report: the list exists by then.
      const second = await apply(conn, template())
      t.diagnostic(
        `first run: ${run.error ? 'stopped' : 'completed'}, ${
          run.errors.length
        } errors, ${run.warnings.length} warnings, List=${first}; second run: ${
          second.error
            ? `stopped (${describeError(second.error)})`
            : 'completed'
        }, ${second.errors.length} errors, List=${await lookupList()}`
      )
      assertClean(run, 'the first run')
      assert.ok(listId, `the list ${later} was not created`)
      assert.equal(first, listId)
    }
  )
})
