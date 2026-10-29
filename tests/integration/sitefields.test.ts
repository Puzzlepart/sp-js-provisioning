import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { SPInstance } from '@pnp/sp/presets/all'
import { apply, IApplyResult } from '../../debug/lib/apply'
import { describeError, redact } from '../../debug/lib/redact'
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
 * The lines a run logged at a level, redacted (they reach JUnit).
 *
 * @param run - The run
 * @param level - `errors` or `warnings`
 */
const lines = (run: IApplyResult, level: 'errors' | 'warnings') =>
  run[level].map((entry) => redact(entry.message))

/**
 * The SiteFields handler on a scratch web: site columns of eleven types from field XML, `{sitename}`
 * in a description and `{listid:}` in a Lookup column's `List`, created on the first run and
 * updated in place with the list-field rules on later runs (docs/specs/sitefields-re-apply.md): an
 * unchanged column is not written (its `SchemaXml` version stays), a change is merged onto what the
 * column holds.
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

  it('creates eleven column types from field XML as declared, and an unchanged second apply writes nothing', async () => {
    const { conn } = scratch()
    const names = namesOf(loadFixture('sitefields'))
    const versions: Array<Record<string, number>> = []
    const { afterFirst, second } = await applyTwice(
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
    // The second run found every column (same ids: applyTwice compared the read-backs) and wrote
    // none: each SchemaXml MERGE raises a column's version by 2 (probe SF1). It used to send the
    // template's XML for every column on every run, dropping what the template left out.
    assert.deepEqual(versions[1], versions[0], 'the second run wrote columns')
    for (const name of names)
      assert.ok(
        second.entries.some(
          (entry) =>
            entry.message.includes(`Site column ${name} (`) &&
            entry.message.endsWith('is unchanged.')
        ),
        `${name} is not logged as unchanged`
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
    // versions: [A, B, B again]; only the Text column was written, once.
    const others = (version: Record<string, number>) =>
      names.filter((name) => name !== text).map((name) => version[name])
    assert.ok(versions[1][text] > versions[0][text], 'B wrote the Text column')
    assert.deepEqual(others(versions[1]), others(versions[0]), 'B wrote others')
    assert.deepEqual(versions[2], versions[1], 'B again wrote columns')

    assertClean(
      await apply(conn, loadFixture('sitefields')),
      'the A run after B'
    )
    assert.deepEqual(await read(conn), a, 'A again restores the Text column')
  })

  it('keeps what an admin set on a column whose template leaves it out', async (t) => {
    // SharePoint replaces a site column's schema wholesale (probe SF2): the engine used to send the
    // template's XML, so a re-import dropped a description, an index or a formatter set since.
    const { conn } = scratch()
    const note = 'SPJSItestSfNote'
    const versions: Array<Record<string, number>> = []
    const read = reader([note], versions)
    assertClean(await apply(conn, loadFixture('sitefields')), 'the A run')
    await conn.web.fields
      .getByInternalNameOrTitle(note)
      .update({ Description: 'Satt av en administrator' }, 'SP.Field')
    const before = await read(conn)
    t.diagnostic(`set by hand: Description, version ${versions[0][note]}`)

    const run = await apply(conn, loadFixture('sitefields'))

    assertClean(run, 'the run after the change by hand')
    const after = await read(conn)
    assert.equal(byName(after, note).Description, 'Satt av en administrator')
    assert.deepEqual(after, before)
    assert.equal(versions[1][note], versions[0][note], 'the column was written')
  })

  it('leaves a column as it is, with one warning, when the template narrows its type', async () => {
    const { conn } = scratch()
    const note = 'SPJSItestSfNote'
    const versions: Array<Record<string, number>> = []
    const read = reader([note], versions)
    assertClean(await apply(conn, loadFixture('sitefields')), 'the A run')
    const before = await read(conn)
    const narrowed = (loadFixture('sitefields').SiteFields as string[])
      .filter((xml) => fieldSchema(xml).attributes.Name === note)
      .map((xml) => xml.replace('Type="Note"', 'Type="Text"'))

    const run = await apply(conn, { SiteFields: narrowed })

    assertClean(run, 'the narrowing run')
    assert.equal(run.warnings.length, 1, lines(run, 'warnings').join('\n'))
    assert.match(
      lines(run, 'warnings')[0],
      /SPJSItestSfNote \(SPJS Itest Note\) is Note, the template says Text/
    )
    assert.deepEqual(await read(conn), before)
    assert.equal(byName(before, note).TypeAsString, 'Note')
    assert.equal(versions[1][note], versions[0][note], 'the column was written')
  })

  it('resolves {listid:} in a lookup to a list from the same template on the first run', async (t) => {
    // The list is created before SiteFields (WebProvisioner), so the column, and the list copy a
    // FieldRef makes of it in the same run, point at it. It used to store the literal token,
    // silently, and only a second run pointed the site column at the list.
    const { conn } = scratch()
    const later = 'SPJS Itest Sitefields Later'
    const copies = 'SPJS Itest Sitefields Copies'
    const lookupName = 'SPJSItestSfLookupLater'
    const template = () => ({
      SiteFields: [
        `<Field ID="{69b1c001-3ffe-414a-8774-507a9da1d89e}" Name="${lookupName}" StaticName="${lookupName}" DisplayName="SPJS Itest Lookup Later" Type="Lookup" List="{listid:${later}}" ShowField="Title" Group="SPJS Itest" />`
      ],
      Lists: [
        {
          Title: later,
          Description: 'Created in the same template as a lookup to it',
          Template: 100,
          ContentTypesEnabled: false
        },
        {
          Title: copies,
          Description: 'Holds a copy of the lookup column, through a FieldRef',
          Template: 100,
          ContentTypesEnabled: false,
          FieldRefs: [
            {
              ID: '69b1c001-3ffe-414a-8774-507a9da1d89e',
              Name: lookupName,
              DisplayName: 'SPJS Itest Kilde'
            }
          ]
        }
      ]
    })
    const versions: Array<Record<string, number>> = []
    const lookupList = async () => {
      const { versions: version, fields } = await readSiteFields(conn, [
        lookupName
      ])
      versions.push(version)
      return fields[0].schema.attributes.List
    }
    const copyList = () =>
      conn.web.lists
        .getByTitle(copies)
        .fields.getByInternalNameOrTitle(lookupName)
        .select('SchemaXml')<{ SchemaXml: string }>()
        .then((field) => fieldSchema(field.SchemaXml).attributes.List)

    const run = await apply(conn, template())
    assertClean(run, 'the first run')
    const listId = await conn.web.lists
      .getByTitle(later)
      .select('Id', 'Description')<{ Id: string; Description: string }>()
    assert.equal(await lookupList(), listId.Id.toLowerCase())
    assert.equal(await copyList(), listId.Id.toLowerCase())
    assert.equal(
      listId.Description,
      'Created in the same template as a lookup to it'
    )
    t.diagnostic(
      redact(
        run.entries
          .filter((entry) =>
            entry.message.includes('created before SiteFields')
          )
          .map((entry) => entry.message)
          .join(' | ')
      )
    )

    const second = await apply(conn, template())
    assertClean(second, 'the second run')
    assert.equal(await lookupList(), listId.Id.toLowerCase())
    assert.equal(versions[1][lookupName], versions[0][lookupName])
  })

  it('leaves an unchanged taxonomy column alone on a second apply', async (t) => {
    // SharePoint ignores a taxonomy column's Customization in a SchemaXml MERGE, stores its own
    // hidden list, web and ShowField, and adds a TextField (probe SF3): merged as the template has
    // them, every run would write the column and change nothing.
    const { conn } = scratch()
    try {
      await SPInstance([conn.web, conn.url], '_api/v2.1/termstore').select(
        'id'
      )<{ id: string }>()
    } catch (error) {
      if (conn.mode === 'certificate' && error && error.status === 403) {
        t.skip('SharePoint refuses the app (403): it needs TermStore.Read.All')
        return
      }
      throw error
    }
    const name = 'SPJSItestTaxonomy'
    const value = (q: string, type: string, text: string) =>
      `<Value p4:type="${q}:${type}" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance" xmlns:${q}="http://www.w3.org/2001/XMLSchema">${text}</Value>`
    // As the vei hub template writes it: no TextField, the term store as a token. The term set need
    // not exist; SharePoint stores any GUID.
    const template = () => ({
      SiteFields: [
        `<Field DisplayName="SPJS Itest Taxonomy" Description="" Type="TaxonomyFieldTypeMulti" Mult="TRUE" ShowField="Term1033" ID="{2f0c8a4e-6b1d-4c39-9e57-a3d81b6f0c24}" StaticName="${name}" Name="${name}" Group="SPJS Itest"><Customization><ArrayOfProperty><Property><Name>SspId</Name>${value(
          'q1',
          'string',
          '{sitecollectiontermstoreid}'
        )}</Property><Property><Name>TermSetId</Name>${value(
          'q2',
          'string',
          '{7d4b2e91-0c5a-4f86-b3e2-91a6c0d4e857}'
        )}</Property></ArrayOfProperty></Customization></Field>`
      ]
    })
    const read = async () => {
      const field = await conn.web.fields
        .getByInternalNameOrTitle(name)
        .select('Id', 'TypeAsString', 'SchemaXml')<{
        Id: string
        TypeAsString: string
        SchemaXml: string
      }>()
      const property = (key: string) =>
        (new RegExp(`<Name>${key}</Name><Value[^>]*>([^<]*)<`).exec(
          field.SchemaXml
        ) || [])[1]
      return {
        Id: field.Id,
        TypeAsString: field.TypeAsString,
        version: /\sVersion="(\d+)"/.exec(field.SchemaXml)[1],
        SspId: property('SspId'),
        TermSetId: property('TermSetId'),
        TextField: property('TextField')
      }
    }

    const first = await apply(conn, template())
    assertClean(first, 'the first run')
    const created = await read()
    const second = await apply(conn, template())
    assertClean(second, 'the second run')
    assert.deepEqual(
      second.warnings.length,
      0,
      lines(second, 'warnings').join('\n')
    )

    assert.deepEqual(await read(), created, 'the second run wrote the column')
    assert.equal(created.TypeAsString, 'TaxonomyFieldTypeMulti')
    assert.match(created.SspId, /^[\da-f-]{36}$/i)
    assert.notEqual(created.SspId, '00000000-0000-0000-0000-000000000000')
    assert.equal(created.TermSetId, '7d4b2e91-0c5a-4f86-b3e2-91a6c0d4e857')
    assert.ok(created.TextField, 'SharePoint added no TextField')
  })

  it('logs a column SharePoint refuses and still creates the next (failures.json)', async () => {
    const { conn } = scratch()

    const run = await apply(conn, loadFixture('failures'), {
      handlers: ['SiteFields']
    })

    assert.equal(run.error, undefined, run.error && describeError(run.error))
    assert.equal(run.errors.length, 1, lines(run, 'errors').join('\n'))
    assert.match(
      lines(run, 'errors')[0],
      /Failed to create site column SPJSItestFailFormula \(SPJS Itest Failing Formula\)/
    )
    const valid = await readSiteFields(conn, ['SPJSItestFailValid'])
    assert.equal(valid.fields[0].TypeAsString, 'Text')
  })
})
