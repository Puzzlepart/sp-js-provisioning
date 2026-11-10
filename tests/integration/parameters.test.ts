import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { ISiteConnection } from '../../debug/lib/connect'
import { readSite } from '../../debug/lib/readback/site'
import { fieldSchema } from '../../debug/lib/readback/sitefields'
import { verify } from '../../debug/lib/verify'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, skipReason, withScratchWeb } from './helpers'

const LIST = 'SPJS Itest Parameters'

/**
 * The values for the fixture's `{parameter:}` tokens. The label has an apostrophe: the DataRows
 * key holds it, and the handler must match the row by its value, escaped for OData, or it adds the
 * row again on every run.
 */
const PARAMETERS = { label: "Kari's plan", default: 'standardverdi' }

/**
 * What the fixture provisions, as plain data.
 *
 * @param conn - The scratch web
 */
async function read(conn: ISiteConnection) {
  const state = await readSite(conn, {
    siteFields: ['SPJSItestPmText'],
    lists: [
      {
        title: LIST,
        rowColumns: ['Title', 'SPJSItestPmListText', 'SPJSItestPmText']
      }
    ]
  })
  const column = state.siteFields.SPJSItestPmText
  const list = state.lists[LIST]
  return {
    siteField: {
      Title: column && column.Title,
      Default: column && fieldSchema(column.SchemaXml).Default
    },
    listField: (
      list.fields.find(
        (field) => field.InternalName === 'SPJSItestPmListText'
      ) || { Title: undefined }
    ).Title,
    rows: list.items.items.map((item) => item.values)
  }
}

/**
 * `{parameter:}` tokens (`IProvisioningConfig.parameters`, `--parameters` in the runner) where the
 * engine replaces them: site column XML, list field XML and DataRows values, the `KeyColumn` value
 * included (docs/specs/integration-testing.md, Fixtures).
 */
describe('Parameters', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('parameters')

  it('replaces {parameter:} in field XML and data rows, matches the row by its value, and verify agrees', async () => {
    const { conn } = scratch()
    const { first, afterFirst } = await applyTwice(
      conn,
      () => loadFixture('parameters'),
      read,
      { parameters: PARAMETERS }
    )
    assert.deepEqual(first.order, ['SiteFields', 'Lists'])
    assert.deepEqual(afterFirst, {
      siteField: { Title: "SPJS Itest Kari's plan", Default: 'standardverdi' },
      listField: "SPJS Itest list Kari's plan",
      // One row after two runs: the second matched it by the key's value.
      rows: [
        {
          Title: "SPJS Itest Kari's plan",
          SPJSItestPmListText: 'standardverdi',
          SPJSItestPmText: "Kari's plan in the site column"
        }
      ]
    })
  })

  it('without the values, verify says which tokens it could not resolve and what differs', async () => {
    const { conn } = scratch()
    const result = await verify(conn, loadFixture('parameters'), {
      handlers: ['SiteFields', 'Lists']
    })
    assert.ok(
      result.notes.some((note) => /\{parameter:label\}/.test(note.reason)),
      'a note names {parameter:label}'
    )
    assert.ok(
      result.findings.some(
        (finding) =>
          finding.handler === 'SiteFields' &&
          finding.path === 'SPJSItestPmText/DisplayName'
      ),
      'the site column title differs from the unresolved one'
    )
  })
})
