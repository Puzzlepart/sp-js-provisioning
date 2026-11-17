import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import { fieldTarget, guidOf } from '../../src/util/fieldxml'
import { loadFixture } from '../fixtures/fixtures'
import {
  applyTwice,
  assertClean,
  skipReason,
  termStoreRefusal,
  withScratchWeb
} from './helpers'

/** The handlers the fixture needs; Taxonomy is applied only when named. */
const HANDLERS = ['Taxonomy', 'SiteFields', 'ContentTypes', 'Lists']
/** The fixture's list, its phase column and the column's term sets. */
const LIST = 'SPJS Itest Fasesjekkliste'
const COLUMN = 'SPJSItestPhase'
/** `taxonomy.json`'s «SPJS Itest Faser»: the site column's term set. */
const SITE_SET = '518a580a-11e9-4f6d-b4b5-3099c454535f'
/** `lists-taxonomy.json`'s «SPJS Itest Faser Vei»: the list field's term set. */
const LIST_SET = 'eeb03669-70cc-417c-8c25-5b9a8f6889d3'

/**
 * The list field's term set and the term each item holds in it, by title, in item order.
 *
 * @param conn - The scratch web
 */
async function readPhases(
  conn: ISiteConnection
): Promise<{ termSet: string; terms: Array<[string, string | null]> }> {
  const { SchemaXml } = await conn.web.lists
    .getByTitle(LIST)
    .fields.getByInternalNameOrTitle(COLUMN)
    .select('SchemaXml')<{ SchemaXml: string }>()
  const items = await conn.web.lists
    .getByTitle(LIST)
    .items.select('Title', COLUMN)
    .orderBy('Id')<Array<Record<string, any>>>()
  return {
    termSet: guidOf(fieldTarget(SchemaXml, 'TermSetId')),
    terms: items.map((item) => [
      item.Title,
      item[COLUMN] ? guidOf(item[COLUMN].TermGuid) : null
    ])
  }
}

/**
 * A list taxonomy field whose content type brings its site column with another term set than the
 * list declares (PP365 stress test row 5: the vei hub's Fasesjekkliste kept the hub's «Fase», and
 * every checklist row's phase was dropped without a word). Two term sets in the `SPJS Itest` term
 * group, created once and left in place like the rest of the group (the Taxonomy suite's
 * `taxonomy.json` and this fixture's own).
 */
describe('Lists (taxonomy)', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('lists-taxonomy')

  it("points the list field at the template's term set, not its site column's, and stores the rows' terms", async (t) => {
    const { conn } = scratch()
    // The term sets: the site column's is the Taxonomy suite's, the list field's this fixture's.
    // A term store that has them is only read; in certificate mode one that refuses the app a
    // write skips the test with SharePoint's words, as the Taxonomy suite does.
    for (const name of ['taxonomy', 'lists-taxonomy']) {
      const run = await apply(conn, loadFixture(name), {
        handlers: ['Taxonomy']
      })
      const refused = termStoreRefusal(run)
      if (refused && conn.mode === 'certificate') {
        t.skip(
          `SharePoint refuses the app writes to the term store: ${refused}`
        )
        return
      }
      assertClean(run, `the ${name}.json Taxonomy run`)
    }
    const rows = loadFixture('lists-taxonomy').Lists[0].DataRows.Rows

    const { afterFirst } = await applyTwice(
      conn,
      () => loadFixture('lists-taxonomy'),
      readPhases,
      { handlers: HANDLERS }
    )

    assert.equal(
      guidOf(
        fieldTarget(loadFixture('lists-taxonomy').SiteFields[0], 'TermSetId')
      ),
      SITE_SET,
      'the site column names another term set than the list'
    )
    assert.equal(afterFirst.termSet, LIST_SET)
    assert.deepEqual(
      afterFirst.terms,
      rows.map((row: any) => [row.Title, row[COLUMN].termId])
    )
  })
})
