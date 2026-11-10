import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { apply, IApplyResult } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import { describeError, redact } from '../../debug/lib/redact'
import {
  bareGuid,
  ITaxonomyState,
  readTaxonomy
} from '../../debug/lib/readback/site'
import { loadFixture } from '../fixtures/fixtures'
import {
  applyTwice,
  assertClean,
  connectToSite,
  recording,
  skipReason
} from './helpers'

const HANDLERS = ['Taxonomy']
const CSOM = 'POST /_vti_bin/client.svc/ProcessQuery'

/**
 * Why a run could not write to the term store, when SharePoint refused it for lack of permission.
 *
 * @param run - The run
 */
function refusal(run: IApplyResult): string | undefined {
  const message = run.error ? redact(describeError(run.error)) : ''
  return /UnauthorizedAccessException|E_ACCESSDENIED|access denied|ingen tilgang/i.test(
    message
  )
    ? message
    : undefined
}

/**
 * The Taxonomy handler against the tenant's default term store (docs/specs/jsom-exit.md, J5). The
 * term store is shared by every site of the tenant, so the suite writes to one term group of its
 * own, `SPJS Itest` (`taxonomy.json`, fixed ids), created by its first run and left in place:
 * nothing is ever deleted from the term store. Later runs find it, so they prove the re-apply (one
 * read, no write), the custom sort order changed and changed back (B), and that `UpdateExistingTerms`
 * never overwrites a custom property; creating the group, the term set and the terms was proven by
 * the first run (2026-11-10) and is pinned by the unit tests. No scratch web: the handler writes
 * nothing to a web.
 */
describe('Taxonomy', { skip: skipReason() }, () => {
  let conn: ISiteConnection

  before(async () => {
    conn = await connectToSite()
  })

  /**
   * Reads back the fixture's group, term set and terms through REST (`_api/v2.1/termstore`), not
   * through the engine's CSOM requests.
   */
  async function read(): Promise<ITaxonomyState> {
    const { TermGroup, TermSets } = loadFixture('taxonomy').Taxonomy
    return readTaxonomy(conn, {
      groupId: TermGroup.Id,
      sets: TermSets.map((set: any) => ({
        id: set.Id,
        terms: set.Terms.map((term: any) => term.Id)
      }))
    })
  }

  /**
   * Applies a fixture once; in certificate mode a term store that refuses the app skips the test
   * with SharePoint's words (the app may lack the right to write to the term store).
   *
   * @param t - The test
   * @param name - The fixture
   */
  async function applyOnce(t: any, name: string): Promise<boolean> {
    const run = await apply(conn, loadFixture(name), { handlers: HANDLERS })
    const refused = refusal(run)
    if (refused && conn.mode === 'certificate') {
      t.skip(`SharePoint refuses the app writes to the term store: ${refused}`)
      return false
    }
    assertClean(run, `the ${name} run`)
    return true
  }

  it("has the fixture's group, term set and terms, with its ids, custom properties and sort order, and a second apply only reads", async (t) => {
    if (!(await applyOnce(t, 'taxonomy'))) return
    const requests: string[] = []
    const marks: number[] = []
    const { afterFirst, second } = await applyTwice(
      recording(conn, requests),
      () => loadFixture('taxonomy'),
      async () => {
        marks.push(requests.length)
        return read()
      },
      { handlers: HANDLERS }
    )
    if (afterFirst.unavailable) {
      t.skip(afterFirst.unavailable)
      return
    }
    const fixture = loadFixture('taxonomy').Taxonomy
    const set = fixture.TermSets[0]
    const stored = afterFirst.sets[set.Id]

    assert.deepEqual(afterFirst.group, {
      Id: bareGuid(fixture.TermGroup.Id),
      Name: fixture.TermGroup.Name
    })
    assert.deepEqual(
      [stored.Id, stored.names, stored.description, stored.isOpen],
      [bareGuid(set.Id), [set.Name], set.Description, false]
    )
    assert.deepEqual(
      set.Terms.map((term: any) => stored.terms[term.Id]),
      set.Terms.map((term: any) => ({
        Id: bareGuid(term.Id),
        labels: [term.Name],
        properties: term.CustomProperties || {}
      }))
    )
    assert.deepEqual(
      stored.customSortOrder,
      set.Terms.map((term: any) => bareGuid(term.Id))
    )

    // Each run read the store once, and the second wrote nothing.
    for (const [index, slice] of [
      requests.slice(0, marks[0]),
      requests.slice(marks[0], marks[1])
    ].entries())
      assert.deepEqual(
        slice.filter((request) => request.startsWith(CSOM)),
        [CSOM],
        `run ${index + 1} sent more than the read`
      )
    assert.ok(
      second.entries.some((entry) =>
        entry.message.endsWith(`Term set ${set.Name} is unchanged.`)
      ),
      'the term set is not logged as unchanged'
    )
  })

  it('applies a reversed sort order (B) without overwriting a custom property, and back', async (t) => {
    if (!(await applyOnce(t, 'taxonomy'))) return
    const a = await read()
    if (a.unavailable) {
      t.skip(a.unavailable)
      return
    }
    const set = loadFixture('taxonomy.b').Taxonomy.TermSets[0]
    const first = set.Terms[0]

    const requests: string[] = []
    const { afterFirst: b } = await applyTwice(
      recording(conn, requests),
      () => loadFixture('taxonomy.b'),
      read,
      { handlers: HANDLERS }
    )

    const ids = set.Terms.map((term: any) => bareGuid(term.Id))
    assert.deepEqual(a.sets[set.Id].customSortOrder, ids)
    assert.deepEqual(b.sets[set.Id].customSortOrder, ids.slice().reverse())
    // B gives the first term another value with UpdateExistingTerms: it keeps the one it has.
    assert.equal(first.CustomProperties.SPJSItestLevel, 'Endret i B')
    assert.equal(
      b.sets[set.Id].terms[first.Id].properties.SPJSItestLevel,
      'Prosjekt'
    )
    assert.deepEqual(b.sets[set.Id].terms, a.sets[set.Id].terms)
    // B wrote once (the order) and then only read: three CSOM requests in all.
    assert.equal(
      requests.filter((request) => request.startsWith(CSOM)).length,
      3,
      requests.join('\n')
    )

    assert.ok(await applyOnce(t, 'taxonomy'), 'the A run after B')
    assert.deepEqual(await read(), a, 'A again restores the sort order')
  })
})
