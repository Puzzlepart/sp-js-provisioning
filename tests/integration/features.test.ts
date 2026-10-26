import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import { readFeatures } from '../../debug/lib/readback/features'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

/**
 * The Features handler on a scratch web: `features.json` activates a web feature,
 * `features.b.json` deactivates it. The feature is Following Content (`FollowingContent`), a
 * web-scoped feature that needs no custom script and toggles cleanly on an STS#3 subweb; a fresh
 * scratch web has it active. Each test first puts the feature in the state it starts from, directly
 * through REST, so the tests do not depend on their order.
 */
describe('Features', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('features')
  const id: string = loadFixture('features').Features[0].id
  const deactivateId: string = loadFixture('features.b').Features[0].id

  /**
   * Activates or deactivates the feature outside the engine, when it is not in that state yet,
   * and returns the read-back the test starts from.
   *
   * @param conn - The scratch web
   * @param active - The state to start from
   */
  async function startWith(conn: ISiteConnection, active: boolean) {
    const now = await readFeatures(conn)
    if (now.active.includes(id) !== active) {
      if (active) await conn.web.features.add(id, false)
      else await conn.web.features.remove(id, false)
    }
    const before = await readFeatures(conn)
    assert.equal(before.active.includes(id), active)
    return before
  }

  /**
   * Runs a pinned test with a short failure. `assertClean` compares the whole `ProvisioningError`
   * with `undefined`, so its failure prints a diff of the HTTP response object; its first line is
   * the evidence, without the tenant's host (SharePoint names the web in its error text). Other
   * failures (a read-back diff) are rethrown as they are.
   *
   * @param run - The test body
   */
  async function brief(run: () => Promise<unknown>): Promise<void> {
    try {
      await run()
    } catch (error) {
      if (!error || !(error.actual instanceof Error)) throw error
      throw new Error(
        String(error.message)
          .split('\n')[0]
          .replace(/https:\/\/[^\s/"]+/g, 'https://<tenant>')
      )
    }
  }

  it('the fixtures toggle the same feature', () => {
    assert.equal(deactivateId, id)
    assert.equal(loadFixture('features').Features[0].deactivate, false)
    assert.equal(loadFixture('features.b').Features[0].deactivate, true)
  })

  it('activates an inactive web feature, and nothing else', async () => {
    const { conn } = scratch()
    const before = await startWith(conn, false)
    const result = await apply(conn, loadFixture('features'))
    assertClean(result)
    assert.deepEqual(result.order, ['Features'])
    assert.deepEqual(
      (await readFeatures(conn)).active,
      [...before.active, id].sort()
    )
  })

  it('deactivates an active web feature (B), and nothing else', async () => {
    const { conn } = scratch()
    const before = await startWith(conn, true)
    const result = await apply(conn, loadFixture('features.b'))
    assertClean(result)
    assert.deepEqual(result.order, ['Features'])
    assert.deepEqual(
      (await readFeatures(conn)).active,
      before.active.filter((active) => active !== id)
    )
  })

  it('changes nothing when the activation is applied a second time', () =>
    brief(async () => {
      const { conn } = scratch()
      await startWith(conn, false)
      const { afterFirst } = await applyTwice(
        conn,
        () => loadFixture('features'),
        readFeatures
      )
      assert.ok(afterFirst.active.includes(id))
    }))

  it('changes nothing when the deactivation (B) is applied a second time', () =>
    brief(async () => {
      const { conn } = scratch()
      await startWith(conn, true)
      const { afterFirst } = await applyTwice(
        conn,
        () => loadFixture('features.b'),
        readFeatures
      )
      assert.ok(!afterFirst.active.includes(id))
    }))

  it('activates again without error with force: true (re-activation)', async () => {
    const { conn } = scratch()
    const before = await startWith(conn, false)
    const forced = () => {
      const template = loadFixture('features')
      template.Features[0].force = true
      return template
    }
    const { afterFirst } = await applyTwice(conn, forced, readFeatures)
    assert.deepEqual(afterFirst.active, [...before.active, id].sort())
  })
})
