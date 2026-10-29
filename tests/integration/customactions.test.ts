import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { PermissionKind } from '@pnp/sp/presets/all'
import { apply } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import {
  ICustomActionState,
  readCustomActions
} from '../../debug/lib/readback/customactions'
import { ICustomAction } from '../../src/schema'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

/** Title prefix of the SPFx extension actions (`customactions.json` and its B variant). */
const SPFX = 'SPJS Itest SPFx'

/** Title prefix of the classic actions (`customactions-script.json`), which NoScript refuses. */
const CLASSIC = 'SPJS Itest Classic'

/** Why the classic-actions test skips on a site with custom script blocked (NoScript). */
const NOSCRIPT =
  'custom script is blocked (DenyAddAndCustomizePages); an SPO admin must allow it'

/**
 * The custom actions a fixture declares, sorted by title like the read-back.
 *
 * @param name - The fixture
 */
function declaredIn(name: string): ICustomAction[] {
  return (loadFixture(name).CustomActions as ICustomAction[])
    .slice()
    .sort((a, b) => a.Title.localeCompare(b.Title, 'en'))
}

/**
 * The read-back reduced to the members each fixture action declares, in the fixture's order, so
 * it compares with `deepEqual` against the fixture. Component ids compare in lower case (the
 * reader lower-cases them).
 *
 * @param states - The read-back
 * @param declared - The fixture's actions, sorted by title
 */
function asDeclared(
  states: ICustomActionState[],
  declared: ICustomAction[]
): { actual: Record<string, any>[]; expected: Record<string, any>[] } {
  const expected = declared.map((action) => {
    const copy: Record<string, any> = { ...action }
    if (copy.ClientSideComponentId)
      copy.ClientSideComponentId = copy.ClientSideComponentId.toLowerCase()
    return copy
  })
  const actual = expected.map((action) => {
    const state = states.find((candidate) => candidate.Title === action.Title)
    if (!state) return { Title: action.Title, missing: true }
    const picked: Record<string, any> = {}
    for (const key of Object.keys(action)) picked[key] = (state as any)[key]
    return picked
  })
  return { actual, expected }
}

/**
 * Whether the web lets the signed-in account add script: NoScript (`DenyAddAndCustomizePages`)
 * removes `AddAndCustomizePages` even for site collection administrators, and subwebs inherit it.
 *
 * @param conn - The web
 */
function customScriptAllowed(conn: ISiteConnection): Promise<boolean> {
  return conn.web.currentUserHasPermissions(PermissionKind.AddAndCustomizePages)
}

/**
 * The CustomActions handler on a scratch web: web-scoped user custom actions, matched by `Title`
 * and only ever added (`docs/ARCHITECTURE.md`, CustomActions). No tokens are replaced in an
 * action, so the fixtures hold literal values: SPFx component ids SharePoint stores without
 * looking them up, and SharePoint's own `~site` and `{ListId}` in the EditControlBlock `Url`.
 *
 * NoScript (`DenyAddAndCustomizePages`, the default on modern sites, inherited by subwebs) lets
 * only SPFx extensions (`ClientSideExtension.*`) through: SharePoint refuses every classic action
 * with 403 `System.UnauthorizedAccessException`, ScriptLink and also an `EditControlBlock` or
 * `Microsoft.SharePoint.StandardMenu` entry with or without a `Url` (probed 2026-10-21). So
 * `customactions.json` holds the SPFx actions and `customactions-script.json` the classic ones,
 * and only one of the last two tests runs on a given site.
 */
describe('CustomActions', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('customactions')

  it('adds SPFx extension actions (application customizer, command bar, context menu) as declared, and adds nothing on a second apply', async () => {
    const { conn } = scratch()
    const declared = declaredIn('customactions')
    const { afterFirst } = await applyTwice(
      conn,
      () => loadFixture('customactions'),
      (c) => readCustomActions(c, SPFX)
    )
    // One action per declared title: nothing missing, nothing added twice. `applyTwice` already
    // proved the second read-back equal to the first, ids included: no duplicate and no
    // delete-and-add on the second run.
    assert.deepEqual(
      afterFirst.map((action) => action.Title),
      declared.map((action) => action.Title)
    )
    const { actual, expected } = asDeclared(afterFirst, declared)
    assert.deepEqual(actual, expected)
    // What the fixture leaves out stays empty: an application customizer is registered on
    // nothing and has no Url.
    const customizer = afterFirst.find(
      (action) => action.Title === `${SPFX} Application Customizer`
    )
    assert.equal(customizer.Url, '')
    assert.equal(customizer.RegistrationId, '')
    assert.equal(customizer.RegistrationType, 0)
  })

  it('adds an action that is missing but does not update one that exists: a changed Sequence is not applied (add-only)', async () => {
    // Current, documented behaviour (ARCHITECTURE.md: "Existing actions are not updated"): the
    // handler matches by Title and skips a match, so the B variant's new Sequence values never
    // reach SharePoint. Asserted as is; a change to update-in-place must change this test.
    const { conn } = scratch()
    assertClean(await apply(conn, loadFixture('customactions')), 'the A run')
    const before = await readCustomActions(conn, SPFX)
    assertClean(await apply(conn, loadFixture('customactions.b')), 'the B run')
    const after = await readCustomActions(conn, SPFX)
    const b = declaredIn('customactions.b')

    // The action only B declares is added, as declared.
    const added = b.filter(
      (action) => !before.some((state) => state.Title === action.Title)
    )
    assert.deepEqual(
      added.map((action) => action.Title),
      [`${SPFX} Added Later`]
    )
    const { actual, expected } = asDeclared(after, added)
    assert.deepEqual(actual, expected)

    // The existing actions are untouched: same ids, the A Sequence, not B's.
    assert.deepEqual(
      after.filter((state) => before.some((old) => old.Id === state.Id)),
      before
    )
    for (const state of before) {
      const changed = b.find((action) => action.Title === state.Title)
      assert.notEqual(
        changed.Sequence,
        state.Sequence,
        `the B variant changes the Sequence of ${state.Title}`
      )
    }
    assert.equal(after.length, before.length + 1)
  })

  it('adds classic actions (a ScriptLink with a ScriptBlock, an EditControlBlock with a Url), and adds nothing on a second apply', async (t) => {
    const { conn } = scratch()
    if (!(await customScriptAllowed(conn))) {
      t.skip(NOSCRIPT)
      return
    }
    const declared = declaredIn('customactions-script')
    const { afterFirst } = await applyTwice(
      conn,
      () => loadFixture('customactions-script'),
      (c) => readCustomActions(c, CLASSIC)
    )
    assert.deepEqual(
      afterFirst.map((action) => action.Title),
      declared.map((action) => action.Title)
    )
    const { actual, expected } = asDeclared(afterFirst, declared)
    assert.deepEqual(actual, expected)
  })

  it('stops in CustomActions with the 403 when custom script is blocked and the template has classic actions', async (t) => {
    // The other side of the previous test: under NoScript the handler does not swallow the
    // refusal, the run ends with a ProvisioningError naming the handler, and nothing classic is
    // stored. (The message is SharePoint's and does not name the refused action.)
    const { conn } = scratch()
    if (await customScriptAllowed(conn)) {
      t.skip(
        'custom script is allowed: the previous test adds the classic actions'
      )
      return
    }
    for (const run of ['the first run', 'the second run']) {
      const result = await apply(conn, loadFixture('customactions-script'))
      assert.equal(result.order.includes('CustomActions'), true, run)
      assert.ok(result.error, `${run} ends with a ProvisioningError`)
      assert.equal(result.error.handler, 'CustomActions', run)
      assert.equal(result.error.cause && result.error.cause.status, 403, run)
      assert.match(result.error.message, /UnauthorizedAccessException/, run)
      assert.deepEqual(await readCustomActions(conn, CLASSIC), [], run)
    }
  })
})
