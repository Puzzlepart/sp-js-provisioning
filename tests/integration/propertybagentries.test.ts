import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { PermissionKind } from '@pnp/sp/presets/all'
import { apply } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import { describeError, redact } from '../../debug/lib/redact'
import { decodePropertyKey } from '../../debug/lib/readback/site'
import { base64EncodeString } from '../../src/util'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, recording, skipReason, withScratchWeb } from './helpers'

const HANDLERS = ['PropertyBagEntries']
const CSOM = 'POST /_vti_bin/client.svc/ProcessQuery'
const INDEXED = 'vti_indexedpropertykeys'

/** Why one side of the suite skips: the reserved site allows custom script or blocks it. */
const BLOCKED =
  'custom script is blocked (DenyAddAndCustomizePages): SharePoint refuses every property bag write; an SPO admin must allow it'
const ALLOWED = 'custom script is allowed on this site'

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
 * The fixture's entries as the web has them, and whether its indexed keys hold the indexed ones.
 *
 * @param conn - The scratch web
 */
async function readBag(
  conn: ISiteConnection
): Promise<{ values: Record<string, any>; indexed: string[] }> {
  const raw = await conn.web.allProperties<Record<string, any>>()
  const bag: Record<string, any> = {}
  for (const key of Object.keys(raw)) bag[decodePropertyKey(key)] = raw[key]
  const keys = loadFixture('propertybagentries').PropertyBagEntries as Array<{
    Key: string
    Indexed?: boolean
  }>
  return {
    values: keys.reduce((all: Record<string, any>, entry) => {
      all[entry.Key] = bag[entry.Key] ?? null
      return all
    }, {}),
    indexed: keys
      .filter((entry) => entry.Indexed)
      .map((entry) => entry.Key)
      .filter((key) =>
        String(bag[INDEXED] || '')
          .split('|')
          .includes(base64EncodeString(key))
      )
  }
}

/**
 * The PropertyBagEntries handler on a scratch web (docs/specs/jsom-exit.md, J6). SharePoint refuses
 * a property bag write on a site that blocks custom script, the default on modern sites, whichever
 * API asks: there the handler must log one error that names the cause and let the template go on.
 * Where custom script is allowed, the entries and their indexed keys are written, PnP's way, and a
 * second apply writes nothing. The reserved site decides which of the two runs.
 */
describe('PropertyBagEntries', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('propertybagentries')

  it('logs one error naming custom script where it is blocked, writes nothing and does not stop the template', async (t) => {
    const { conn } = scratch()
    if (await customScriptAllowed(conn)) {
      t.skip(ALLOWED)
      return
    }
    const before = await readBag(conn)

    const run = await apply(conn, loadFixture('propertybagentries'), {
      handlers: HANDLERS
    })

    assert.equal(run.error, undefined, run.error && describeError(run.error))
    const errors = run.errors.map((entry) => redact(entry.message))
    assert.equal(errors.length, 1, errors.join('\n'))
    assert.match(
      errors[0],
      /Property bag entries SPJSItestPbNew, SPJSItestPbOverwrite, vti_indexedpropertykeys were not written: SharePoint refused the write, as it does on a site that blocks custom script \(DenyAddAndCustomizePages; a SharePoint admin can allow it\): .+ \(System\.UnauthorizedAccessException\)/
    )
    t.diagnostic(errors[0])
    assert.deepEqual(await readBag(conn), before)
  })

  it('writes the entries and their indexed keys where custom script is allowed, and a second apply writes nothing', async (t) => {
    const { conn } = scratch()
    if (!(await customScriptAllowed(conn))) {
      t.skip(BLOCKED)
      return
    }
    const requests: string[] = []
    const marks: number[] = []
    const { afterFirst } = await applyTwice(
      recording(conn, requests),
      () => loadFixture('propertybagentries'),
      async () => {
        marks.push(requests.length)
        return readBag(conn)
      },
      { handlers: HANDLERS }
    )

    assert.deepEqual(afterFirst, {
      values: { SPJSItestPbNew: 'ny verdi', SPJSItestPbOverwrite: 'fra malen' },
      indexed: ['SPJSItestPbNew']
    })
    assert.deepEqual(
      requests
        .slice(marks[0], marks[1])
        .filter((request) => request.startsWith(CSOM)),
      [],
      'the second run wrote the property bag'
    )
  })
})
