import assert from 'node:assert/strict'
import { after, before } from 'node:test'
import { apply, IApplyOptions, IApplyResult } from '../../debug/lib/apply'
import { accessDenied, connect, ISiteConnection } from '../../debug/lib/connect'
import { authMode, env } from '../../debug/lib/env'
import { describeError, redact } from '../../debug/lib/redact'
import {
  createScratchWeb,
  deleteScratchWeb,
  IScratchWeb,
  removeScratchWeb,
  scratchName,
  stampOf
} from '../../debug/lib/scratch'

export { skipReason } from '../../debug/lib/env'

/**
 * Connects to the reserved site; a refused sign-in fails with what to do about it (renew the
 * session, give the account access) instead of PnPjs's raw 403. Redacted: it reaches JUnit.
 */
export async function connectToSite(): Promise<ISiteConnection> {
  try {
    return await connect(env.siteUrl)
  } catch (error) {
    throw new Error(
      redact(
        accessDenied(error, env.siteUrl, authMode()) ?? describeError(error)
      )
    )
  }
}

/**
 * A scratch web for the suite's tests: created in `before`, deleted in `after` with the site
 * groups whose title starts with its title (unless `INTEGRATION_KEEP=true`). When
 * `tests/integration/run.ts` shares one web per run (`INTEGRATION_SCRATCH=run`), the suite uses
 * that one and leaves it to the runner. Call it inside a `describe`; it returns a getter, since the
 * web exists only once `before` has run.
 *
 * A `before` that times out leaves the creation running (a promise cannot be stopped): `after`
 * waits for it and deletes what it made, by name when the creation failed half-way and could not
 * clean up after itself. A run once left a web behind that way, its requests hanging.
 *
 * @param suite - The suite, for the web's name (`SPJS Itest <suite> <stamp>`)
 */
export function withScratchWeb(suite: string): () => IScratch {
  let parent: ISiteConnection | undefined
  let scratch: IScratch | undefined
  let name: { title: string; url: string } | undefined
  let creating: Promise<IScratchWeb> | undefined
  before(
    async () => {
      parent = await connectToSite()
      const shared = process.env.INTEGRATION_SCRATCH_WEB
      if (shared) {
        scratch = {
          title: process.env.INTEGRATION_SCRATCH_TITLE || '',
          url: shared,
          stamp: process.env.INTEGRATION_SCRATCH_STAMP || '',
          conn: await connect(`${parent.url}/${shared}`, parent.mode),
          parent,
          shared: true
        }
        return
      }
      const now = new Date()
      name = scratchName(suite, stampOf(now))
      creating = createScratchWeb(parent, suite, now)
      scratch = { ...(await creating), parent, shared: false }
    },
    // A few requests, each limited to REQUEST_TIMEOUT per attempt: room for slow ones that
    // succeed. after() deals with a creation still running when this gives up.
    { timeout: 600_000 }
  )
  after(
    async () => {
      if (scratch && scratch.shared) return
      if (!scratch && creating && parent) {
        const created = await creating.catch(() => undefined)
        if (created) scratch = { ...created, parent, shared: false }
      }
      if (scratch && env.keep) {
        // eslint-disable-next-line no-console
        console.log(
          `# kept scratch web ${scratch.conn.webInfo.ServerRelativeUrl}`
        )
        return
      }
      try {
        if (scratch) await deleteScratchWeb(parent, scratch)
        else if (parent && name) await removeScratchWeb(parent, name)
      } catch (error) {
        // node:test drops an after() error when before() failed: say what is left.
        // eslint-disable-next-line no-console
        console.error(
          redact(
            `# scratch web ${
              scratch ? scratch.url : name && name.url
            } may be left behind: ${describeError(
              error
            )}; remove it with npm run cleanup -- --webs`
          )
        )
        throw error
      }
    },
    // Long enough to outlast a creation still running after before() gave up, and its own
    // clean-up, then delete.
    { timeout: 900_000 }
  )
  return () => {
    if (!scratch) throw new Error('The scratch web is created in before()')
    return scratch
  }
}

/**
 * A scratch web with the reserved site it lives in.
 */
export interface IScratch extends IScratchWeb {
  parent: ISiteConnection
  shared: boolean
}

/**
 * Asserts that a run neither stopped nor logged an error.
 *
 * @param result - The run
 * @param label - Which run, for the message
 */
export function assertClean(result: IApplyResult, label = 'the run'): void {
  // Only strings reach the assertion: a failing `assert.equal(result.error, undefined)` would
  // print the whole error, whose `cause.response` holds the response headers (`Set-Cookie`).
  if (result.error)
    assert.fail(
      `${label} stopped in ${result.error.handler}: ${describeError(
        result.error
      )}`
    )
  assert.deepEqual(
    result.errors.map((entry) => redact(entry.message)),
    [],
    `${label} logged errors`
  )
}

/**
 * The two runs and read-backs of {@link applyTwice}.
 */
export interface ITwice<T> {
  first: IApplyResult
  second: IApplyResult
  afterFirst: T
  afterSecond: T
}

/**
 * Applies a template, reads the result back, applies it again and reads back again. Each run must
 * end without a `ProvisioningError` and without a logged error, and the second read-back must equal
 * the first: re-applying a template changes nothing. Each run gets a fresh template from
 * `template()`, since handlers mutate their input.
 *
 * @param conn - The scratch web
 * @param template - Returns a fresh copy of the template
 * @param read - Reads back what the template provisions, as plain, comparable data
 * @param options - Apply options (handlers, parameters)
 */
export async function applyTwice<T>(
  conn: ISiteConnection,
  template: () => Record<string, any>,
  read: (conn: ISiteConnection) => Promise<T>,
  options: IApplyOptions = {}
): Promise<ITwice<T>> {
  const first = await apply(conn, template(), options)
  assertClean(first, 'the first run')
  const afterFirst = await read(conn)
  const second = await apply(conn, template(), options)
  assertClean(second, 'the second run')
  const afterSecond = await read(conn)
  assert.deepEqual(
    afterSecond,
    afterFirst,
    'the second run changed what the first provisioned'
  )
  return { first, second, afterFirst, afterSecond }
}
