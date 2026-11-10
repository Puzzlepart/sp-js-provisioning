import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { ISiteConnection } from '../../debug/lib/connect'
import { fixtureUser } from '../../debug/lib/placeholders'
import { redact } from '../../debug/lib/redact'
import { verify } from '../../debug/lib/verify'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, skipReason, withScratchWeb } from './helpers'

/** The sections of `all.node.json`, in run order. */
const SECTIONS = [
  'SiteFields',
  'Features',
  'Security',
  'Lists',
  'CustomActions',
  'ComposedLook',
  'ClientSidePages',
  'Navigation',
  'WebSettings'
]

/**
 * Every Node-capable handler fixture in one template (`all.node.json`, made by
 * `npm run build:fixtures`) on one fresh scratch web: the handlers see each other's results, as
 * in a real template (site columns the lists use, groups the list permissions name, lists the
 * lookups point at). Runs with `npm run test:integration -- all` or `-- everything`, not by
 * default: the handler suites cover each part.
 */
describe('Everything', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('everything')
  let user: string | undefined

  before(async () => {
    user = await fixtureUser(scratch().conn)
  })

  it('applies every Node-capable fixture together, twice, and verify finds nothing, exact on the fresh web', async (t) => {
    if (!user) {
      t.skip(
        'app-only without INTEGRATION_USER: no user for the members and User values'
      )
      return
    }
    const { conn } = scratch()
    const template = () => loadFixture('all.node', { user })
    /** What `verify` finds comparing everything: a fresh web the template was just applied to. */
    const exact = async (c: ISiteConnection) =>
      (await verify(c, template(), { exact: true })).findings.map((finding) =>
        redact(
          `${finding.handler} ${finding.path}: expected ${JSON.stringify(
            finding.expected
          )}, found ${JSON.stringify(finding.actual)}`
        )
      )
    const { first, afterFirst } = await applyTwice(conn, template, exact)
    assert.deepEqual(first.order, SECTIONS)
    assert.deepEqual(afterFirst, [], 'verify --exact found differences')
  })
})
