import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import * as path from 'path'
import { parseAge, prefixedArtefacts } from '../../debug/lib/cleanup'
import { UsageError } from '../../debug/lib/cli'
import { assertWritable } from '../../debug/lib/connect'
import { isWithinSite, readEnv, skipReason } from '../../debug/lib/env'
import {
  dateOfStamp,
  removeScratchWeb,
  scratchName,
  stampIn,
  stampOf
} from '../../debug/lib/scratch'
import { describeError, redact } from '../../debug/lib/redact'
import { loadFixture } from '../fixtures/fixtures'
import { redactXml, suiteFiles } from '../integration/run'

const site = 'https://contoso.sharepoint.com/sites/spjs-itest'

describe('harness: scratch web names', () => {
  it('stamps in UTC and reads the time back from a name', () => {
    const now = new Date(Date.UTC(2026, 9, 21, 15, 10, 15))
    const stamp = stampOf(now)

    assert.equal(stamp, '261021151015')
    assert.equal(dateOfStamp(stamp).getTime(), now.getTime())
    assert.deepEqual(scratchName('Lists Data', stamp), {
      title: 'SPJS Itest Lists Data 261021151015',
      url: 'spjs-itest-lists-data-261021151015'
    })
    assert.equal(stampIn('spjs-itest-lists-261021151015'), stamp)
    assert.equal(stampIn('SPJS Itest lists 261021151015-eiere'), stamp)
    assert.equal(stampIn('SPJS Itest Readers'), undefined)
    assert.equal(stampIn('spjs-itest-lists-999999999999'), undefined)
  })

  it('deletes by name only a scratch web, before any request', async () => {
    // No site: a refused name never reaches SharePoint.
    const parent: any = {}
    const title = 'SPJS Itest lists 261021151015'
    for (const url of [
      'customer-hub',
      'spjs-itest-lists',
      'spjs-itest-lists-261021151016'
    ])
      await assert.rejects(
        removeScratchWeb(parent, { title, url }),
        /Not a scratch web/,
        url
      )
    await assert.rejects(
      removeScratchWeb(parent, {
        title: 'Project Hub 261021151015',
        url: 'spjs-itest-lists-261021151015'
      }),
      /Not a scratch web title/
    )
  })
})

describe('harness: the site guard', () => {
  const configuration = readEnv({ INTEGRATION_SITE_URL: `${site}/` })

  it('allows the reserved site and webs below it, in any case', () => {
    for (const target of [
      site,
      `${site}/`,
      `${site}/spjs-itest-lists-1`,
      site.toUpperCase()
    ])
      assertWritable(target, undefined, configuration)
    assert.equal(isWithinSite(`${site}/sub/sub`, site), true)
  })

  it('refuses another site, also one whose URL starts the same', () => {
    for (const target of [
      'https://contoso.sharepoint.com/sites/other',
      `${site}2`,
      'https://fabrikam.sharepoint.com/sites/spjs-itest'
    ])
      assert.throws(
        () => assertWritable(target, undefined, configuration),
        (error: Error) =>
          error instanceof UsageError &&
          error.message.includes(`Pass --allow-site ${target}`)
      )
  })

  it('refuses a target that leaves the site through dot segments or encoding', () => {
    // `fetch` resolves `..`: the string started with the reserved site, the
    // requests went to another one.
    for (const target of [
      `${site}/../customer-hub`,
      `${site}/sub/../../customer-hub`,
      `${site}/%2e%2e/customer-hub`,
      `${site}/%2E%2E%2Fcustomer-hub`,
      `${site}\\..\\customer-hub`,
      'ftp://contoso.sharepoint.com/sites/spjs-itest'
    ])
      assert.throws(
        () => assertWritable(target, undefined, configuration),
        UsageError,
        target
      )
    // A `..` that stays inside the site is fine.
    assertWritable(`${site}/a/../b`, undefined, configuration)
  })

  it('writes outside it only when --allow-site repeats the target', () => {
    const other = 'https://contoso.sharepoint.com/sites/hub'
    assertWritable(other, `${other}/`, configuration)
    assert.throws(() => assertWritable(other, site, configuration), UsageError)
    // Without INTEGRATION_SITE_URL every target needs --allow-site.
    assert.throws(
      () => assertWritable(site, undefined, readEnv({})),
      UsageError
    )
  })
})

describe('harness: skipping', () => {
  it('names what is missing, and throws it under INTEGRATION_STRICT', () => {
    const exists = () => true
    assert.match(
      skipReason(readEnv({}), () => false),
      /no sign-in configured/
    )
    assert.match(
      skipReason(readEnv({}), exists),
      /INTEGRATION_SITE_URL not set/
    )
    assert.equal(
      skipReason(readEnv({ INTEGRATION_SITE_URL: site }), exists),
      undefined
    )
    assert.throws(
      () => skipReason(readEnv({ INTEGRATION_STRICT: 'true' }), exists),
      /INTEGRATION_STRICT is set and INTEGRATION_SITE_URL not set/
    )
  })
})

describe('harness: cleanup ages', () => {
  it('parses minutes, hours and days', () => {
    assert.equal(parseAge('30m'), 30 * 60_000)
    assert.equal(parseAge('2h'), 2 * 3_600_000)
    assert.equal(parseAge('1D'), 86_400_000)
    assert.throws(() => parseAge('2 weeks'), /Invalid age/)
  })
})

describe('harness: cleanup by prefix', () => {
  /** A PnPjs collection or object stand-in: `.select()`/`.filter()` chain, a call reads. */
  const fake = (data: any): any => {
    const read: any = () =>
      data instanceof Error ? Promise.reject(data) : Promise.resolve(data)
    read.select = () => read
    read.filter = () => read
    return read
  }

  /** A web whose associated groups have these ids (an Error: that read fails). */
  const associated = (owner: any, member: any, visitor: any) => ({
    associatedOwnerGroup: fake(owner),
    associatedMemberGroup: fake(member),
    associatedVisitorGroup: fake(visitor)
  })
  const notFound = () => Object.assign(new Error('Not found'), { status: 404 })
  const connection = (web: any, rootWeb: any) =>
    ({ web, sp: { site: { rootWeb } } } as any)
  const groups = [
    { Id: 3, Title: 'SPJS Itest Owners' },
    { Id: 4, Title: 'SPJS Itest Readers' },
    { Id: 5, Title: 'SPJS Itest lists 261021151015-eiere' },
    { Id: 6, Title: 'Prosjektstyre' },
    { Id: 8, Title: 'SPJS Itest Hub Members' }
  ]
  const webWith = (ids: any[]) => ({
    lists: fake([]),
    contentTypes: fake([]),
    fields: fake([]),
    userCustomActions: fake([]),
    navigation: { quicklaunch: fake([]), topNavigationBar: fake([]) },
    siteGroups: fake(groups),
    ...associated(ids[0], ids[1], ids[2])
  })

  it("never removes the web's or the root web's associated groups, and leaves stamped ones to --webs", async () => {
    // A reserved site titled `SPJS Itest` has `SPJS Itest Owners` as its owners group; a web
    // below a root titled `SPJS Itest Hub` must not take the root's groups either.
    const web = webWith([{ Id: 3 }, { Id: 7 }, notFound()])
    const root = associated({ Id: 9 }, { Id: 8 }, { Id: 10 })
    const items = await prefixedArtefacts(connection(web, root), 'SPJS Itest')

    assert.deepEqual(
      items.map((item) => `${item.kind}: ${item.name}`),
      ['site group: SPJS Itest Readers']
    )
  })

  it('plans nothing when it cannot tell which groups are associated', async () => {
    // A 503 (or a timeout) on that read used to leave the owners group unprotected.
    const unavailable = Object.assign(new Error('Service unavailable'), {
      status: 503
    })
    const web = webWith([unavailable, { Id: 7 }, { Id: 10 }])
    await assert.rejects(
      prefixedArtefacts(
        connection(web, associated({ Id: 9 }, { Id: 8 }, { Id: 10 })),
        'SPJS Itest'
      ),
      /Service unavailable/
    )
  })
})

describe('harness: suites', () => {
  const available = ['lists', 'lists-datarows', 'reachability', 'security']

  it('runs reachability first, then the suites asked for, by any case', () => {
    const names = (files: string[]) => files.map((file) => path.basename(file))
    assert.deepEqual(names(suiteFiles([], available)), [
      'reachability.test.ts',
      'lists.test.ts',
      'lists-datarows.test.ts',
      'security.test.ts'
    ])
    // A handler name selects its suites.
    assert.deepEqual(names(suiteFiles(['Lists'], available)), [
      'reachability.test.ts',
      'lists.test.ts',
      'lists-datarows.test.ts'
    ])
    assert.deepEqual(names(suiteFiles(['lists-datarows'], available)), [
      'reachability.test.ts',
      'lists-datarows.test.ts'
    ])
    assert.deepEqual(names(suiteFiles(['Security'], available)), [
      'reachability.test.ts',
      'security.test.ts'
    ])
    assert.deepEqual(names(suiteFiles(['reachability'], available)), [
      'reachability.test.ts'
    ])
    assert.throws(
      () => suiteFiles(['Bogus'], available),
      /Unknown suite\(s\): bogus/
    )
  })
})

describe('harness: fixtures', () => {
  it('loads a fresh copy without $schema', () => {
    const first = loadFixture('smoke')
    first.Lists[0].Title = 'changed'

    assert.equal(loadFixture('smoke').Lists[0].Title, 'SPJS Itest Smoke')
    assert.equal(first.$schema, undefined)
  })
})

describe('harness: redaction', () => {
  // Built at run time, so this file does not trip the scan for tenant URLs itself.
  const tenant = ['north', 'wind'].join('')

  it('removes cookie values, bearer tokens and tenant hosts, not the example hosts', () => {
    const cookie = 'A'.repeat(40)
    assert.equal(
      redact(
        `set-cookie: FedAuth=${cookie}; path=/ Cookie: FedAuth=${cookie}; rtFa=${cookie}`
      ),
      'set-cookie: <redacted>'
    )
    assert.equal(
      redact(`Cookie: rtFa=${cookie}; x=1`),
      'Cookie: rtFa=<redacted>; x=1'
    )
    assert.equal(
      redact(`Authorization: Bearer ${'ey' + 'J'.repeat(20)}`),
      'Authorization: Bearer <redacted>'
    )
    assert.equal(
      redact(
        `https://${tenant}.sharepoint.com/sites/x and https://${tenant}-admin.sharepoint.com`
      ),
      'https://<tenant>.sharepoint.com/sites/x and https://<tenant>-admin.sharepoint.com'
    )
    assert.equal(
      redact('https://contoso.sharepoint.com/sites/x'),
      'https://contoso.sharepoint.com/sites/x'
    )
  })

  it('describes an error by its message and status, never its response', () => {
    const error = Object.assign(
      new Error(`Forbidden at https://${tenant}.sharepoint.com/sites/x`),
      {
        status: 403,
        response: { headers: { 'set-cookie': `FedAuth=${'A'.repeat(40)}` } }
      }
    )
    assert.equal(
      describeError(error),
      'Forbidden at https://<tenant>.sharepoint.com/sites/x (403)'
    )
    assert.equal(describeError('plain'), 'plain')
    assert.equal(describeError(undefined), 'undefined')
  })

  it('redacts a JUnit report and keeps it valid XML', () => {
    // The CI runner rewrites test-results/*.xml with this.
    assert.equal(
      redactXml(
        `<failure message="List does not exist at site with URL 'https://${tenant}.sharepoint.com/sites/x'">`
      ),
      `<failure message="List does not exist at site with URL 'https://&lt;tenant&gt;.sharepoint.com/sites/x'">`
    )
  })
})
