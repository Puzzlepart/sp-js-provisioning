import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { rmSync } from 'node:fs'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
// Deliberately no `import '@pnp/sp/presets/all'` here: test files run in their
// own process, so this proves connect() loads what it needs by itself.
import { DefaultParse } from '@pnp/queryable'
import { DefaultHeaders, DefaultInit, spfi } from '@pnp/sp'
import { fetchWithRetry } from '../../debug/lib/auth'
import { accessDenied, connect } from '../../debug/lib/connect'
import { readEnv } from '../../debug/lib/env'

describe('runner: connect', () => {
  it('signs in with the session cookie and builds the SPFx context from what SharePoint says', async () => {
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'spjs-connect-'))
    const file = path.join(directory, 'user.json')
    fs.writeFileSync(
      file,
      JSON.stringify({
        cookies: [
          {
            name: 'FedAuth',
            value: 'session',
            domain: 'contoso.sharepoint.com',
            expires: -1
          }
        ]
      })
    )
    const requests: Array<{ url: string; cookie: string; signal: boolean }> = []
    const realFetch = globalThis.fetch
    globalThis.fetch = (async (url: string, init: any) => {
      requests.push({
        url: url.toString(),
        cookie: new Headers(init.headers).get('Cookie'),
        signal: Boolean(init.signal)
      })
      const body = url.toString().includes('/_api/site')
        ? {
            Id: 'site-id',
            Url: 'https://contoso.sharepoint.com/sites/hub',
            ServerRelativeUrl: '/sites/hub'
          }
        : {
            Id: 'web-id',
            Title: 'Sub',
            Url: 'https://contoso.sharepoint.com/sites/hub/sub',
            ServerRelativeUrl: '/sites/hub/sub',
            Language: 1044
          }
      return new Response(JSON.stringify(body), {
        status: 200,
        headers: { 'Content-Type': 'application/json' }
      })
    }) as any
    try {
      const connection = await connect(
        'https://contoso.sharepoint.com/sites/hub/sub/',
        'cookies',
        readEnv({ INTEGRATION_COOKIES_FILE: file })
      )

      assert.equal(
        connection.url,
        'https://contoso.sharepoint.com/sites/hub/sub'
      )
      assert.equal(connection.mode, 'cookies')
      // `{site}` and `{sitecollection}` resolve from pageContext.site: the site
      // collection, as on a page, also when the target is a subweb.
      assert.deepEqual(connection.context.pageContext.site, {
        id: 'site-id',
        absoluteUrl: 'https://contoso.sharepoint.com/sites/hub',
        serverRelativeUrl: '/sites/hub'
      })
      assert.equal(
        connection.context.pageContext.web.serverRelativeUrl,
        '/sites/hub/sub'
      )
      assert.equal(requests.length, 2)
      for (const request of requests) {
        assert.equal(request.cookie, 'FedAuth=session')
        // Every request is bounded (REQUEST_TIMEOUT).
        assert.equal(request.signal, true)
        assert.match(
          request.url,
          /^https:\/\/contoso\.sharepoint\.com\/sites\/hub\/sub\/_api\//
        )
      }
    } finally {
      globalThis.fetch = realFetch
      rmSync(directory, { recursive: true, force: true })
    }
  })
})

describe('runner: requests with a time limit per attempt', () => {
  const site = 'https://contoso.sharepoint.com/sites/x'
  const json = (status: number, headers: Record<string, string> = {}) =>
    new Response(JSON.stringify({ Title: 'SPJS Itest' }), {
      status,
      headers: { 'Content-Type': 'application/json', ...headers }
    })

  /**
   * Runs `fn` with `fetch` answering from `answers` in turn, and the signals it was called with.
   * An answer `'hang'` never comes: it settles only when the signal aborts.
   */
  async function withFetch(
    answers: Array<Response | Error | 'hang'>,
    fn: () => Promise<void>
  ): Promise<AbortSignal[]> {
    const realFetch = globalThis.fetch
    const signals: AbortSignal[] = []
    // AbortSignal.timeout's timer does not hold the process open (a real request's socket does).
    const keepAlive = setTimeout(() => undefined, 10_000)
    globalThis.fetch = ((_url: string, init: any) => {
      signals.push(init.signal)
      const answer = answers.shift()
      if (answer instanceof Error) return Promise.reject(answer)
      if (answer !== 'hang') return Promise.resolve(answer)
      return new Promise((_resolve, reject) => {
        if (init.signal.aborted) reject(init.signal.reason)
        else
          init.signal.addEventListener('abort', () =>
            reject(init.signal.reason)
          )
      })
    }) as any
    try {
      await fn()
    } finally {
      clearTimeout(keepAlive)
      globalThis.fetch = realFetch
    }
    return signals
  }

  const retried: string[] = []
  const sp = (ms: number) =>
    spfi(site).using(
      DefaultHeaders(),
      DefaultInit(),
      fetchWithRetry(ms, 3, (what) => retried.push(what)),
      DefaultParse()
    )

  it('fails an attempt SharePoint does not answer, without sending it again', async () => {
    const started = Date.now()
    const signals = await withFetch(['hang', json(200)], async () => {
      await assert.rejects(sp(50).web(), /timeout/i)
    })
    assert.ok(Date.now() - started < 5_000, 'bounded, not 3 x 300 s')
    // Sent again, a write SharePoint is still carrying out could be repeated.
    assert.equal(signals.length, 1)
    assert.equal(signals[0].aborted, true)
  })

  it('waits on a throttled answer and tries again with a fresh limit', async () => {
    // Retry-After 0: the default back-off (400 ms), longer than the 50 ms limit; one limit for
    // the whole request would fail the second attempt.
    const signals = await withFetch(
      [json(429, { 'Retry-After': '0' }), json(200)],
      async () => {
        assert.equal((await sp(50).web()).Title, 'SPJS Itest')
      }
    )
    assert.equal(signals.length, 2)
    assert.notEqual(signals[0], signals[1])
    // Throttling is reported, so a slow run says why.
    assert.ok(retried.includes('GET answered 429; waiting 0 s'), retried.join())
  })

  it('tries a network error again, three attempts in all', async () => {
    const failed = () => new TypeError('fetch failed')
    await withFetch([failed(), json(200)], async () => {
      assert.equal((await sp(1_000).web()).Title, 'SPJS Itest')
    })
    const signals = await withFetch(
      [failed(), failed(), failed()],
      async () => {
        await assert.rejects(sp(1_000).web(), /fetch failed/)
      }
    )
    assert.equal(signals.length, 3)
  })
})

describe('runner: access denied', () => {
  const site = 'https://contoso.sharepoint.com/sites/spjs-itest'

  it('names the account to give access to the site', () => {
    const cookies = accessDenied({ status: 403 }, site, 'cookies')
    assert.match(cookies, /refused the cookies sign-in to .*spjs-itest \(403\)/)
    assert.match(cookies, /E2E_USERNAME/)
    assert.match(cookies, /site collection administrator/)
    assert.match(
      accessDenied({ status: 401 }, site, 'certificate'),
      /Sites\.FullControl\.All/
    )
  })

  it('leaves other errors alone', () => {
    assert.equal(accessDenied({ status: 500 }, site, 'cookies'), undefined)
    assert.equal(accessDenied(new Error('network'), site, 'cookies'), undefined)
  })
})
