import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { rmSync } from 'node:fs'
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
// Deliberately no `import '@pnp/sp/presets/all'` here: test files run in their
// own process, so this proves connect() loads what it needs by itself.
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
    const requests: Array<{ url: string; cookie: string }> = []
    const realFetch = globalThis.fetch
    globalThis.fetch = (async (url: string, init: any) => {
      requests.push({
        url: url.toString(),
        cookie: new Headers(init.headers).get('Cookie')
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

describe('runner: access denied', () => {
  const site = 'https://contoso.sharepoint.com/sites/spjs-itest'

  it('names the account to give access to the site', () => {
    const cookies = accessDenied({ status: 403 }, site, 'cookies')
    assert.match(cookies, /refused the cookies sign-in to .*spjs-itest \(403\)/)
    assert.match(cookies, /E2E_USERNAME/)
    assert.match(cookies, /owner/)
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
