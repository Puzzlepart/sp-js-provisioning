import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { LogLevel } from '@pnp/logging'
import { DefaultParse, RejectOnError, ResolveOnData } from '@pnp/queryable'
import { spPost, SPQueryable, Web } from '@pnp/sp/presets/all'
import { RetryNotAcceptable } from '../../src/util/retry'
import { WebProvisioner } from '../../src/webprovisioner'
import { captureLogs, levelOf } from '../helpers/logs'

const BASE = 'https://contoso.sharepoint.com/sites/demo'

/**
 * A web whose transport (`send`, as `BrowserFetchWithRetry` is in SPFx) answers each request with
 * the next status in `statuses`, the last one for every request after it: 200 with a web as JSON,
 * anything else with an empty body, as SharePoint's 406 has.
 *
 * @param statuses - The statuses, in order
 * @param headers - Headers for a response that is not 200
 */
function webAnswering(
  statuses: number[],
  headers: Record<string, string> = {}
) {
  const sent: string[] = []
  const transport = (instance: any) => {
    instance.on.send(async (url: URL, init: RequestInit) => {
      const status = statuses.length > 1 ? statuses.shift() : statuses[0]
      sent.push(
        `${init.method} ${decodeURIComponent(
          String(url).slice(BASE.length + 1)
        )} ${status}`
      )
      return status === 200
        ? new Response(
            JSON.stringify({ Title: 'Demo', ServerRelativeUrl: '/sites/demo' }),
            { status, headers: { 'Content-Type': 'application/json' } }
          )
        : new Response('', { status, statusText: 'Not Acceptable', headers })
    })
    return instance
  }
  const web = Web(BASE).using(
    DefaultParse(),
    ResolveOnData(),
    RejectOnError(),
    transport
  )
  return { web, sent }
}

describe('RetryNotAcceptable', () => {
  it('tries a GET SharePoint answers 406 again, and returns what the next try gets', async () => {
    const { web, sent } = webAnswering([406, 200])
    const retries: string[] = []

    const data = await web
      .using(
        RetryNotAcceptable({
          interval: 1,
          onRetry: (message) => retries.push(message)
        })
      )
      .select('Title')<{ Title: string }>()

    assert.equal(data.Title, 'Demo')
    assert.deepEqual(sent, [
      'GET _api/web?$select=Title 406',
      'GET _api/web?$select=Title 200'
    ])
    assert.equal(retries.length, 1)
    assert.match(
      retries[0],
      /^SharePoint answered 406 \(Not Acceptable\) to GET https:\/\/contoso\.sharepoint\.com\/sites\/demo\/_api\/web\?%24select=Title: trying again in 0\.001 s \(1 of 3\)$/
    )
  })

  it('waits what Retry-After says', async () => {
    const { web } = webAnswering([406, 200], { 'Retry-After': '0' })
    const retries: string[] = []

    await web
      .using(
        RetryNotAcceptable({ onRetry: (message) => retries.push(message) })
      )
      .select('Title')()

    // Not the default first wait of 1 s.
    assert.match(retries[0], /trying again in 0 s \(1 of 3\)$/)
  })

  it('gives up after three more tries, rejecting with the 406', async () => {
    const { web, sent } = webAnswering([406])

    await assert.rejects(
      web.using(RetryNotAcceptable({ interval: 1 })).select('Title')(),
      (error: any) => error.status === 406
    )
    assert.equal(sent.length, 4)
  })

  it('does not try a write again, nor another status', async () => {
    const post = webAnswering([406])
    await assert.rejects(
      spPost(
        SPQueryable(
          post.web.using(RetryNotAcceptable({ interval: 1 })),
          'lists'
        )
      ),
      (error: any) => error.status === 406
    )
    assert.deepEqual(post.sent, ['POST _api/web/lists 406'])

    const notFound = webAnswering([404, 200])
    await assert.rejects(
      notFound.web.using(RetryNotAcceptable({ interval: 1 })).select('Title')(),
      (error: any) => error.status === 404
    )
    assert.equal(notFound.sent.length, 1)
  })
})

describe('WebProvisioner: a GET SharePoint answers 406', () => {
  it("is tried again, on a copy of the caller's web, and logged at Info", async () => {
    // The first request (the web, in setup) answers 406; Retry-After 0 keeps the test fast.
    const { web, sent } = webAnswering([406, 200], {
      'Retry-After': '0'
    })
    const [callerSend] = web.on.send.toArray()

    const { entries } = await captureLogs(() =>
      new WebProvisioner(web).setup({}).applyTemplate({})
    )

    assert.deepEqual(sent.slice(0, 2), ['GET _api/web 406', 'GET _api/web 200'])
    assert.equal(
      levelOf(entries, /\(RetryNotAcceptable\): SharePoint answered 406/),
      LogLevel.Info
    )
    // The caller's web keeps its own transport, unwrapped.
    assert.deepEqual(web.on.send.toArray(), [callerSend])
  })
})
