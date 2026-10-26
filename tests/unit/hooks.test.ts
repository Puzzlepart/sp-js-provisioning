import assert from 'node:assert/strict'
import { afterEach, beforeEach, describe, it } from 'node:test'
import { Hooks } from '../../src/handlers/hooks'

/** How long a hook run may take before the test calls it hung. */
const DEADLINE = 2000

/**
 * A Hooks handler that polls a 202's `Location` without the five-second wait.
 */
class FastHooks extends Hooks {
  protected pollInterval = 1

  constructor() {
    super({})
  }
}

/**
 * A request the handler sent through `fetch`.
 */
interface ISent {
  url: string
  method: string
  headers: Record<string, string>
  body?: string
}

/**
 * A fake global `fetch` that records what the handler sends and answers from `answers`, in order.
 * A URL that is not absolute is refused the way Node's `fetch` refuses it.
 */
function fakeFetch(answers: Array<() => Response>) {
  const sent: ISent[] = []
  const fetch = async (url: string, init: any = {}): Promise<Response> => {
    if (!/^https?:\/\//.test(String(url)))
      throw new TypeError(`Failed to parse URL from ${url}`)
    sent.push({
      url: String(url),
      method: init.method,
      headers: { ...init.headers },
      body: init.body
    })
    const answer = answers.shift()
    if (!answer) throw new Error(`No answer left for ${init.method} ${url}`)
    return answer()
  }
  return { fetch, sent }
}

/**
 * A fake web whose property bag is `properties`.
 *
 * @param properties - What `web.allProperties()` returns
 */
function fakeWeb(properties: Record<string, any> = {}): any {
  return { allProperties: async () => properties }
}

/**
 * `promise`, or a rejection when it has not settled within `ms`: a hung run fails the test
 * instead of stalling the suite.
 *
 * @param promise - The run
 * @param ms - The deadline
 */
async function within<T>(promise: Promise<T>, ms = DEADLINE): Promise<T> {
  let timer: any
  const deadline = new Promise<never>((resolve, reject) => {
    timer = setTimeout(
      () => reject(new Error(`the hook run did not settle within ${ms} ms`)),
      ms
    )
  })
  try {
    return await Promise.race([promise, deadline])
  } finally {
    clearTimeout(timer)
  }
}

const accepted = (location?: string) => () =>
  new Response(undefined, {
    status: 202,
    headers: location ? { Location: location } : {}
  })
const ok = () => new Response('{}', { status: 200 })
const failed = () =>
  new Response(JSON.stringify({ error: 'job failed' }), {
    status: 500,
    statusText: 'Internal Server Error'
  })

const slow = {
  Title: 'Slow',
  Url: 'https://hooks.contoso.com/jobs',
  Method: 'POST',
  Headers: { 'Content-Type': 'application/json', 'X-Hook': 'slow' },
  Body: { job: 'provision' }
}

describe('Hooks', () => {
  const realFetch = (globalThis as any).fetch
  afterEach(() => {
    ;(globalThis as any).fetch = realFetch
  })
  let use: (answers: Array<() => Response>) => ISent[]
  beforeEach(() => {
    use = (answers) => {
      const fake = fakeFetch(answers)
      ;(globalThis as any).fetch = fake.fetch
      return fake.sent
    }
  })

  it('sends a GET hook with its headers and no body', async () => {
    const sent = use([ok])
    await within(
      new FastHooks().ProvisionObjects(fakeWeb({ a: 1 }), [
        {
          Title: 'Get',
          Url: 'https://hooks.contoso.com/ping',
          Method: 'GET',
          Headers: { 'X-Hook': 'get' }
        } as any
      ])
    )
    assert.deepEqual(sent, [
      {
        url: 'https://hooks.contoso.com/ping',
        method: 'GET',
        headers: { 'X-Hook': 'get' },
        body: undefined
      }
    ])
  })

  it('posts Body merged with the property bag, the property bag winning a clash', async () => {
    const sent = use([ok])
    await within(
      new FastHooks().ProvisionObjects(
        fakeWeb({ vti_title: 'Web', job: 'from the bag' }),
        [{ ...slow, Body: { job: 'provision', kind: 'hub' } } as any]
      )
    )
    assert.equal(sent[0].method, 'POST')
    assert.deepEqual(JSON.parse(sent[0].body), {
      job: 'from the bag',
      kind: 'hub',
      vti_title: 'Web'
    })
  })

  it('resolves when a 202 poll ends in 200 (D4)', async () => {
    const sent = use([accepted('https://hooks.contoso.com/jobs/1'), ok])
    await within(new FastHooks().ProvisionObjects(fakeWeb(), [slow as any]))
    assert.deepEqual(
      sent.map((request) => [request.method, request.url]),
      [
        ['POST', 'https://hooks.contoso.com/jobs'],
        ['GET', 'https://hooks.contoso.com/jobs/1']
      ]
    )
    // The poll carries the hook's headers.
    assert.deepEqual(sent[1].headers, slow.Headers)
    assert.equal(sent[1].body, undefined)
  })

  it('keeps polling while the Location answers 202', async () => {
    const location = 'https://hooks.contoso.com/jobs/1'
    const sent = use([accepted(location), accepted(location), ok])
    await within(new FastHooks().ProvisionObjects(fakeWeb(), [slow as any]))
    assert.deepEqual(
      sent.map((request) => request.method),
      ['POST', 'GET', 'GET']
    )
  })

  it('rejects when a poll fails, naming the hook and the status', async () => {
    use([accepted('https://hooks.contoso.com/jobs/1'), failed])
    await assert.rejects(
      within(new FastHooks().ProvisionObjects(fakeWeb(), [slow as any])),
      (error: Error) => {
        assert.match(error.message, /500 - Internal Server Error \| job failed/)
        assert.match(error.message, /Hook 1\/1: Slow/)
        return true
      }
    )
  })

  it('rejects when a poll cannot be sent', async () => {
    use([
      accepted('https://hooks.contoso.com/jobs/1'),
      () => {
        throw new TypeError('fetch failed')
      }
    ])
    await assert.rejects(
      within(new FastHooks().ProvisionObjects(fakeWeb(), [slow as any])),
      /fetch failed - Hook 1\/1: Slow/
    )
  })

  it('names the hook whose first request cannot be sent', async () => {
    // A network or CORS failure used to surface as a bare "Failed to fetch", naming no hook.
    const cannotSend = () => {
      throw new TypeError('fetch failed')
    }
    const get = {
      Title: 'Get',
      Url: 'https://hooks.contoso.com/ping',
      Method: 'GET'
    }
    use([cannotSend])
    await assert.rejects(
      within(new FastHooks().ProvisionObjects(fakeWeb(), [get as any])),
      /fetch failed - Hook 1\/1: Get/
    )
    // Two hooks run together: the POST is the one that fails, and is the one named.
    use([ok, cannotSend])
    await assert.rejects(
      within(
        new FastHooks().ProvisionObjects(fakeWeb(), [get as any, slow as any])
      ),
      /fetch failed - Hook 2\/2: Slow/
    )
  })

  it('polls a relative Location against the URL that answered', async () => {
    // After a redirect the response URL is not the hook's.
    const redirected = () => {
      const response = accepted('jobs/1')()
      Object.defineProperty(response, 'url', {
        value: 'https://other.contoso.com/queue/'
      })
      return response
    }
    const sent = use([redirected, ok])
    await within(new FastHooks().ProvisionObjects(fakeWeb(), [slow as any]))
    assert.equal(sent[1].url, 'https://other.contoso.com/queue/jobs/1')
  })

  it('rejects a Location it cannot resolve, naming the hook', async () => {
    use([accepted('http://[')])
    await assert.rejects(
      within(new FastHooks().ProvisionObjects(fakeWeb(), [slow as any])),
      /Invalid Location 'http:\/\/\[' to poll - Hook 1\/1: Slow/
    )
  })

  it('polls a relative Location against the hook URL', async () => {
    // The Location used to go to fetch as is: Node refused a relative URL, a
    // browser resolved it against the SharePoint page.
    const sent = use([accepted('/jobs/1'), ok])
    await within(new FastHooks().ProvisionObjects(fakeWeb(), [slow as any]))
    assert.equal(sent[1].url, 'https://hooks.contoso.com/jobs/1')
  })

  it('rejects a 202 without a Location, naming the hook, and polls nothing', async () => {
    // It used to poll `null`: in a browser a GET to <page folder>/null with
    // the hook's headers.
    const sent = use([accepted()])
    await assert.rejects(
      within(new FastHooks().ProvisionObjects(fakeWeb(), [slow as any])),
      /202 Accepted without a Location to poll - Hook 1\/1: Slow/
    )
    assert.deepEqual(
      sent.map((request) => request.method),
      ['POST']
    )
  })
})
