import { ISiteConnection } from '../connect'

/**
 * A request a hook listener received, as Node gives it.
 */
export interface IHookRequest {
  method: string
  /** Path and query. */
  url: string
  /** Header names in lower case. */
  headers: Record<string, string | string[] | undefined>
  body: string
  /** `Date.now()` on arrival. */
  receivedAt: number
}

/**
 * A hook call as the template declares it: method, path, the headers a template sets and the
 * parsed body. Without what the HTTP client adds on its own (user agent, encoding, host).
 */
export interface IHookCall {
  method: string
  url: string
  headers: Record<string, string>
  /** The JSON body, parsed; absent for a call without one. */
  body?: any
}

/**
 * What the Hooks handler did in one run: the calls a listener received, sorted, and the web's
 * property bag, which the handler merges into each POST body.
 */
export interface IHooksState {
  calls: IHookCall[]
  properties: Record<string, any>
}

/** The request headers a hook call keeps: the ones a template sets. */
export const HOOK_HEADERS = /^(accept|content-type|x-spjs-.*)$/

/**
 * A received request as a comparable call.
 *
 * @param request - The request
 */
export function hookCall(request: IHookRequest): IHookCall {
  const headers: Record<string, string> = {}
  for (const name of Object.keys(request.headers).sort()) {
    const value = request.headers[name]
    if (HOOK_HEADERS.test(name) && value !== undefined)
      headers[name] = Array.isArray(value) ? value.join(', ') : value
  }
  const call: IHookCall = {
    method: request.method,
    url: request.url,
    headers
  }
  if (request.body) call.body = JSON.parse(request.body)
  return call
}

/**
 * Reads back one run of the Hooks handler: the calls a listener received during the run (hooks
 * run concurrently, so they are sorted by URL and method), and the web's property bag.
 *
 * @param conn - The web the template was applied to
 * @param requests - The requests the listener received during the run
 */
export async function readHooks(
  conn: ISiteConnection,
  requests: IHookRequest[]
): Promise<IHooksState> {
  const properties = await conn.web.allProperties<Record<string, any>>()
  const calls = requests
    .map(hookCall)
    .sort((a, b) =>
      `${a.url} ${a.method}` < `${b.url} ${b.method}` ? -1 : 1
    )
  return { calls, properties }
}
