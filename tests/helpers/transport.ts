import { RejectOnError, ResolveOnData } from '@pnp/queryable'
import { Web } from '@pnp/sp/presets/all'

/** A request as it would go on the wire. */
export interface IWireRequest {
  /** GET, POST, or the tunnelled MERGE/DELETE PnPjs sends as a POST with X-HTTP-Method. */
  method: string
  /** The URL below the web (`_api/web/...`), decoded, without the query. */
  path: string
  /** The query, decoded (`$filter`, `$select`, `$top`, ...). */
  query: Record<string, string>
  /** The JSON body, parsed, when there is one; a CSOM request's XML as text. */
  body?: any
}

/**
 * A web whose requests never leave the process: each is recorded as it would go on the wire and
 * answered by `respond` (an empty object when it returns undefined; it may throw to fail the
 * request). Child queryables inherit the observer, so everything a handler builds from the web is
 * recorded. This is the contract layer's transport: the tests assert the recorded requests.
 *
 * @param respond - The answer for a request
 * @param base - The web's URL; a server-relative one models an SPFx context
 */
export function createWireWeb(
  respond?: (request: IWireRequest) => any,
  base = 'https://contoso.sharepoint.com/sites/demo'
): { web: any; requests: IWireRequest[] } {
  const requests: IWireRequest[] = []
  const prefix = base.replace(/\/+$/, '') + '/'
  const recording = (instance: any) => {
    instance.on.pre(async (url: string, init: any) => {
      const [address, search = ''] = url.split('?')
      const query = search
        .split('&')
        .filter(Boolean)
        .reduce((all: Record<string, string>, pair) => {
          const [key, value = ''] = pair.split('=')
          all[decodeURIComponent(key)] = decodeURIComponent(
            value.replace(/\+/g, ' ')
          )
          return all
        }, {})
      const headers = new Headers((init && init.headers) || {})
      const request: IWireRequest = {
        method: headers.get('X-HTTP-Method') || (init && init.method) || 'GET',
        path: decodeURIComponent(
          address.startsWith(prefix) ? address.slice(prefix.length) : address
        ),
        query,
        // A CSOM request (ProcessQuery) is XML: kept as text, for tests/helpers/csom.ts.
        body:
          init && typeof init.body === 'string' && init.body
            ? /^\s*</.test(init.body)
              ? init.body
              : JSON.parse(init.body)
            : undefined
      }
      requests.push(request)
      const result = respond ? respond(request) : undefined
      return [url, init, result === undefined ? {} : result]
    })
    return instance
  }
  // ResolveOnData/RejectOnError settle the request promise; they normally come from spfi.
  const web = Web(base).using(ResolveOnData(), RejectOnError(), recording)
  return { web, requests }
}

/**
 * The requests that write (anything but GET).
 *
 * @param requests - The recorded requests
 */
export const writes = (requests: IWireRequest[]): IWireRequest[] =>
  requests.filter((request) => request.method !== 'GET')
