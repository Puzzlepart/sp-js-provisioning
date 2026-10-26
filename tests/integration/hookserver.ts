import * as http from 'http'
import { after, before } from 'node:test'
import { IHookRequest } from '../../debug/lib/readback/hooks'

/**
 * A local HTTP listener for hooks. Hooks run in the calling process, so a listener on 127.0.0.1
 * is reachable from a run against SharePoint Online.
 *
 * Routes: `POST /slow` answers `202 Accepted` with an absolute `Location` (`/slow/status`), which
 * answers `200`; every other request answers `200` with `{}`.
 */
export interface IHookServer {
  /** The listener's origin, e.g. `http://127.0.0.1:53124`: the `@@hookUrl@@` value. */
  url: string
  /** Every request received, in order of arrival. */
  requests: IHookRequest[]
  /** The requests received since the last call. */
  take(): IHookRequest[]
  close(): Promise<void>
}

/**
 * Starts a hook listener on a free port on 127.0.0.1.
 */
export function startHookServer(): Promise<IHookServer> {
  const requests: IHookRequest[] = []
  let taken = 0
  let origin = ''
  const server = http.createServer((request, response) => {
    const chunks: any[] = []
    request.on('data', (chunk: any) => chunks.push(chunk))
    request.on('end', () => {
      requests.push({
        method: request.method || '',
        url: request.url || '',
        headers: { ...(request.headers as any) },
        body: Buffer.concat(chunks).toString('utf8'),
        receivedAt: Date.now()
      })
      if (request.method === 'POST' && request.url === '/slow') {
        response.writeHead(202, { Location: `${origin}/slow/status` })
        response.end()
        return
      }
      response.writeHead(200, { 'Content-Type': 'application/json' })
      response.end('{}')
    })
  })
  return new Promise((resolve, reject) => {
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      origin = `http://127.0.0.1:${server.address().port}`
      resolve({
        url: origin,
        requests,
        take: () => {
          const since = requests.slice(taken)
          taken = requests.length
          return since
        },
        close: () =>
          new Promise<void>((resolveClose) => {
            // `fetch` keeps its connections alive: close them, or `close()` waits for them.
            const closeAll = (server as any).closeAllConnections
            if (closeAll) closeAll.call(server)
            server.close(() => resolveClose())
          })
      })
    })
  })
}

/**
 * A hook listener for the suite's tests: started in `before`, closed in `after`. Call it inside a
 * `describe`; it returns a getter, since the listener exists only once `before` has run.
 */
export function withHookServer(): () => IHookServer {
  let server: IHookServer | undefined
  before(async () => {
    server = await startHookServer()
  })
  after(async () => {
    if (server) await server.close()
  })
  return () => {
    if (!server) throw new Error('The hook listener is started in before()')
    return server
  }
}
