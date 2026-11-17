import { delay, TimelinePipe } from '@pnp/core'
import { Queryable, QueryableSendObserver } from '@pnp/queryable'

/** How {@link RetryNotAcceptable} tries again. */
export interface IRetryNotAcceptableProps {
  /** Tries after the first (default 3). */
  retries?: number
  /** The first wait in milliseconds, doubled for each try after it (default 1000). */
  interval?: number
  /** Told of each retry, in words. */
  onRetry?: (message: string) => void
}

/**
 * Tries a GET again when SharePoint answers 406 (Not Acceptable): it does so with an empty body,
 * under load, right after a page is published (PP365 stress test, 2026-11-16: a project setup
 * stopped in ClientSidePages when PnPjs read the page's item after `publish`). Waits what
 * `Retry-After` says, or 1, 2 and 4 seconds, then hands the last response on as it is.
 *
 * It wraps the queryable's own `send`, so the caller's transport and retries stay as they are:
 * PnPjs's `SPFx` behavior (`BrowserFetchWithRetry`) and the local runner already try a 429, 503 or
 * 504 again, honouring `Retry-After`, and stacking more retries on throttling would only make it
 * worse. Other methods are not tried again: a write may have been done.
 *
 * @param props - How to try again
 */
export function RetryNotAcceptable(
  props: IRetryNotAcceptableProps = {}
): TimelinePipe<Queryable> {
  const { retries = 3, interval = 1000, onRetry } = props
  return (instance: Queryable) => {
    const [send] = instance.on.send.toArray() as QueryableSendObserver[]
    if (!send) return instance
    instance.on.send.replace(async function (
      this: Queryable,
      url: URL,
      init: RequestInit
    ): Promise<Response> {
      let wait = interval
      for (let attempt = 1; ; attempt++) {
        const response: Response = await Reflect.apply(send, this, [url, init])
        if (
          response.status !== 406 ||
          attempt > retries ||
          String(init.method || 'GET').toUpperCase() !== 'GET'
        )
          return response
        const after = Number.parseInt(
          response.headers.get('Retry-After') || '',
          10
        )
        const ms = Number.isNaN(after) ? wait : after * 1000
        if (onRetry)
          onRetry(
            `SharePoint answered 406 (Not Acceptable) to GET ${String(
              url
            )}: trying again in ${ms / 1000} s (${attempt} of ${retries})`
          )
        await delay(ms)
        wait *= 2
      }
    })
    return instance
  }
}
