/**
 * Node stub for `spfx-jsom`.
 *
 * `spfx-jsom` loads the SharePoint JSOM (`SP.ClientContext`) into a browser page
 * and transitively pulls `@microsoft/sp-*`, which can't even be imported under
 * Node (it `require`s `.resx` files). The local provisioning runner replaces it
 * with this stub so the engine module graph loads.
 *
 * REST-only handlers (SiteFields, Lists) merely *initialize* a jsom context and
 * never touch it — they keep working. Handlers that actually use JSOM
 * (ContentTypes, Taxonomy) hit the throwing proxy below and fail with a clear
 * message; the runner auto-skips them by default.
 */
function jsomUnavailable(): never {
  throw new Error(
    'JSOM (spfx-jsom) is not available in the local Node runner. ' +
      'ContentTypes and Taxonomy can only be provisioned in the browser/SPFx.'
  )
}

const throwingProxy: any = new Proxy(
  {},
  {
    get: () => jsomUnavailable(),
    apply: () => jsomUnavailable()
  }
)

/**
 * Stand-in for the real `JsomContext`. The index signature keeps the handlers
 * type-checking against this stub (the real context exposes `web`,
 * `clientContext` and more, all of which resolve to the throwing proxy here).
 */
export class JsomContext {
  [key: string]: any
}

/** Throws - JSOM queries cannot be executed under Node. */
export function ExecuteJsomQuery(...args: any[]): any {
  return jsomUnavailable()
}

const initSpfxJsom = async (...args: any[]): Promise<any> => {
  return { jsomContext: { web: throwingProxy, clientContext: throwingProxy } }
}

export default initSpfxJsom
