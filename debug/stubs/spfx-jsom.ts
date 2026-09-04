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

export class JsomContext {}

export function ExecuteJsomQuery(): never {
  return jsomUnavailable()
}

const initSpfxJsom = async (
  _webServerRelativeUrl?: string
): Promise<{ jsomContext: any }> => {
  return { jsomContext: { web: throwingProxy, clientContext: throwingProxy } }
}

export default initSpfxJsom
