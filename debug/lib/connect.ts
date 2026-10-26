import { spfi, SPFI } from '@pnp/sp'
// Side effect: adds `site`, `web` and the rest to SPFI. The type import below alone is erased.
import '@pnp/sp/presets/all'
import { IWeb } from '@pnp/sp/presets/all'
import { signedIn } from './auth'
import { UsageError } from './cli'
import {
  AuthMode,
  authMode,
  env,
  IEnv,
  isWithinSite,
  normaliseSiteUrl
} from './env'

/**
 * What `_api/site` says about the site collection.
 */
export interface ISiteInfo {
  Id: string
  Url: string
  ServerRelativeUrl: string
}

/**
 * What `_api/web` says about the web.
 */
export interface IWebInfo {
  Id: string
  Title: string
  Url: string
  ServerRelativeUrl: string
  Language: number
}

/**
 * A site the runner talks to: a signed-in client, what the site says about itself, and the SPFx
 * context stand-in the engine is set up with.
 */
export interface ISiteConnection {
  url: string
  sp: SPFI
  web: IWeb
  siteInfo: ISiteInfo
  webInfo: IWebInfo
  context: any
  mode: AuthMode
}

/**
 * The SPFx context a web part on the web would have, as far as the engine reads it:
 * `pageContext.site` (for the `{site}`, `{sitecollection}` and `{wpgallery}` tokens) is the site
 * collection, `pageContext.web` the web, as on a page.
 *
 * @param site - What `_api/site` says
 * @param web - What `_api/web` says
 */
export function spfxContext(site: ISiteInfo, web: IWebInfo): any {
  return {
    pageContext: {
      site: {
        id: site.Id,
        absoluteUrl: site.Url,
        serverRelativeUrl: site.ServerRelativeUrl
      },
      web: {
        id: web.Id,
        title: web.Title,
        absoluteUrl: web.Url,
        serverRelativeUrl: web.ServerRelativeUrl,
        language: web.Language
      }
    }
  }
}

/**
 * Connects to a site in the configured sign-in mode, and reads what the site says about itself.
 *
 * @param url - The site
 * @param mode - Sign-in mode (from `.env` by default)
 * @param configuration - The configuration (from `.env` by default)
 */
export async function connect(
  url: string,
  mode: AuthMode | undefined = authMode(),
  configuration: IEnv = env
): Promise<ISiteConnection> {
  if (!mode)
    throw new Error(
      'No sign-in configured: set CI_TENANT, CI_CLIENT_ID and CI_CERT_BASE64, or point ' +
        'INTEGRATION_COOKIES_FILE at a Playwright storage state (see .env.example)'
    )
  const siteUrl = url.replace(/\/+$/, '')
  const sp = spfi(siteUrl).using(signedIn(mode, siteUrl, configuration))
  const [siteInfo, webInfo] = await Promise.all([
    sp.site.select('Id', 'Url', 'ServerRelativeUrl')<ISiteInfo>(),
    sp.web.select(
      'Id',
      'Title',
      'Url',
      'ServerRelativeUrl',
      'Language'
    )<IWebInfo>()
  ])
  return {
    url: siteUrl,
    sp,
    // A fresh queryable on every use: PnPjs 4's `select()` sets the query on the instance, so a
    // shared one would carry a reader's `$select` into the next request.
    get web() {
      return sp.web
    },
    siteInfo,
    webInfo,
    context: spfxContext(siteInfo, webInfo),
    mode
  }
}

/**
 * A clear message for SharePoint refusing the sign-in (401/403), naming the account to give
 * access to the site; `undefined` for any other error.
 *
 * @param error - The error from `connect()`
 * @param site - The site
 * @param mode - The sign-in mode
 */
export function accessDenied(
  error: any,
  site: string,
  mode: AuthMode | undefined
): string | undefined {
  const status = error && error.status
  if (status !== 401 && status !== 403) return undefined
  const account =
    mode === 'certificate'
      ? 'The app (CI_CLIENT_ID) needs Sites.FullControl.All, or Sites.Selected with a grant on the site.'
      : "The cookie session is PP365's e2e user (E2E_USERNAME in prosjektportalen365/e2e/.env): " +
        'make that account a site collection administrator of the site (the tests need it; ' +
        '`provision` alone needs an owner). If it already is, renew the session with ' +
        '`npx playwright test --project=setup` in prosjektportalen365/e2e.'
  return `SharePoint refused the ${mode} sign-in to ${site} (${status}).\n${account}`
}

/**
 * Refuses a write target that is not the reserved site (`INTEGRATION_SITE_URL`) or a web below
 * it, unless the caller names that exact target with `--allow-site`. A template run cannot be
 * undone, so nothing writes to a site by accident (a customer hub, PP365's test sites).
 *
 * @param target - The site a command is about to write to
 * @param allowSite - `--allow-site`: the target, repeated to confirm it
 * @param configuration - The configuration
 */
export function assertWritable(
  target: string,
  allowSite?: string,
  configuration: IEnv = env
): void {
  const normalised = normaliseSiteUrl(target)
  if (!normalised)
    throw new UsageError(
      `Refusing to write to ${target}: not a plain http(s) URL (no encoded dots or slashes)`
    )
  if (allowSite && normaliseSiteUrl(allowSite) === normalised) return
  if (configuration.siteUrl && isWithinSite(target, configuration.siteUrl))
    return
  throw new UsageError(
    `Refusing to write to ${target}: it is not INTEGRATION_SITE_URL (${
      configuration.siteUrl ?? 'not set'
    }) or a web below it. Pass --allow-site ${target} to write there anyway.`
  )
}

/**
 * After connecting: refuses a web SharePoint reports outside the reserved site, whatever the URL
 * given looked like (the last line of the site guard).
 *
 * @param conn - The connection
 * @param allowSite - `--allow-site`, as given
 * @param configuration - The configuration
 */
export function assertConnectedWritable(
  conn: ISiteConnection,
  allowSite?: string,
  configuration: IEnv = env
): void {
  assertWritable(conn.webInfo.Url, allowSite, configuration)
}
