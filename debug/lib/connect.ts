import { spfi, SPFI } from '@pnp/sp'
// Side effect: adds `site`, `web` and the rest to SPFI. The type import below alone is erased.
import '@pnp/sp/presets/all'
import { IWeb } from '@pnp/sp/presets/all'
import { signedIn } from './auth'
import { AuthMode, authMode, env, IEnv } from './env'

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
    web: sp.web,
    siteInfo,
    webInfo,
    context: spfxContext(siteInfo, webInfo),
    mode
  }
}
