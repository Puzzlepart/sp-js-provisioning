import { ConfidentialClientApplication } from '@azure/msal-node'
import { combine, delay, isUrlAbsolute, TimelinePipe } from '@pnp/core'
import { DefaultParse, Queryable } from '@pnp/queryable'
import { DefaultHeaders, DefaultInit, RequestDigest } from '@pnp/sp'
import * as fs from 'fs'
import * as forge from 'node-forge'
import { AuthMode, env, ICertificateSignIn, IEnv } from './env'

/**
 * How a PnPjs client talks to SharePoint from Node, in either sign-in mode. Ported from PP365's
 * `integration/src/auth.ts` (releases/1.15). Nothing here writes a token, a cookie or the
 * certificate anywhere, and errors name the variable, never its value.
 */

/**
 * Relative URLs (`_api/web`) against the site, as SPFx resolves them against the page's web.
 * Prepended, so it runs before PnPjs parses the URL.
 *
 * @param baseUrl - The site
 */
function resolveAgainst(baseUrl: string): TimelinePipe<Queryable> {
  return (instance) => {
    instance.on.pre.prepend(async (url: string, init: any, result: any) => [
      isUrlAbsolute(url) ? url : combine(baseUrl, url),
      init,
      result
    ])
    return instance
  }
}

/**
 * A PFX opened without a password. One exported with an empty password derives its keys from a
 * terminating null, one exported with none from no bytes at all; node-forge opens only the one it
 * is told, so both are tried.
 *
 * @param pfxBase64 - The PFX, base64
 */
function openPfx(pfxBase64: string): forge.pkcs12.Pkcs12Pfx {
  let asn1: forge.asn1.Asn1
  try {
    asn1 = forge.asn1.fromDer(forge.util.decode64(pfxBase64))
  } catch {
    throw new Error('CI_CERT_BASE64 is not a PFX in base64')
  }
  try {
    return forge.pkcs12.pkcs12FromAsn1(asn1, '')
  } catch {
    try {
      return forge.pkcs12.pkcs12FromAsn1(
        asn1,
        true,
        (null as unknown) as string
      )
    } catch (error) {
      throw new Error(
        `CI_CERT_BASE64 does not open without a password: ${
          (error as Error).message
        }`
      )
    }
  }
}

/**
 * The private key and SHA-256 thumbprint of the app's certificate, read from the PFX in memory.
 *
 * @param pfxBase64 - The PFX, base64, without a password
 */
export function readCertificate(
  pfxBase64: string
): { privateKey: string; thumbprintSha256: string } {
  const pfx = openPfx(pfxBase64)
  const shroudedOid = forge.pki.oids.pkcs8ShroudedKeyBag
  const keyOid = forge.pki.oids.keyBag
  const certOid = forge.pki.oids.certBag
  const keyBag =
    pfx.getBags({ bagType: shroudedOid })[shroudedOid]?.[0] ??
    pfx.getBags({ bagType: keyOid })[keyOid]?.[0]
  const certificate = pfx.getBags({ bagType: certOid })[certOid]?.[0]?.cert
  if (!keyBag?.key || !certificate)
    throw new Error('CI_CERT_BASE64 holds no private key and certificate')
  const der = forge.asn1
    .toDer(forge.pki.certificateToAsn1(certificate))
    .getBytes()
  return {
    privateKey: forge.pki.privateKeyToPem(keyBag.key),
    thumbprintSha256: forge.md.sha256.create().update(der).digest().toHex()
  }
}

/** One MSAL client per run: it keeps the token and renews it, whichever PnPjs client asks. */
let appClient: ConfidentialClientApplication | undefined

/**
 * A token for SharePoint, as the app (app-only).
 *
 * @param signIn - The app's sign-in
 * @param origin - The tenant's SharePoint origin (`https://<tenant>.sharepoint.com`)
 */
async function appToken(
  signIn: ICertificateSignIn,
  origin: string
): Promise<string> {
  appClient ??= new ConfidentialClientApplication({
    auth: {
      clientId: signIn.clientId,
      authority: `https://login.microsoftonline.com/${signIn.tenant}`,
      clientCertificate: readCertificate(signIn.pfxBase64)
    }
  })
  const result = await appClient.acquireTokenByClientCredential({
    scopes: [`${origin}/.default`]
  })
  if (!result?.accessToken)
    throw new Error(`No token for ${origin} from tenant ${signIn.tenant}`)
  return result.accessToken
}

/**
 * How long one attempt at a request may take. Without a limit a request SharePoint never answered
 * waited for Node's own 300 s timeout, three times over (PnPjs's `BrowserFetchWithRetry` sent it
 * again): one scratch web's set-up took 15 minutes, and the web was left behind.
 */
export const REQUEST_TIMEOUT = 120_000

/**
 * PnPjs's `BrowserFetchWithRetry` with a time limit on each attempt. A throttled request (429, 503,
 * 504) waits what `Retry-After` says and tries again with a fresh limit; a network error is tried
 * again; an attempt not answered in time fails the request, as sending it again could repeat a
 * write SharePoint is still carrying out. A request with a signal of its own keeps it.
 *
 * @param ms - The limit per attempt
 * @param retries - Attempts in all
 * @param onRetry - Told of each retry (throttling, a network error): stderr by default
 */
export function fetchWithRetry(
  ms: number,
  retries = 3,
  onRetry: (what: string) => void = (what) =>
    // A run that throttling slows down says so (no URL: it goes to logs).
    // eslint-disable-next-line no-console
    console.error(`# SharePoint: ${what}, trying again`)
): TimelinePipe<Queryable> {
  return (instance) => {
    instance.on.send.replace(async function (
      url: URL,
      init: any
    ): Promise<Response> {
      let wait = 200
      for (let attempt = 1; ; attempt++) {
        let response: Response
        try {
          response = await fetch(url.toString(), {
            ...init,
            signal: init.signal ?? (AbortSignal as any).timeout(ms)
          })
        } catch (error) {
          if (
            attempt >= retries ||
            /AbortError|TimeoutError/.test(error && error.name)
          )
            throw error
          onRetry(`${init.method || 'GET'} failed (${error && error.message})`)
          continue
        }
        if (
          response.ok ||
          attempt >= retries ||
          ![429, 503, 504].includes(response.status)
        )
          return response
        const after = parseInt(response.headers.get('Retry-After') || '', 10)
        wait = after > 0 ? after * 1000 : wait * 2
        onRetry(
          `${init.method || 'GET'} answered ${response.status}; waiting ${Math.round(
            wait / 1000
          )} s`
        )
        await delay(wait)
      }
    })
    return instance
  }
}

/**
 * Signed in as the app: a bearer token on every request. SharePoint needs no request digest with
 * one.
 *
 * @param baseUrl - The site relative URLs resolve against
 * @param configuration - The configuration
 */
function certificateBehavior(
  baseUrl: string,
  configuration: IEnv
): TimelinePipe<Queryable> {
  const signIn = configuration.certificate
  if (!signIn)
    throw new Error(
      'Certificate sign-in needs CI_TENANT, CI_CLIENT_ID and CI_CERT_BASE64'
    )
  const { origin } = new URL(baseUrl)
  return (instance) => {
    instance.using(
      DefaultHeaders(),
      DefaultInit(),
      fetchWithRetry(REQUEST_TIMEOUT),
      DefaultParse()
    )
    instance.on.auth.replace(async (url: URL, init: any) => {
      init.headers = {
        ...init.headers,
        Authorization: `Bearer ${await appToken(signIn, origin)}`
      }
      return [url, init]
    })
    return resolveAgainst(baseUrl)(instance)
  }
}

/**
 * A cookie from Playwright's storage state.
 */
interface IStoredCookie {
  name: string
  value: string
  domain: string
  expires: number
}

/**
 * The session cookies for a host, as a `Cookie` header, from a Playwright storage state (PP365's
 * `e2e/.auth/user.json`). Throws, naming the file and the renew command, when the `FedAuth`
 * session is missing or has expired.
 *
 * @param file - The storage state file
 * @param host - The SharePoint host
 * @param now - The time to check expiry against (ms)
 */
export function sessionCookies(
  file: string,
  host: string,
  now: number = Date.now()
): string {
  const { cookies } = JSON.parse(fs.readFileSync(file, 'utf8')) as {
    cookies: IStoredCookie[]
  }
  const forHost = cookies.filter(({ domain }) => {
    const bare = domain.replace(/^\./, '')
    return host === bare || host.endsWith(`.${bare}`)
  })
  const session = forHost.find(({ name }) => name === 'FedAuth')
  if (!session || (session.expires > 0 && session.expires * 1000 < now)) {
    throw new Error(
      `The SharePoint session for ${host} is missing or has expired in ${file}: ` +
        'sign in again with `npx playwright test --project=setup` in prosjektportalen365/e2e'
    )
  }
  return forHost.map(({ name, value }) => `${name}=${value}`).join('; ')
}

/**
 * SharePoint's v2.1 APIs (the term store's, which `{sitecollectiontermstoreid}` reads) refuse
 * cookie calls that do not come from a browser. The session is a browser's, so the calls say so.
 */
const BROWSER_AGENT = 'Mozilla/5.0 (compatible; sp-js-provisioning)'

/**
 * Signed in as the user, with the session a Playwright sign-in left; writes carry a request
 * digest, as on a page.
 *
 * @param baseUrl - The site relative URLs resolve against
 * @param configuration - The configuration
 */
function cookieBehavior(
  baseUrl: string,
  configuration: IEnv
): TimelinePipe<Queryable> {
  const cookie = sessionCookies(
    configuration.cookiesFile,
    new URL(baseUrl).hostname
  )
  return (instance) => {
    instance.using(
      DefaultHeaders(),
      DefaultInit(),
      fetchWithRetry(REQUEST_TIMEOUT),
      DefaultParse(),
      RequestDigest()
    )
    instance.on.auth.replace(async (url: URL, init: any) => {
      init.headers = {
        ...init.headers,
        Cookie: cookie,
        'User-Agent': BROWSER_AGENT
      }
      return [url, init]
    })
    return resolveAgainst(baseUrl)(instance)
  }
}

/**
 * The behaviour that connects a PnPjs client to SharePoint in a sign-in mode.
 *
 * @param mode - Sign-in mode
 * @param baseUrl - The site relative URLs resolve against
 * @param configuration - The configuration (from `.env` by default)
 */
export function signedIn(
  mode: AuthMode,
  baseUrl: string,
  configuration: IEnv = env
): TimelinePipe<Queryable> {
  return mode === 'certificate'
    ? certificateBehavior(baseUrl, configuration)
    : cookieBehavior(baseUrl, configuration)
}
