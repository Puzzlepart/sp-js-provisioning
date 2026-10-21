import { config } from 'dotenv'
import * as fs from 'fs'
import * as path from 'path'

/**
 * The repository root, where `.env` lives.
 */
export const ROOT = path.resolve(__dirname, '..', '..')

config({ path: path.join(ROOT, '.env'), quiet: true })

/**
 * How the runner signs in: as an app with its certificate (app-only, CI), or as PP365's e2e user
 * with the session cookies its Playwright sign-in left (local).
 */
export type AuthMode = 'certificate' | 'cookies'

/**
 * An app's sign-in: the tenant, the app's client id and its certificate (a PFX, base64, with no
 * password). Same variables as PP365's integration tests, so one secret serves both repositories.
 */
export interface ICertificateSignIn {
  tenant: string
  clientId: string
  pfxBase64: string
}

/**
 * The configuration, from the environment and the repository's `.env`.
 */
export interface IEnv {
  /** The site reserved for this repository's runs and tests, without a trailing slash. */
  siteUrl?: string
  /** `INTEGRATION_AUTH`, when it forces a sign-in mode. */
  auth?: AuthMode
  certificate?: ICertificateSignIn
  /** Playwright storage state with SharePoint session cookies (PP365's e2e user). */
  cookiesFile: string
}

/**
 * A variable's value, or `undefined` when it is empty or still the `<placeholder>` of
 * `.env.example`.
 *
 * @param name - Variable
 * @param variables - Where to read it from
 */
export function configured(
  name: string,
  variables: Record<string, string | undefined> = process.env
): string | undefined {
  const value = variables[name]?.trim()
  return value && !/<[^>]+>/.test(value) ? value : undefined
}

/**
 * Reads the configuration.
 *
 * @param variables - Where to read it from (the environment by default)
 */
export function readEnv(
  variables: Record<string, string | undefined> = process.env
): IEnv {
  const value = (name: string) => configured(name, variables)
  const forced = value('INTEGRATION_AUTH')
  const tenant = value('CI_TENANT')
  const clientId = value('CI_CLIENT_ID')
  const pfxBase64 = value('CI_CERT_BASE64')
  return {
    siteUrl: value('INTEGRATION_SITE_URL')?.replace(/\/+$/, ''),
    auth: forced === 'certificate' || forced === 'cookies' ? forced : undefined,
    certificate:
      tenant && clientId && pfxBase64
        ? { tenant, clientId, pfxBase64 }
        : undefined,
    cookiesFile: path.resolve(
      ROOT,
      value('INTEGRATION_COOKIES_FILE') ??
        path.join('..', 'prosjektportalen365', 'e2e', '.auth', 'user.json')
    )
  }
}

/**
 * The configuration of this process.
 */
export const env = readEnv()

/**
 * The sign-in to use: the one `INTEGRATION_AUTH` names, else the certificate when it is
 * configured, else the cookies when their file exists; `undefined` when neither is there.
 *
 * @param configuration - The configuration
 * @param exists - Whether a file exists
 */
export function authMode(
  configuration: IEnv = env,
  exists: (file: string) => boolean = fs.existsSync
): AuthMode | undefined {
  if (configuration.auth) return configuration.auth
  if (configuration.certificate) return 'certificate'
  if (exists(configuration.cookiesFile)) return 'cookies'
  return undefined
}
