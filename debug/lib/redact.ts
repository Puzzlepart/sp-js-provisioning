import * as fs from 'fs'
import * as path from 'path'

/**
 * One rule for everything a run leaves where others can read it: the job log, the JUnit and HTML
 * reports, the artifact and the test report site (docs/specs/integration-testing.md, D5). The
 * repository is public, and so are its job logs and artifacts.
 *
 * `redact()` replaces, in this order:
 *
 * - secrets: session cookie values, any `Set-Cookie`, bearer tokens, PEM blocks, certificates and
 *   the value of `CI_CERT_BASE64`, with `<redacted>`;
 * - what the run was configured with (`INTEGRATION_USER`, `CI_CLIENT_ID`, `CI_TENANT`, the path of
 *   `INTEGRATION_SITE_URL`), with `<user>`, `<client-id>`, `<tenant>` and `/sites/<site>`;
 * - any other e-mail address or user principal name (also URL-encoded, and as OneDrive writes it
 *   in a path), with `<user>`;
 * - SharePoint hosts and `onmicrosoft.com` domains, with `<tenant>`;
 * - the site in a `/sites/...` or `/teams/...` path, with `<site>` (what is below it stays: the
 *   scratch web's name says which suite it was);
 * - every GUID that is not one of the repository's own (fixtures, tests, the engine, the hosting
 *   packages), with `<guid>`: ids of the tenant, the app, sites, webs, lists and requests.
 *
 * The documentation's example hosts and domains (contoso, fabrikam, example.org) are kept, with
 * their paths. A text that has been redacted is not changed by redacting it again, which is what
 * `unredacted()` and `npm run check:reports` rely on.
 */

/** The placeholders `redact()` writes, as `<name>`. */
export const PLACEHOLDERS = [
  'redacted',
  'user',
  'client-id',
  'tenant',
  'site',
  'guid'
]

/** Tenant names of documentation: never a real tenant. */
const EXAMPLE_TENANT = /^(contoso|fabrikam)$/i

/** Mail domains of documentation and of the repository's host: never a tenant's user. */
const EXAMPLE_DOMAIN = /(^|\.)(contoso\.com|fabrikam\.com|(contoso|fabrikam)\.onmicrosoft\.com|example\.(com|org|net)|github\.com)$/i

const GUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi

/** File types read for the repository's own GUIDs. */
const TEXT_FILE = /\.(ts|js|json|md|xml|txt)$/i

let ownGuids: Set<string> | undefined

/**
 * The GUIDs in the repository's own files: field, content type and component ids of the fixtures,
 * the tests and the engine, and of the hosting packages when their checkout is there
 * (`INTEGRATION_HOSTING_DIR`, else next to this one). Anyone can read them on GitHub. Read once
 * per process, the first time a text holds a GUID.
 */
function repositoryGuids(): Set<string> {
  if (ownGuids) return ownGuids
  const found = new Set<string>()
  const root = path.resolve(__dirname, '..', '..')
  const hosting = process.env.INTEGRATION_HOSTING_DIR
    ? path.resolve(process.env.INTEGRATION_HOSTING_DIR)
    : path.resolve(root, '..', 'prosjektportalen-hosting')
  const scan = (dir: string, depth: number) => {
    let entries: string[]
    try {
      entries = fs.readdirSync(dir)
    } catch {
      return
    }
    for (const name of entries) {
      if (name === 'node_modules' || name.startsWith('.')) continue
      const file = path.join(dir, name)
      let stat: fs.Stats
      try {
        stat = fs.statSync(file)
      } catch {
        continue
      }
      if (stat.isDirectory()) {
        if (depth < 8) scan(file, depth + 1)
      } else if (TEXT_FILE.test(name) && stat.size < 8 * 1024 * 1024) {
        const matches = fs.readFileSync(file, 'utf8').match(GUID)
        for (const guid of matches || []) found.add(guid.toLowerCase())
      }
    }
  }
  for (const dir of [
    path.join(root, 'tests'),
    path.join(root, 'src'),
    path.join(hosting, 'packages')
  ])
    scan(dir, 0)
  ownGuids = found
  return found
}

/** A text as a regular expression that matches it literally. */
const literal = (text: string) => text.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')

/**
 * The replacements for what the run was configured with, from the environment: `[pattern,
 * replacement]` pairs. Built again only when the environment changes.
 */
let configured: {
  key: string
  rules: Array<[RegExp, string]>
  certificate?: string
}

function configuredRules(): typeof configured {
  const {
    INTEGRATION_USER: user,
    CI_CLIENT_ID: clientId,
    CI_TENANT: tenant,
    INTEGRATION_SITE_URL: siteUrl,
    CI_CERT_BASE64: certificate
  } = process.env
  const key = [
    user,
    clientId,
    tenant,
    siteUrl,
    certificate && certificate.length
  ].join('\n')
  if (configured && configured.key === key) return configured
  const rules: Array<[RegExp, string]> = []
  const real = (value?: string) =>
    !!value && value.trim().length > 3 && !/^<.*>$/.test(value.trim())
  if (real(user) && !EXAMPLE_DOMAIN.test(user.split('@')[1] || ''))
    rules.push([new RegExp(literal(user.trim()), 'gi'), '<user>'])
  if (real(clientId))
    rules.push([new RegExp(literal(clientId.trim()), 'gi'), '<client-id>'])
  if (real(tenant)) {
    const name = /^([a-z0-9-]+)\.onmicrosoft\.com$/i.exec(tenant.trim())
    // A `<name>.onmicrosoft.com` tenant is replaced by the domain rule, which keeps the suffix.
    if (!name)
      rules.push([new RegExp(literal(tenant.trim()), 'gi'), '<tenant>'])
  }
  if (real(siteUrl))
    try {
      const segments = new URL(siteUrl.trim()).pathname
        .split('/')
        .filter(Boolean)
      if (segments.length >= 2)
        rules.push([
          new RegExp(
            `/${literal(segments[0])}/${literal(
              segments[1]
            )}(?=[/?#"'\\s<>&),;:.\\]]|$)`,
            'gi'
          ),
          `/${segments[0]}/<site>`
        ])
    } catch {
      // Not a URL: nothing to take a path from.
    }
  configured = {
    key,
    rules,
    ...(certificate && certificate.length > 40 ? { certificate } : {})
  }
  return configured
}

/**
 * Text that may reach a terminal, a test report or a CI artifact, without secrets, the tenant, its
 * users or its ids (see the module).
 *
 * @param text - The text
 */
export function redact(text: string): string {
  const { rules, certificate } = configuredRules()
  let out = String(text)
  if (certificate && out.includes(certificate))
    out = out.split(certificate).join('<redacted>')
  out = out
    .replace(
      /\b(FedAuth|rtFa|SIMI|EdgeAccessCookie)=[^;\s"',]*/gi,
      '$1=<redacted>'
    )
    .replace(/(set-cookie['"]?\s*[:=]\s*['"]?)[^'"\n]*/gi, '$1<redacted>')
    .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/g, 'Bearer <redacted>')
    // A PEM block on one line, then line by line (a log is relayed a line at a time, and a lone
    // marker must not swallow the rest of a report): its markers, and every line that is nothing
    // but base64, which is also a secret someone wrapped.
    .replace(/-----BEGIN [A-Z ]+-----.*?-----END [A-Z ]+-----/g, '<redacted>')
    .replace(/-----(?:BEGIN|END) [A-Z ]+-----/g, '<redacted>')
    .replace(/\bMII[A-Za-z0-9+/]{60,}={0,2}/g, '<redacted>')
    .replace(/^[A-Za-z0-9+/]{60,}={0,2}$/gm, '<redacted>')
  for (const [pattern, replacement] of rules)
    out = out.replace(pattern, replacement)
  return out
    .replace(
      /[A-Za-z0-9._%+-]{1,64}(?:@|%40)[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}/g,
      (match) =>
        EXAMPLE_DOMAIN.test(match.split(/@|%40/).pop()) ? match : '<user>'
    )
    .replace(
      /\/personal\/[a-z0-9]+(?:_[a-z0-9]+)*_onmicrosoft_com/gi,
      '/personal/<user>'
    )
    .replace(/\b([a-z0-9-]{1,63})\.onmicrosoft\.com/gi, (match, name) =>
      EXAMPLE_TENANT.test(name) ? match : '<tenant>.onmicrosoft.com'
    )
    .replace(
      /\b([a-z0-9]+(?:-(?!(?:admin|my)\.sharepoint\.com)[a-z0-9]+)*)((?:-admin|-my)?)\.sharepoint\.com/gi,
      (match, host, suffix) =>
        EXAMPLE_TENANT.test(host) ? match : `<tenant>${suffix}.sharepoint.com`
    )
    .replace(
      /((?:\b[a-z0-9-]+)?\.sharepoint\.com)?\/(sites|teams)\/[^/\s"'<>&?#,;:)\]]+/gi,
      (match, host, managed) =>
        host &&
        EXAMPLE_TENANT.test(host.split('.')[0].replace(/-(admin|my)$/i, ''))
          ? match
          : `${host || ''}/${managed}/<site>`
    )
    .replace(GUID, (match) =>
      repositoryGuids().has(match.toLowerCase()) ? match : '<guid>'
    )
}

/**
 * `redact()` for XML or HTML text: the placeholders it writes, escaped (`<tenant>` in an attribute
 * would make the XML invalid).
 *
 * @param xml - The XML
 */
export function redactXml(xml: string): string {
  return redact(xml).replace(
    new RegExp(`<(${PLACEHOLDERS.join('|')})>`, 'g'),
    '&lt;$1&gt;'
  )
}

/**
 * What `redact()` would still change in a text: for each such line, its number and the kinds of
 * value it holds (the placeholders `redact()` would write), never the values. Empty for a text
 * that is fully redacted, which every report must be before it is uploaded or published.
 *
 * @param text - The text of a report or a log
 */
export function unredacted(
  text: string
): Array<{ line: number; kinds: string[] }> {
  const count = (line: string, kind: string) =>
    line.split(`<${kind}>`).length - 1
  // A report escapes the placeholders (`&lt;tenant&gt;`): read them as redact() wrote them.
  const escaped = new RegExp(`&lt;(${PLACEHOLDERS.join('|')})&gt;`, 'g')
  const found: Array<{ line: number; kinds: string[] }> = []
  String(text)
    .split('\n')
    .forEach((raw, index) => {
      const line = raw.replace(escaped, '<$1>')
      const redacted = redact(line)
      if (redacted === line) return
      const kinds = PLACEHOLDERS.filter(
        (kind) => count(redacted, kind) > count(line, kind)
      )
      found.push({
        line: index + 1,
        kinds: kinds.length ? kinds : ['redacted']
      })
    })
  return found
}

/**
 * The message of an error, redacted: never the error object itself, whose `cause.response` holds
 * the response headers (with any `Set-Cookie`).
 *
 * @param error - The error
 */
export function describeError(error: any): string {
  if (error === undefined || error === null) return String(error)
  const status = error.status ?? (error.cause && error.cause.status)
  const message =
    typeof error === 'string' ? error : error.message || String(error)
  return redact(
    status && !message.includes(`[${status}]`)
      ? `${message} (${status})`
      : message
  )
}
