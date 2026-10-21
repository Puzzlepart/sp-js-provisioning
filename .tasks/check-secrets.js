/**
 * Fails when a committed file that tests or the runner read holds a tenant URL, a session cookie,
 * a bearer token or a certificate. Example hosts (contoso, fabrikam, also with -admin or -my) are
 * allowed.
 *
 * Usage: node .tasks/check-secrets.js
 */
const { execFileSync } = require('node:child_process')
const { readFileSync } = require('node:fs')
const { join } = require('node:path')

const EXAMPLE_HOSTS = ['contoso', 'fabrikam']

/**
 * Whether a line names a SharePoint host that is not an example host.
 *
 * @param {string} line
 */
function hasTenantUrl(line) {
  return (line.match(/[a-z0-9-]+(?=\.sharepoint\.com)/gi) || []).some(
    (host) => !EXAMPLE_HOSTS.includes(host.toLowerCase().replace(/-(admin|my)$/, ''))
  )
}

const PATTERNS = [
  { name: 'a tenant URL', test: hasTenantUrl },
  // Real session cookies are long; tests use short stand-ins such as `FedAuth=session`.
  { name: 'a session cookie', test: (line) => /(FedAuth|rtFa)=[A-Za-z0-9+/=%_.-]{40,}/.test(line) },
  { name: 'a bearer token', test: (line) => /Bearer ey[A-Za-z0-9_-]{10,}/.test(line) },
  { name: 'a certificate', test: (line) => /MII[A-Za-z0-9+/]{40,}/.test(line) }
]

/**
 * What a file's text leaks, as `<line>: <what>` entries.
 *
 * @param {string} text
 */
function findingsIn(text) {
  const findings = []
  text.split('\n').forEach((line, index) => {
    for (const pattern of PATTERNS)
      if (pattern.test(line)) findings.push(`${index + 1}: ${pattern.name}`)
  })
  return findings
}

module.exports = { findingsIn, hasTenantUrl }

if (require.main === module) {
  const root = join(__dirname, '..')
  // Committed files and new ones not yet added (but not ignored ones such as .env).
  const files = execFileSync(
    'git',
    ['ls-files', '-z', '--cached', '--others', '--exclude-standard', '--', 'tests', 'debug', '.env.example'],
    { cwd: root, encoding: 'utf8' }
  )
    .split('\0')
    .filter(Boolean)
  const findings = files.flatMap((file) =>
    findingsIn(readFileSync(join(root, file), 'utf8')).map((finding) => `${file}:${finding}`)
  )
  if (findings.length) {
    console.error(`check-secrets: ${findings.length} finding(s):\n  ${findings.join('\n  ')}`)
    process.exit(1)
  }
  console.log(`check-secrets: OK, ${files.length} files`)
}
