/* eslint-disable no-console */
/**
 * `npm run check:reports -- <folder>`: fails when a report in the folder still holds something
 * `redact()` would replace (`debug/lib/redact.ts`: a tenant host, a site path, a user, a tenant
 * id, a cookie, a token, a certificate). The integration workflow runs it before it uploads the
 * reports as an artifact and publishes them on the test report site: a hit stops both. It prints
 * the file, the line and the kind of value, never the value.
 *
 * It is for the reports of an integration run. The unit tests' JUnit file (`npm run test:ci`:
 * `test-results/unit.xml`) is written by Node without `redact()`; the unit tests never reach a
 * tenant, and the example paths in their names (`/sites/demo`) would be reported here.
 *
 * Exit codes: 0 nothing found (or no such folder: nothing to publish), 1 a finding, 2 usage.
 */
import * as fs from 'fs'
import * as path from 'path'
// Loads `.env`, so a local check knows the site and the user the run was configured with, as
// redact() did when the reports were written (in CI they are in the environment).
import '../../debug/lib/env'
import { unredacted } from '../../debug/lib/redact'

/** The report files that are text: JUnit, the HTML page, logs. */
const REPORT_FILE = /\.(xml|html?|txt|json|log|md)$/i

/**
 * The text files below a folder.
 *
 * @param dir - The folder
 */
export function reportFiles(dir: string): string[] {
  return fs
    .readdirSync(dir)
    .sort()
    .reduce((files: string[], name) => {
      const file = path.join(dir, name)
      if (fs.statSync(file).isDirectory())
        return files.concat(reportFiles(file))
      return REPORT_FILE.test(name) ? files.concat(file) : files
    }, [])
}

/**
 * What the reports in a folder still hold that `redact()` would replace, as
 * `<file>:<line>: <kinds>` lines, with the number of files read.
 *
 * @param dir - The folder
 */
export function checkReports(
  dir: string
): { files: number; findings: string[] } {
  const files = reportFiles(dir)
  const findings: string[] = []
  for (const file of files)
    for (const found of unredacted(fs.readFileSync(file, 'utf8')))
      findings.push(
        `${path.relative(dir, file)}:${found.line}: ${found.kinds
          .map((kind) => `<${kind}>`)
          .join(', ')}`
      )
  return { files: files.length, findings }
}

if (require.main === module) {
  const dir = process.argv[2]
  if (!dir) {
    console.error('Usage: npm run check:reports -- <folder>')
    process.exitCode = 2
  } else if (!fs.existsSync(dir)) {
    console.log(`check-reports: no ${dir}, nothing to check`)
  } else {
    const { files, findings } = checkReports(dir)
    if (findings.length) {
      console.error(
        `check-reports: ${findings.length} line(s) in ${dir} hold what redact() would replace ` +
          `(file:line: kinds). The reports are not uploaded or published:\n  ${findings
            .slice(0, 50)
            .join('\n  ')}${findings.length > 50 ? '\n  ...' : ''}`
      )
      process.exitCode = 1
    } else console.log(`check-reports: OK, ${files} file(s) in ${dir}`)
  }
}
