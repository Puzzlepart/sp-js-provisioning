/* eslint-disable no-console */
/**
 * Runs the integration suites against the reserved site (`INTEGRATION_SITE_URL`), one file at a
 * time: `reachability.test.ts` first, in a process of its own (`node --test` runs the files it is
 * given in alphabetical order), then the suites asked for. When reachability fails, the suites do
 * not run: they would all fail the same way.
 *
 *   npm run test:integration                         every suite
 *   npm run test:integration -- Lists Security       those suites (handler names, any case)
 *   npm run test:integration -- reachability         only the reachability checks
 *   npm run check:tenant                             the same, strict: a missing site fails
 *   npm run test:integration -- Lists --keep         keep the scratch webs for inspection
 *   npm run test:integration:ci                      strict, spec + JUnit to test-results/
 *
 * Other `--test-*` flags (e.g. `--test-name-pattern=...`) are passed to `node --test` for the
 * suites (not reachability). A pattern that skips the first test of Lists or Lists (DataRows)
 * fails the others: they build on its lists. `--strict` sets `INTEGRATION_STRICT=true`.
 * With `INTEGRATION_SCRATCH=run` the suites share one scratch web, created and deleted here.
 */
import { spawn, spawnSync } from 'child_process'
import * as fs from 'fs'
import { mkdirSync } from 'node:fs'
import * as path from 'path'
import { connect, ISiteConnection } from '../../debug/lib/connect'
import { env, ROOT, skipReason } from '../../debug/lib/env'
import { describeError, redact } from '../../debug/lib/redact'
import {
  createScratchWeb,
  deleteScratchWeb,
  IScratchWeb
} from '../../debug/lib/scratch'

const SUITES = path.join(__dirname)

/**
 * The suites there are, by name: `lists` for `lists.test.ts`.
 */
function availableSuites(): string[] {
  return fs
    .readdirSync(SUITES)
    .filter((file) => file.endsWith('.test.ts'))
    .map((file) => file.replace(/\.test\.ts$/, ''))
    .sort()
}

/**
 * The suite files to run, reachability first.
 *
 * @param names - Suite or handler names; none for every suite
 * @param available - The suites there are
 */
export function suiteFiles(names: string[], available: string[]): string[] {
  // A handler name selects its suites: `lists` runs `lists` and `lists-datarows`.
  const selected = (name: string) =>
    available.filter((suite) => suite === name || suite.startsWith(`${name}-`))
  const unknown = names.filter((name) => !selected(name.toLowerCase()).length)
  if (unknown.length)
    throw new Error(
      `Unknown suite(s): ${unknown
        .map((name) => name.toLowerCase())
        .join(', ')}. Suites: ${available.join(', ')}`
    )
  const wanted = names.length
    ? names
        .map((name) => name.toLowerCase())
        .reduce((all, name) => [...all, ...selected(name)], [] as string[])
        .filter((name, index, all) => all.indexOf(name) === index)
    : available.filter((name) => name !== 'reachability')
  const ordered = [
    'reachability',
    ...wanted.filter((name) => name !== 'reachability')
  ]
  return ordered.map((name) => path.join(SUITES, `${name}.test.ts`))
}

/**
 * Copies a child's output line by line through `redact()`.
 *
 * @param from - The child's stream
 * @param to - Where it goes
 */
function relayRedacted(from: any, to: { write(text: string): any }): void {
  let rest = ''
  from.setEncoding('utf8')
  from.on('data', (chunk: string) => {
    const lines = (rest + chunk).split('\n')
    rest = lines.pop() || ''
    for (const line of lines) to.write(`${redact(line)}\n`)
  })
  from.on('end', () => {
    if (rest) to.write(redact(rest))
  })
}

/**
 * Redacts a JUnit report in place, if there is one. The placeholders are escaped: `<tenant>` in
 * an attribute would make the XML invalid.
 *
 * @param file - The report
 */
export function redactReport(file: string): void {
  if (!fs.existsSync(file)) return
  const xml = fs.readFileSync(file, 'utf8')
  fs.writeFileSync(file, redactXml(xml))
}

/**
 * `redact()` for XML text: the placeholders it writes, escaped.
 *
 * @param xml - The XML
 */
export function redactXml(xml: string): string {
  return redact(xml).replace(/<(tenant|redacted)>/g, '&lt;$1&gt;')
}

async function main(): Promise<number> {
  const argv = process.argv.slice(2)
  const ci = argv.includes('--ci')
  const strict = ci || argv.includes('--strict')
  // On a CI runner (CI=true, as GitHub Actions sets it) also without --ci: check:tenant runs there.
  const redacting = ci || process.env.CI === 'true'
  const keep = argv.includes('--keep') || env.keep
  const passThrough = argv.filter((arg) => arg.startsWith('--test-'))
  const names = argv.filter((arg) => !arg.startsWith('--'))
  const files = suiteFiles(names, availableSuites())

  const childEnv: Record<string, string> = { ...process.env } as any
  if (keep) childEnv.INTEGRATION_KEEP = 'true'
  if (strict) childEnv.INTEGRATION_STRICT = 'true'
  const reporters = (junit: string) =>
    ci
      ? [
          '--test-reporter=spec',
          '--test-reporter-destination=stdout',
          '--test-reporter=junit',
          `--test-reporter-destination=test-results/${junit}`
        ]
      : ['--test-reporter=spec']
  if (ci) mkdirSync(path.join(ROOT, 'test-results'), { recursive: true })
  const tsxCli = require.resolve('tsx/cli')
  const run = async (
    testFiles: string[],
    junit: string,
    extra: string[]
  ): Promise<number> => {
    const args = [
      tsxCli,
      '--tsconfig',
      path.join(ROOT, 'tests', 'tsconfig.json'),
      '--test',
      '--test-concurrency=1',
      ...reporters(junit),
      ...extra,
      ...testFiles
    ]
    if (!redacting) {
      const result = spawnSync(process.execPath, args, {
        cwd: ROOT,
        env: childEnv,
        stdio: 'inherit'
      })
      return result.status === null ? 1 : result.status
    }
    // CI keeps the log and the reports: a SharePoint error message names the site's absolute URL,
    // so both pass through redact() (the spec: no tenant URL in a log or artifact).
    const child = spawn(process.execPath, args, {
      cwd: ROOT,
      env: childEnv,
      stdio: ['ignore', 'pipe', 'pipe']
    })
    relayRedacted(child.stdout, process.stdout)
    relayRedacted(child.stderr, process.stderr)
    const status = await new Promise<number>((resolve) =>
      child.on('close', (code: number | null) =>
        resolve(code === null ? 1 : code)
      )
    )
    redactReport(path.join(ROOT, 'test-results', junit))
    return status
  }

  // Whole, and without the pass-through flags: a name pattern would filter its checks away.
  const [reachability, ...suites] = files
  const reached = await run([reachability], 'reachability.xml', [])
  if (reached !== 0) {
    console.error('reachability failed: the suites were not run.')
    return reached
  }
  if (!suites.length) return 0

  // One scratch web for the whole run, when asked for and the site is reachable.
  let shared: IScratchWeb | undefined
  let parent: ISiteConnection | undefined
  if (env.scratch === 'run' && !skipReason({ ...env, strict: false })) {
    parent = await connect(env.siteUrl)
    shared = await createScratchWeb(parent, 'run')
    childEnv.INTEGRATION_SCRATCH_WEB = shared.url
    childEnv.INTEGRATION_SCRATCH_STAMP = shared.stamp
    childEnv.INTEGRATION_SCRATCH_TITLE = shared.title
    console.log(`Shared scratch web: ${shared.conn.webInfo.ServerRelativeUrl}`)
  }

  const status = await run(suites, 'integration.xml', passThrough)
  if (shared && parent) {
    if (keep)
      console.log(
        `Kept shared scratch web: ${shared.conn.webInfo.ServerRelativeUrl}`
      )
    else await deleteScratchWeb(parent, shared)
  }
  return status
}

if (require.main === module)
  main()
    .then((code) => {
      process.exitCode = code
    })
    .catch((error) => {
      console.error(describeError(error))
      process.exitCode = 1
    })
