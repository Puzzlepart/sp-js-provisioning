import assert from 'node:assert/strict'
import * as fs from 'fs'
import * as path from 'path'
import { describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import {
  handlerOrder,
  Mode,
  planHandlers,
  projectSteps,
  resolveTemplate,
  transformSchema,
  verifyTargets
} from '../../debug/lib/cli'
import { ISiteConnection } from '../../debug/lib/connect'
import { ROOT, skipReason } from '../../debug/lib/env'
import { describeError, redact } from '../../debug/lib/redact'
import { verify } from '../../debug/lib/verify'
import { DefaultHandlerSort } from '../../src/handlers/exports'
import { withScratchWeb } from './helpers'
import {
  expectedOf,
  IPackageCase,
  PACKAGE_CASES,
  unlisted
} from './packages.expected'

/** The prosjektportalen-hosting checkout: `INTEGRATION_HOSTING_DIR`, else next to this one. */
export const HOSTING_DIR = process.env.INTEGRATION_HOSTING_DIR
  ? path.resolve(process.env.INTEGRATION_HOSTING_DIR)
  : path.resolve(ROOT, '..', 'prosjektportalen-hosting')

/**
 * Why the packages cannot be tested here: no site or sign-in (`skipReason`), or no hosting checkout.
 * Under `INTEGRATION_STRICT=true` (CI) a missing checkout fails instead, as a missing site does.
 */
function packagesSkipReason(): string | undefined {
  const reason = skipReason()
  if (reason) return reason
  if (fs.existsSync(path.join(HOSTING_DIR, 'packages'))) return undefined
  const missing = `no prosjektportalen-hosting checkout at ${path.relative(
    ROOT,
    HOSTING_DIR
  )} (INTEGRATION_HOSTING_DIR)`
  if (process.env.INTEGRATION_STRICT === 'true') throw new Error(missing)
  return missing
}

/**
 * Makes the pages and lists a template's menu points at, where neither the web nor the templates
 * have them: a PP365 project site has its standard pages (`SitePages/Prosjektstatus.aspx`) and
 * lists (`Lists/Prosjektlogg`) from the setup's standard template, a scratch web does not, and
 * SharePoint refuses a menu node whose target is not there. A page or a list one of the templates
 * declares is left for the template to make.
 *
 * @param conn - The scratch web
 * @param schemas - The templates the case applies
 * @returns What was made, as the menu's URLs
 */
async function arrangeMenuTargets(
  conn: ISiteConnection,
  schemas: Array<Record<string, any>>
): Promise<string[]> {
  const urls: string[] = []
  const walk = (nodes: any[]) => {
    for (const node of nodes || []) {
      if (typeof node.Url === 'string') urls.push(node.Url)
      walk(node.Children)
    }
  }
  const declared = new Set<string>()
  for (const schema of schemas) {
    for (const collection of ['QuickLaunch', 'TopNavigationBar'])
      walk(schema.Navigation && schema.Navigation[collection])
    for (const list of schema.Lists || [])
      declared.add(`list ${String(list.Title).toLowerCase()}`)
    for (const page of schema.ClientSidePages || [])
      declared.add(`page ${String(page.Name).toLowerCase()}`)
  }
  if (urls.length === 0) return []
  const web = conn.webInfo.ServerRelativeUrl.replace(/\/$/, '')
  // Through the lists, not by path: a request that names the web's path twice
  // (`getFileByServerRelativePath`) passes SharePoint's URL length limit on a scratch web.
  const lists = await conn.web.lists
    .select('RootFolder/ServerRelativeUrl')
    .expand('RootFolder')<
    Array<{ RootFolder: { ServerRelativeUrl: string } }>
  >()
  const have = new Set(
    lists.map((list) => list.RootFolder.ServerRelativeUrl.toLowerCase())
  )
  const pages = await conn.web.lists.ensureSitePagesLibrary()
  for (const file of await pages.rootFolder.files.select('Name')<
    Array<{ Name: string }>
  >())
    have.add(`${web}/SitePages/${file.Name}`.toLowerCase())
  const made: string[] = []
  for (const url of urls) {
    const page = /^SitePages\/([^/?#]+\.aspx)$/i.exec(url)
    if (page) {
      const file = `${web}/SitePages/${page[1]}`.toLowerCase()
      if (declared.has(`page ${page[1].toLowerCase()}`) || have.has(file))
        continue
      await (await conn.web.addClientsidePage(page[1])).save(true)
      have.add(file)
      made.push(url)
      continue
    }
    const list = /^Lists\/([^/?#]+)\//i.exec(url)
    if (!list) continue
    const name = decodeURIComponent(list[1])
    const target = `${web}/Lists/${name}`.toLowerCase()
    if (declared.has(`list ${name.toLowerCase()}`) || have.has(target)) continue
    await conn.web.lists.add(name)
    have.add(target)
    made.push(url)
  }
  return made
}

/**
 * With `INTEGRATION_PACKAGES=all`: every other package's hub template and project template (with
 * the extensions the setup selects to begin with), reported and not asserted.
 */
function allCases(): IPackageCase[] {
  if (process.env.INTEGRATION_PACKAGES !== 'all') return []
  const listed = (pkg: string, mode: Mode) =>
    PACKAGE_CASES.some((known) => known.pkg === pkg && known.mode === mode)
  return fs
    .readdirSync(path.join(HOSTING_DIR, 'packages'))
    .sort()
    .reduce((cases: IPackageCase[], pkg) => {
      const file = path.join(HOSTING_DIR, 'packages', pkg, 'manifest.json')
      if (!fs.existsSync(file)) return cases
      const provisioning =
        JSON.parse(fs.readFileSync(file, 'utf8')).provisioning || {}
      if (provisioning.hubTemplate && !listed(pkg, 'hub'))
        cases.push({ pkg, mode: 'hub' })
      if (provisioning.template && !listed(pkg, 'project'))
        cases.push({ pkg, mode: 'project' })
      return cases
    }, [])
}

/**
 * The templates a mode applies, in order, and the ones to verify after.
 *
 * @param pkg - The package folder
 * @param mode - The mode
 * @param lang - The locale
 * @param extensions - `--mode project`: the extensions (default: the ones the setup selects)
 */
function stepsOf(
  pkg: string,
  mode: Mode,
  lang?: string,
  extensions?: string[]
) {
  const resolved = resolveTemplate({ package: pkg, mode, lang, extensions })
  const schema = transformSchema(
    resolved.schema,
    { mode, skipTaxonomy: false },
    resolved.manifest
  )
  if (mode !== 'project') {
    const step = { source: resolved.label, schema }
    return { apply: [step], verified: [step] }
  }
  const steps = projectSteps(schema, resolved.extensions ?? [])
  return {
    apply: [...steps.apply, ...steps.hooks],
    verified: verifyTargets(steps.apply)
  }
}

/**
 * The prosjektportalen-hosting packages (docs/specs/integration-testing.md, Fixtures): each case
 * applies a package twice to its own scratch web the way PP365 would in that mode, then verifies
 * it. Until the JSOM exit (`docs/specs/jsom-exit.md`) the runner skips Taxonomy, ContentTypes,
 * Files and PropertyBagEntries in Node, so what depends on them is an explicit list per case.
 * Runs with `npm run test:integration -- packages`, not by default.
 */
describe('Packages', { skip: packagesSkipReason() }, () => {
  const cases = fs.existsSync(path.join(HOSTING_DIR, 'packages'))
    ? [...PACKAGE_CASES, ...allCases()]
    : []
  for (const testCase of cases) {
    const label = [
      testCase.pkg,
      testCase.mode,
      testCase.lang,
      ...(testCase.extensions || [])
    ]
      .filter(Boolean)
      .join(' ')
    describe(label, () => {
      // Short: the web's path is in every request, some name it twice, and SharePoint refuses a
      // URL over its length limit (found with `pkg pp-forskningsprosjekt project forskningsmal`).
      const scratch = withScratchWeb(
        ['pkg', testCase.pkg.replace(/^pp-/, ''), testCase.mode, testCase.lang]
          .filter(Boolean)
          .join(' ')
      )

      it('applies twice with only the expected errors, and verify finds only the expected differences', async (t) => {
        const { conn } = scratch()
        const steps = stepsOf(
          path.join(HOSTING_DIR, 'packages', testCase.pkg),
          testCase.mode,
          testCase.lang,
          testCase.extensions
        )
        const made = await arrangeMenuTargets(
          conn,
          steps.apply.map((step) => step.schema)
        )
        if (made.length)
          t.diagnostic(`made the menu's targets first: ${made.join(', ')}`)
        const handlers = handlerOrder(DefaultHandlerSort)
        const errors: string[] = []
        for (const run of [1, 2])
          for (const step of steps.apply) {
            const result = await apply(conn, step.schema)
            if (result.error)
              assert.fail(
                `run ${run} (${step.source}) stopped in ${
                  result.error.handler
                }: ${describeError(result.error)}`
              )
            for (const entry of result.errors) {
              const line = redact(`[${entry.source ?? '?'}] ${entry.message}`)
              if (!errors.includes(line)) errors.push(line)
            }
          }
        const findings: string[] = []
        for (const step of steps.verified) {
          const result = await verify(conn, step.schema, {
            handlers: planHandlers(step.schema, handlers).run
          })
          for (const finding of result.findings)
            findings.push(
              redact(
                `[${finding.handler}] ${
                  finding.path
                }: expected ${JSON.stringify(
                  finding.expected
                )}, found ${JSON.stringify(finding.actual)}`
              )
            )
        }
        if (!testCase.lists) {
          // INTEGRATION_PACKAGES=all: reported, the counts first.
          t.diagnostic(
            `${label}: ${errors.length} error(s), ${findings.length} finding(s)`
          )
          for (const line of [...errors, ...findings]) t.diagnostic(line)
          return
        }
        const expected = expectedOf(testCase)
        // One assertion for both, named after the case: every case shares this test's name.
        assert.deepEqual(
          {
            case: label,
            errors: unlisted(errors, expected.errors),
            findings: unlisted(findings, expected.findings)
          },
          {
            case: label,
            errors: { unexpected: [], missing: [] },
            findings: { unexpected: [], missing: [] }
          },
          `${label}: not what the list says the Node run cannot do (a missing entry was fixed: remove it)`
        )
      })
    })
  }
})
