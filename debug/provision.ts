/* eslint-disable no-console */
/**
 * Local provisioning runner: apply a template or a Prosjektportalen package to a real SharePoint
 * site straight from Node, without the SPFx build/deploy loop.
 *
 * It reproduces what PP365's PackageInstaller does in the browser
 * (`WebProvisioner(web).setup({ spfxContext, logging }).applyTemplate(schema)`), signed in with
 * the configuration in `.env` (see `.env.example`): the session cookies of a Playwright sign-in
 * locally, or an app's certificate in CI. It runs the engine source (`src/`) through tsx, so a
 * handler change takes effect on the next run without a build.
 *
 * Usage:
 *   npm run provision -- --package <path> [--site <url>] [options]
 *   npm run provision -- --help
 */
import {
  describeSchema,
  handlerOrder,
  IArgs,
  IStep,
  parseArgs,
  planHandlers,
  projectSteps,
  resolveTemplate,
  transformSchema,
  usage,
  UsageError,
  verifyTargets
} from './lib/cli'
import { DefaultHandlerSort } from '../src/handlers/exports'
import { apply, IApplyResult } from './lib/apply'
import {
  accessDenied,
  assertConnectedWritable,
  assertWritable,
  connect,
  ISiteConnection
} from './lib/connect'
import { authMode, env } from './lib/env'
import { describeError } from './lib/redact'
import { formatResult, IFinding, verify } from './lib/verify'

/**
 * A template the run applies, with the handlers it runs and skips.
 */
interface IPlannedStep extends IStep {
  plan: { run: string[]; skipped: string[] }
}

/**
 * Prints the plan for a run.
 *
 * @param site - The target site
 * @param label - What the template came from
 * @param args - The arguments
 * @param steps - The templates to apply, in order, with their handlers
 */
function printPlan(
  site: string,
  label: string,
  args: IArgs,
  steps: IPlannedStep[]
): void {
  console.log(`\nSite:     ${site || '(not set)'}`)
  console.log(`Package:  ${label}`)
  console.log(
    `Mode:     ${args.mode}${args.skipTaxonomy ? ' (skip taxonomy)' : ''}${
      args.twice ? ', twice' : ''
    }${args.strict ? ', strict' : ''}${
      args.verify ? `, verify${args.exact ? ' (exact)' : ''}` : ''
    }`
  )
  for (const step of steps) {
    if (steps.length > 1) console.log(`\nStep:     ${step.source}`)
    console.log(`Template sections:\n${describeSchema(step.schema)}`)
    console.log(`Will run: ${step.plan.run.join(', ') || '(none)'}`)
    if (step.plan.skipped.length)
      console.log(
        `Skipped:  ${step.plan.skipped.join(
          ', '
        )} (browser/SPFx-only in Node; pass --handlers to force)`
      )
  }
  console.log('')
}

/**
 * The templates to apply, in order: the template (after the hub/cloud/taxonomy transforms), or in
 * `--mode project` the project setup's `ApplyTemplate` and `Hooks` steps.
 *
 * @param args - The arguments
 * @param handlers - The engine's handlers, in run order
 */
function plannedSteps(
  args: IArgs,
  handlers: string[]
): { label: string; steps: IPlannedStep[]; verified: IPlannedStep[] } {
  const resolved = resolveTemplate(args)
  const schema = transformSchema(resolved.schema, args, resolved.manifest)
  const planned = (step: IStep): IPlannedStep => ({
    ...step,
    plan: planHandlers(step.schema, handlers, args.handlers)
  })
  if (args.mode !== 'project') {
    const step = planned({ source: resolved.label, schema })
    return { label: resolved.label, steps: [step], verified: [step] }
  }
  const project = projectSteps(schema, resolved.extensions ?? [])
  const steps = [...project.apply, ...project.hooks].map(planned)
  return {
    label: resolved.label,
    steps,
    verified: verifyTargets(project.apply).map(planned)
  }
}

/**
 * Applies the steps once, printing the engine's log lines (with their data, e.g. the SharePoint
 * error body). Stops at a step that threw, as the project setup does.
 *
 * @param connection - The web
 * @param steps - The templates to apply
 * @param args - The arguments
 * @param attempt - Which run (with `--twice`)
 */
async function runSteps(
  connection: ISiteConnection,
  steps: IPlannedStep[],
  args: IArgs,
  attempt: number
): Promise<Array<{ step: IPlannedStep; result: IApplyResult }>> {
  const results: Array<{ step: IPlannedStep; result: IApplyResult }> = []
  for (const step of steps) {
    if (steps.length > 1) console.log(`\n--- ${step.source} ---`)
    const result = await apply(connection, step.schema, {
      handlers: step.plan.run,
      parameters: args.parameters,
      prefix: `(provision${args.twice ? ` #${attempt}` : ''})`,
      onHandler: (handler) => console.log(`\n> ${handler}`)
    })
    results.push({ step, result })
    if (result.error) break
  }
  return results
}

/**
 * Verifies the applied templates against the web and prints the result.
 *
 * @param connection - The web
 * @param verified - The templates to verify, with the handlers that ran
 * @param args - The arguments
 * @returns The findings
 */
async function verifySteps(
  connection: ISiteConnection,
  verified: IPlannedStep[],
  args: IArgs
): Promise<IFinding[]> {
  console.log('\n=== Verify ===')
  const findings: IFinding[] = []
  for (const step of verified) {
    const result = await verify(connection, step.schema, {
      exact: args.exact,
      handlers: step.plan.run,
      parameters: args.parameters
    })
    if (verified.length > 1)
      result.findings.forEach((finding) => (finding.source = step.source))
    if (verified.length > 1) console.log(`\n${step.source}:`)
    for (const line of formatResult(result)) console.log(`   ${line}`)
    findings.push(...result.findings)
  }
  return findings
}

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2))
  const site = (args.site ?? env.siteUrl ?? '').replace(/\/+$/, '')
  const handlers = handlerOrder(DefaultHandlerSort)
  if (args.help) {
    console.log(usage(handlers))
    return 0
  }
  const { label, steps, verified } = plannedSteps(args, handlers)
  printPlan(site, label, args, steps)
  if (args.dryRun) {
    console.log('Dry run: not signing in, not applying.')
    return 0
  }
  if (!site)
    throw new UsageError('Missing --site (or INTEGRATION_SITE_URL in .env)')
  assertWritable(site, args.allowSite)

  let connection: ISiteConnection
  try {
    connection = await connect(site)
  } catch (error) {
    const message = accessDenied(error, site, authMode())
    if (!message) throw error
    console.error(`\n${message}`)
    return 1
  }
  // The web SharePoint reports, not the URL given, decides (a `..` in --site is resolved there).
  assertConnectedWritable(connection, args.allowSite)
  console.log(
    `Signed in with ${connection.mode} to "${connection.webInfo.Title}" (${connection.webInfo.Url})\n`
  )

  const runs: Array<Array<{ step: IPlannedStep; result: IApplyResult }>> = []
  for (let attempt = 1; attempt <= (args.twice ? 2 : 1); attempt++) {
    if (args.twice) console.log(`\n=== Run ${attempt} of 2 ===`)
    const run = await runSteps(connection, steps, args, attempt)
    runs.push(run)
    if (run.some(({ result }) => result.error)) break
  }

  let failed = false
  let stopped = false
  runs.forEach((run, index) => {
    const name = args.twice ? `Run ${index + 1}` : 'Run'
    for (const { step, result } of run) {
      const where = steps.length > 1 ? `${name} (${step.source})` : name
      if (result.error) {
        failed = true
        stopped = true
        console.error(`\nFAILED ${where}: provisioning stopped.`)
        if (result.error.handler)
          console.error(`   Failing handler: ${result.error.handler}`)
        console.error(`   Message: ${describeError(result.error)}`)
        continue
      }
      const logged = `${result.errors.length} error(s), ${result.warnings.length} warning(s) logged`
      if (result.errors.length) {
        if (args.strict) failed = true
        console.log(
          `\n${
            args.strict ? 'FAILED' : 'DONE with errors'
          } ${where}: ${logged}:`
        )
        for (const entry of result.errors)
          console.log(`   [${entry.source ?? '?'}] ${entry.message}`)
      } else {
        console.log(`\nOK ${where}: ${logged}.`)
      }
    }
  })
  if (args.verify) {
    if (stopped) console.log('\nNot verified: provisioning stopped.')
    else if ((await verifySteps(connection, verified, args)).length > 0)
      failed = true
  } else if (args.strict && !failed)
    console.log(
      '\nNote: --strict sees what the handlers log as errors; add --verify to read the result ' +
        'back from the site and compare it with the template.'
    )
  return failed ? 1 : 0
}

main()
  .then((code) => {
    process.exitCode = code
  })
  .catch((error) => {
    if (error instanceof UsageError) {
      console.error(`${error.message}\n\nRun with --help for usage.`)
      process.exitCode = error.exitCode
      return
    }
    // The message only: the error object's response holds the response headers (`Set-Cookie`).
    console.error(describeError(error))
    process.exitCode = 1
  })
