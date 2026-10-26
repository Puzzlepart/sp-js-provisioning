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
  parseArgs,
  planHandlers,
  resolveTemplate,
  transformSchema,
  usage,
  UsageError
} from './lib/cli'
import { DefaultHandlerSort } from '../src/handlers/exports'
import { apply } from './lib/apply'
import {
  accessDenied,
  assertConnectedWritable,
  assertWritable,
  connect,
  ISiteConnection
} from './lib/connect'
import { authMode, env } from './lib/env'
import { describeError } from './lib/redact'

/**
 * Prints the plan for a run.
 *
 * @param site - The target site
 * @param label - What the template came from
 * @param args - The arguments
 * @param schema - The template to apply
 * @param plan - The handlers to run and skip
 */
function printPlan(
  site: string,
  label: string,
  args: IArgs,
  schema: Record<string, any>,
  plan: { run: string[]; skipped: string[] }
): void {
  console.log(`\nSite:     ${site || '(not set)'}`)
  console.log(`Package:  ${label}`)
  console.log(
    `Mode:     ${args.mode}${args.skipTaxonomy ? ' (skip taxonomy)' : ''}${
      args.twice ? ', twice' : ''
    }${args.strict ? ', strict' : ''}`
  )
  console.log(`Template sections:\n${describeSchema(schema)}`)
  console.log(`\nWill run: ${plan.run.join(', ') || '(none)'}`)
  if (plan.skipped.length)
    console.log(
      `Skipped:  ${plan.skipped.join(
        ', '
      )} (browser/SPFx-only in Node; pass --handlers to force)`
    )
  console.log('')
}

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2))
  const site = (args.site ?? env.siteUrl ?? '').replace(/\/+$/, '')
  const handlers = handlerOrder(DefaultHandlerSort)
  if (args.help) {
    console.log(usage(handlers))
    return 0
  }
  const { schema: rawSchema, label, manifest } = resolveTemplate(args)
  const schema = transformSchema(rawSchema, args, manifest)
  const plan = planHandlers(schema, handlers, args.handlers)
  printPlan(site, label, args, schema, plan)
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

  /**
   * One run, printing the engine's log lines (with their data, e.g. the SharePoint error body).
   */
  const run = (attempt: number) =>
    apply(connection, schema, {
      handlers: plan.run,
      parameters: args.parameters,
      prefix: `(provision${args.twice ? ` #${attempt}` : ''})`,
      onHandler: (handler) => console.log(`\n> ${handler}`)
    })

  const results = []
  for (let attempt = 1; attempt <= (args.twice ? 2 : 1); attempt++) {
    if (args.twice) console.log(`\n=== Run ${attempt} of 2 ===`)
    results.push(await run(attempt))
    if (results[results.length - 1].error) break
  }

  let failed = false
  results.forEach((result, index) => {
    const name = args.twice ? `Run ${index + 1}` : 'Run'
    if (result.error) {
      failed = true
      console.error(`\nFAILED ${name}: provisioning stopped.`)
      if (result.error.handler)
        console.error(`   Failing handler: ${result.error.handler}`)
      console.error(`   Message: ${describeError(result.error)}`)
      return
    }
    const logged = `${result.errors.length} error(s), ${result.warnings.length} warning(s) logged`
    if (result.errors.length) {
      if (args.strict) failed = true
      console.log(
        `\n${args.strict ? 'FAILED' : 'DONE with errors'} ${name}: ${logged}:`
      )
      for (const entry of result.errors)
        console.log(`   [${entry.source ?? '?'}] ${entry.message}`)
    } else {
      console.log(`\nOK ${name}: ${logged}.`)
    }
  })
  if (args.strict && !failed)
    console.log(
      '\nNote: --strict sees what the handlers log as errors; it does not read the result ' +
        'back from the site (docs/specs/integration-testing.md, Phase 1).'
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
