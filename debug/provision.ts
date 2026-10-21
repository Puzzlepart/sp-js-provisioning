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
import { authMode, env } from './lib/env'

/**
 * The engine's `util` reads `window`/`document` (it is a browser-first library). Provide just
 * enough of them for the Node code paths. Must run before the engine is imported.
 *
 * @param site - The target site
 */
function installBrowserShims(site: string): void {
  const url = new URL(site)
  const g = globalThis as any
  if (!g.window) {
    g.window = {
      location: {
        protocol: url.protocol,
        host: url.host,
        hostname: url.hostname,
        port: url.port,
        href: site
      },
      btoa: (s: string) => Buffer.from(s, 'binary').toString('base64')
    }
  }
  if (!g.document) {
    g.document = {
      location: { protocol: url.protocol, hostname: url.hostname }
    }
  }
}

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

  // Shims first: before any engine code is imported.
  installBrowserShims(site || 'https://localhost')
  const { DefaultHandlerSort } = await import('../src/handlers/exports')
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

  const { accessDenied, connect } = await import('./lib/connect')
  const { LogLevel, startRun } = await import('./lib/logs')
  const { WebProvisioner } = await import('../src/index')

  let connection: import('./lib/connect').ISiteConnection
  try {
    connection = await connect(site)
  } catch (error) {
    const message = accessDenied(error, site, authMode())
    if (!message) throw error
    console.error(`\n${message}`)
    return 1
  }
  console.log(
    `Signed in with ${connection.mode} to "${connection.webInfo.Title}" (${connection.webInfo.Url})\n`
  )

  /**
   * One run, with fresh logging: the collector, and the engine's console listener, which prints
   * each line with its data (e.g. the SharePoint error body).
   */
  const run = async (attempt: number) => {
    const logs = startRun()
    const provisioner = new WebProvisioner(connection.web).setup({
      spfxContext: connection.context,
      parameters: args.parameters,
      logging: {
        prefix: `(provision${args.twice ? ` #${attempt}` : ''})`,
        activeLogLevel: LogLevel.Info
      }
    })
    let error: any
    try {
      await provisioner.applyTemplate(schema, plan.run, (handler) => {
        console.log(`\n> ${handler}`)
      })
    } catch (caught) {
      error = caught
    }
    return { error, errors: logs.errors(), warnings: logs.warnings() }
  }

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
      console.error(`   Message: ${result.error.message || result.error}`)
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
    console.error(error)
    process.exitCode = 1
  })
