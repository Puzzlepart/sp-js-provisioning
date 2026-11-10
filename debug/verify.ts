/* eslint-disable no-console */
/**
 * `npm run verify`: compares a template or a Prosjektportalen package with a web and reports what
 * differs (docs/specs/integration-testing.md, Harness). Read-only: it signs in as
 * `npm run provision` does, reads the web and writes nothing, so it needs no site guard.
 *
 * Exit codes: 0 no finding, 1 a finding (or the sign-in was refused), 2 a usage error.
 *
 * Usage:
 *   npm run verify -- --package <path> [--site <url>] [--exact] [--json]
 *   npm run verify -- --help
 */
import {
  handlerOrder,
  IStep,
  parseArgs,
  projectSteps,
  resolveTemplate,
  transformSchema,
  usage,
  UsageError,
  verifyTargets
} from './lib/cli'
import { DefaultHandlerSort } from '../src/handlers/exports'
import { accessDenied, connect, ISiteConnection } from './lib/connect'
import { authMode, env } from './lib/env'
import { describeError } from './lib/redact'
import { formatResult, IVerifyResult, verify } from './lib/verify'

async function main(): Promise<number> {
  const args = parseArgs(process.argv.slice(2), 'verify')
  const site = (args.site ?? env.siteUrl ?? '').replace(/\/+$/, '')
  if (args.help) {
    console.log(usage(handlerOrder(DefaultHandlerSort), 'verify'))
    return 0
  }
  const resolved = resolveTemplate(args)
  const schema = transformSchema(resolved.schema, args, resolved.manifest)
  const targets: IStep[] =
    args.mode === 'project'
      ? verifyTargets(projectSteps(schema, resolved.extensions ?? []).apply)
      : [{ source: resolved.label, schema }]
  if (!site)
    throw new UsageError('Missing --site (or INTEGRATION_SITE_URL in .env)')

  let connection: ISiteConnection
  try {
    connection = await connect(site)
  } catch (error) {
    const message = accessDenied(error, site, authMode())
    if (!message) throw error
    console.error(message)
    return 1
  }
  if (!args.json)
    console.log(
      `Verifying ${resolved.label} (${args.mode}${
        args.exact ? ', exact' : ''
      }) against "${connection.webInfo.Title}" (${connection.webInfo.Url})\n`
    )

  const results: Array<IVerifyResult & { source: string }> = []
  for (const target of targets) {
    const result = await verify(connection, target.schema, {
      exact: args.exact,
      handlers: args.handlers,
      parameters: args.parameters
    })
    if (targets.length > 1)
      result.findings.forEach((finding) => (finding.source = target.source))
    results.push({ source: target.source, ...result })
  }
  const findings = results.reduce(
    (count, result) => count + result.findings.length,
    0
  )
  if (args.json) console.log(JSON.stringify({ findings, results }, null, 2))
  else
    for (const result of results) {
      if (results.length > 1) console.log(`${result.source}:`)
      for (const line of formatResult(result))
        console.log(`${results.length > 1 ? '   ' : ''}${line}`)
    }
  return findings > 0 ? 1 : 0
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
