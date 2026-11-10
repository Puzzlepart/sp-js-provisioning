/* eslint-disable no-console */
/**
 * `npm run inspect`: prints what a web holds of what the handlers write, as JSON
 * (docs/specs/integration-testing.md, Commands). Read-only: it signs in as `npm run provision`
 * does and writes nothing to the site.
 *
 * Usage:
 *   npm run inspect -- [--site <url>] [--prefix <text>] [--out <file>]
 */
import * as fs from 'fs'
import { INSPECT_USAGE, parseInspectArgs, UsageError } from './lib/cli'
import { accessDenied, connect, ISiteConnection } from './lib/connect'
import { authMode, env } from './lib/env'
import { inspectWeb } from './lib/inspect'
import { describeError } from './lib/redact'

async function main(): Promise<number> {
  const args = parseInspectArgs(process.argv.slice(2))
  if (args.help) {
    console.log(INSPECT_USAGE)
    return 0
  }
  const site = (args.site ?? env.siteUrl ?? '').replace(/\/+$/, '')
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
  const json = JSON.stringify(
    await inspectWeb(connection, { prefix: args.prefix }),
    null,
    2
  )
  if (args.out) {
    fs.writeFileSync(args.out, `${json}\n`)
    console.error(`Wrote ${args.out}`)
  } else console.log(json)
  return 0
}

main()
  .then((code) => {
    process.exitCode = code
  })
  .catch((error) => {
    if (error instanceof UsageError) {
      console.error(`${error.message}\n\n${INSPECT_USAGE}`)
      process.exitCode = error.exitCode
      return
    }
    // The message only: the error object's response holds the response headers (`Set-Cookie`).
    console.error(describeError(error))
    process.exitCode = 1
  })
