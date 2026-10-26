/* eslint-disable no-console */
/**
 * Removes what the integration tests and the runner leave behind. Prints the plan; removes only
 * with `--apply`.
 *
 *   npm run cleanup -- --webs                     scratch webs and their groups older than 2 h
 *   npm run cleanup -- --webs --older-than 30m    ... older than 30 minutes
 *   npm run cleanup -- --webs --all               ... of any age (no other run may be active)
 *   npm run cleanup -- --prefix 'SPJS Itest'      prefixed lists, pages, content types, site
 *                                                 fields, custom actions, navigation nodes and
 *                                                 site groups on the site
 *   ... --apply                                   remove them
 *
 * The site is INTEGRATION_SITE_URL unless --site names a web below it (or --allow-site repeats a
 * site outside it).
 */
import { assertConnectedWritable, assertWritable, connect } from './lib/connect'
import {
  ICleanupItem,
  parseAge,
  prefixedArtefacts,
  scratchLeftovers
} from './lib/cleanup'
import { UsageError } from './lib/cli'
import { env } from './lib/env'
import { describeError } from './lib/redact'

interface ICleanupArgs {
  site?: string
  allowSite?: string
  webs: boolean
  prefix?: string
  olderThan: number
  all: boolean
  apply: boolean
}

/**
 * Parses the arguments.
 *
 * @param argv - The arguments after `npm run cleanup --`
 */
function parseCleanupArgs(argv: string[]): ICleanupArgs {
  const out: ICleanupArgs = {
    webs: false,
    olderThan: parseAge('2h'),
    all: false,
    apply: false
  }
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    const next = () => {
      const value = argv[++i]
      if (value === undefined || value.startsWith('--'))
        throw new UsageError(`Missing value for ${arg}`)
      return value
    }
    switch (arg) {
      case '--site':
        out.site = next()
        break
      case '--allow-site':
        out.allowSite = next()
        break
      case '--webs':
        out.webs = true
        break
      case '--prefix':
        out.prefix = next()
        break
      case '--older-than':
        out.olderThan = parseAge(next())
        break
      case '--all':
        out.all = true
        break
      case '--apply':
        out.apply = true
        break
      default:
        throw new UsageError(`Unknown argument: ${arg}`)
    }
  }
  if (!out.webs && !out.prefix)
    throw new UsageError(
      'Nothing to clean up: pass --webs and/or --prefix <text>'
    )
  return out
}

async function main(): Promise<number> {
  const args = parseCleanupArgs(process.argv.slice(2))
  const site = (args.site ?? env.siteUrl ?? '').replace(/\/+$/, '')
  if (!site)
    throw new UsageError('Missing --site (or INTEGRATION_SITE_URL in .env)')
  assertWritable(site, args.allowSite)
  const conn = await connect(site)
  assertConnectedWritable(conn, args.allowSite)
  console.log(`Site: ${conn.webInfo.Url} (signed in with ${conn.mode})`)

  const items: ICleanupItem[] = []
  if (args.webs)
    items.push(
      ...(await scratchLeftovers(conn, {
        olderThan: args.olderThan,
        all: args.all
      }))
    )
  if (args.prefix) items.push(...(await prefixedArtefacts(conn, args.prefix)))

  if (!items.length) {
    console.log('Nothing to remove.')
    return 0
  }
  console.log(`\n${args.apply ? 'Removing' : 'Would remove'} ${items.length}:`)
  let failed = 0
  for (const item of items) {
    if (!args.apply) {
      console.log(`  - ${item.kind}: ${item.name}`)
      continue
    }
    try {
      await item.remove()
      console.log(`  removed ${item.kind}: ${item.name}`)
    } catch (error) {
      failed++
      console.error(
        `  FAILED ${item.kind}: ${item.name}: ${describeError(error)}`
      )
    }
  }
  if (!args.apply) console.log('\nDry run: pass --apply to remove them.')
  return failed ? 1 : 0
}

main()
  .then((code) => {
    process.exitCode = code
  })
  .catch((error) => {
    console.error(describeError(error))
    process.exitCode = error instanceof UsageError ? error.exitCode : 1
  })
