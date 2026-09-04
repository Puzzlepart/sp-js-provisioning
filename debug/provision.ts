/* eslint-disable no-console */
/**
 * Local provisioning runner — apply a Prosjektportalen template package to a real
 * SharePoint site straight from Node, without the SPFx build/deploy loop.
 *
 * It reproduces what the catalog's `PackageInstaller` does in the browser
 * (`WebProvisioner(web).setup({ spfxContext, logging }).applyTemplate(schema)`),
 * but authenticates with `@pnp/nodejs` (MSAL app-only) so you can iterate on the
 * engine handlers (Lists, ContentTypes, SiteFields, Taxonomy, …) and see every
 * log line — including the SharePoint error payloads — in your terminal.
 *
 * Usage:
 *   npm run provision -- --site <url> --package <path> [options]
 *
 *   --site <url>        Target site, e.g. https://contoso.sharepoint.com/sites/hub
 *   --package <path>    A package folder (with manifest.json), a .pppkg/.zip,
 *                       or a raw hub-template JSON file.
 *   --lang <locale>     Use the localized hub template, e.g. en-US (default: nb-NO).
 *   --mode <hub|cloud>  hub = full hub template (default, like a Mode A import).
 *                       cloud = only SiteFields + ContentTypes + binding lists
 *                       (like provisionCloudTemplateHubDependencies).
 *   --handlers <csv>    Only run these handlers, e.g. SiteFields,ContentTypes,Lists.
 *   --skip-taxonomy     Drop the Taxonomy section before applying.
 *   --dry-run           Print the resolved template (handlers + counts) and exit.
 *
 * Auth: copy debug/provision.settings.example.ts → debug/provision.settings.ts
 * (gitignored) and fill in your app registration. Run with Node 18+ (your system
 * default is fine — the Node 16 requirement only applies to the SPFx build).
 */
import * as fs from 'fs'
import * as os from 'os'
import * as path from 'path'
import { execFileSync } from 'child_process'

const HANDLERS = [
  'Taxonomy',
  'SiteFields',
  'ContentTypes',
  'Features',
  'Lists',
  'Files',
  'CustomActions',
  'ComposedLook',
  'ClientSidePages',
  'PropertyBagEntries',
  'Navigation',
  'Hooks'
]

/**
 * Handlers the engine can't run outside the browser/SPFx:
 *  - Taxonomy and ContentTypes provision via JSOM (`spfx-jsom` → SP.ClientContext),
 *    which is browser-only — stubbed out in Node.
 *  - Files throws when an SPFx context is present.
 *  - PropertyBagEntries rejects in Node.
 * Auto-skipped unless the caller forces a handler set with --handlers. (Forcing
 * ContentTypes/Taxonomy will fail with a clear "JSOM not available" error.)
 *
 * What DOES run server-side: SiteFields and Lists (incl. folders, views, rows and
 * content-type bindings) — they use `@pnp/sp` REST.
 */
const NODE_UNSUPPORTED = ['Taxonomy', 'ContentTypes', 'Files', 'PropertyBagEntries']

interface IArgs {
  site: string
  package: string
  lang?: string
  mode: 'hub' | 'cloud'
  handlers?: string[]
  skipTaxonomy: boolean
  dryRun: boolean
}

function parseArgs(argv: string[]): IArgs {
  const out: any = { mode: 'hub', skipTaxonomy: false, dryRun: false }
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]
    const next = () => argv[++i]
    switch (a) {
      case '--site':
        out.site = next()
        break
      case '--package':
      case '--pkg':
        out.package = next()
        break
      case '--lang':
        out.lang = next()
        break
      case '--mode':
        out.mode = next()
        break
      case '--handlers':
        out.handlers = next()
          .split(',')
          .map((s) => s.trim())
          .filter(Boolean)
        break
      case '--skip-taxonomy':
        out.skipTaxonomy = true
        break
      case '--dry-run':
        out.dryRun = true
        break
      case '--help':
      case '-h':
        printUsageAndExit(0)
        break
      default:
        console.error(`Unknown argument: ${a}`)
        printUsageAndExit(1)
    }
  }
  if (!out.site || !out.package) {
    console.error('Missing required --site and/or --package.')
    printUsageAndExit(1)
  }
  if (out.mode !== 'hub' && out.mode !== 'cloud') {
    console.error(`Invalid --mode "${out.mode}" (expected hub or cloud).`)
    printUsageAndExit(1)
  }
  return out as IArgs
}

function printUsageAndExit(code: number): never {
  console.log(
    [
      'Usage: npm run provision -- --site <url> --package <path> [options]',
      '',
      '  --site <url>        Target SharePoint site',
      '  --package <path>    Package folder, .pppkg/.zip, or hub-template JSON',
      '  --lang <locale>     Localized hub template (e.g. en-US). Default nb-NO.',
      '  --mode <hub|cloud>  hub (full, default) or cloud (hub dependencies only)',
      '  --handlers <csv>    Only run these handlers',
      '  --skip-taxonomy     Drop the Taxonomy section',
      '  --dry-run           Print the resolved template and exit',
      '',
      `  Handlers: ${HANDLERS.join(', ')}`
    ].join('\n')
  )
  process.exit(code)
}

/**
 * The engine's `util` references `window`/`document` unconditionally (it is a
 * browser-first library). Provide just enough of a DOM so the non-browser code
 * paths run. Must be called before importing the engine.
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
    g.document = { location: { protocol: url.protocol, hostname: url.hostname } }
  }
}

/**
 * A minimal SPFx-context stand-in. The engine only reads
 * `pageContext.site.{absoluteUrl,serverRelativeUrl}` (for `{site}` token
 * replacement), which is all we supply.
 */
function buildSpfxContext(site: string): any {
  const url = new URL(site)
  const serverRelativeUrl = url.pathname || '/'
  const page = {
    site: { absoluteUrl: site, serverRelativeUrl },
    web: { absoluteUrl: site, serverRelativeUrl }
  }
  return { pageContext: page }
}

async function loadSettings(): Promise<any> {
  try {
    const mod = await import('./provision.settings')
    return (mod as any).default ?? mod
  } catch (error) {
    console.error(
      [
        'Could not load debug/provision.settings.ts.',
        'Copy the example and fill in your app registration:',
        '',
        '  cp debug/provision.settings.example.ts debug/provision.settings.ts',
        ''
      ].join('\n')
    )
    throw error
  }
}

function buildMsalConfig(settings: any): any {
  const auth: any = {
    clientId: settings.clientId,
    authority: `https://login.microsoftonline.com/${settings.tenantId}`
  }
  if (settings.clientSecret) {
    auth.clientSecret = settings.clientSecret
  } else if (settings.certificate?.privateKeyPath) {
    auth.clientCertificate = {
      thumbprint: settings.certificate.thumbprint,
      privateKey: fs.readFileSync(settings.certificate.privateKeyPath, 'utf8')
    }
  } else {
    throw new Error(
      'provision.settings.ts must define either clientSecret or certificate.{thumbprint,privateKeyPath}.'
    )
  }
  return { auth }
}

/**
 * Resolve the sp-js-provisioning `Schema` to apply from the --package argument,
 * mirroring how PackageInstaller picks `manifest.provisioning.hubTemplate`.
 */
function resolveTemplate(args: IArgs): { schema: any; label: string } {
  let dir = args.package
  let label = args.package

  const stat = fs.statSync(args.package)

  // A .pppkg/.zip — extract to a temp dir first.
  if (stat.isFile() && /\.(pppkg|zip)$/i.test(args.package)) {
    const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pp-provision-'))
    execFileSync('unzip', ['-o', '-q', args.package, '-d', tmp])
    dir = tmp
    label = `${args.package} (extracted)`
  } else if (stat.isFile()) {
    // A raw template JSON.
    return { schema: JSON.parse(fs.readFileSync(args.package, 'utf8')), label }
  }

  // A package folder (or extracted .pppkg): read the manifest, pick the template.
  const manifestPath = path.join(dir, 'manifest.json')
  if (!fs.existsSync(manifestPath)) {
    throw new Error(`No manifest.json found in ${dir}`)
  }
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
  const localized =
    args.lang && manifest.provisioning?.localized?.[args.lang]?.hubTemplate
  const hubTemplate = localized || manifest.provisioning?.hubTemplate
  if (!hubTemplate) {
    throw new Error(`manifest.provisioning.hubTemplate missing in ${manifestPath}`)
  }
  const templatePath = path.join(dir, hubTemplate)
  label = `${manifest.name ?? manifest.id} → ${hubTemplate}`
  return { schema: JSON.parse(fs.readFileSync(templatePath, 'utf8')), label }
}

/** Apply the hub/cloud/taxonomy transforms PackageInstaller applies. */
function transformSchema(schema: any, args: IArgs): any {
  if (args.mode === 'cloud') {
    const filtered: any = {}
    if (schema.SiteFields?.length) filtered.SiteFields = schema.SiteFields
    if (schema.ContentTypes?.length) filtered.ContentTypes = schema.ContentTypes
    const bindingLists = (schema.Lists ?? []).filter(
      (list: any) =>
        Array.isArray(list.ContentTypeBindings) &&
        list.ContentTypeBindings.length > 0 &&
        !list.DataRows
    )
    if (bindingLists.length) filtered.Lists = bindingLists
    return filtered
  }
  if (args.skipTaxonomy && schema.Taxonomy) {
    const { Taxonomy, ...rest } = schema
    return rest
  }
  return schema
}

function describe(schema: any): string {
  return Object.keys(schema)
    .map((key) => {
      const value = schema[key]
      const count = Array.isArray(value) ? `[${value.length}]` : ''
      return `  • ${key}${count}`
    })
    .join('\n')
}

/**
 * Decide which handlers to run. An explicit --handlers wins as-is; otherwise run
 * every handler present in the template except the ones that can't run in Node.
 */
function planHandlers(
  schema: any,
  args: IArgs
): { run: string[]; skipped: string[] } {
  if (args.handlers) return { run: args.handlers, skipped: [] }
  const present = Object.keys(schema).filter((key) => HANDLERS.includes(key))
  return {
    run: present.filter((key) => !NODE_UNSUPPORTED.includes(key)),
    skipped: present.filter((key) => NODE_UNSUPPORTED.includes(key))
  }
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2))

  // Shims first — before any engine code is imported.
  installBrowserShims(args.site)

  const { schema: rawSchema, label } = resolveTemplate(args)
  const schema = transformSchema(rawSchema, args)
  const plan = planHandlers(schema, args)

  console.log(`\nSite:     ${args.site}`)
  console.log(`Package:  ${label}`)
  console.log(`Mode:     ${args.mode}${args.skipTaxonomy ? ' (skip taxonomy)' : ''}`)
  console.log(`Template handlers:\n${describe(schema)}`)
  console.log(`\nWill run: ${plan.run.join(', ') || '(none)'}`)
  if (plan.skipped.length) {
    console.log(
      `Skipped:  ${plan.skipped.join(', ')} (browser/SPFx-only — pass --handlers to force)`
    )
  }
  console.log('')

  if (args.dryRun) {
    console.log('Dry run — not applying.')
    return
  }

  const settings = await loadSettings()

  // Imported after the shims are in place.
  const { spfi } = await import('@pnp/sp')
  await import('@pnp/sp/presets/all')
  const { SPDefault } = await import('@pnp/nodejs')
  const { FunctionListener, Logger, LogLevel } = await import('@pnp/logging')
  const { WebProvisioner } = await import('../src/index')

  const url = new URL(args.site)
  const sp = spfi(args.site).using(
    SPDefault({
      msal: {
        config: buildMsalConfig(settings),
        scopes: [`${url.protocol}//${url.host}/.default`]
      }
    })
  )

  // Surface the data payload of warnings/errors (e.g. the SharePoint 500 body),
  // which the default ConsoleListener doesn't print.
  Logger.subscribe(
    FunctionListener((entry) => {
      if (entry.level >= LogLevel.Warning && entry.data) {
        console.log('   ↳ data:', JSON.stringify(entry.data, null, 2))
      }
    })
  )

  const provisioner = new WebProvisioner((sp as any).web).setup({
    spfxContext: buildSpfxContext(args.site),
    logging: { prefix: '(provision-debug)', activeLogLevel: LogLevel.Info }
  } as any)

  console.log('Applying template…\n')
  try {
    await provisioner.applyTemplate(schema, plan.run, (handler) => {
      console.log(`\n▶ ${handler}`)
    })
    console.log('\n✅ Done — template applied without a terminal error.')
  } catch (error: any) {
    console.error('\n❌ Provisioning failed.')
    if (error?.handler) console.error(`   Failing handler: ${error.handler}`)
    console.error(`   Message: ${error?.message ?? error}`)
    if (error?.stack) console.error(error.stack)
    process.exitCode = 1
  }
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
