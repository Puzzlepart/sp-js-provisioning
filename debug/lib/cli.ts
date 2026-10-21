import { execFileSync } from 'child_process'
import * as fs from 'fs'
import { rmSync } from 'node:fs'
import * as os from 'os'
import * as path from 'path'

/**
 * The runner's command line and template handling, kept free of I/O against SharePoint so it can be
 * unit-tested (`tests/unit/runner.test.ts`).
 */

/**
 * Handlers the engine cannot run outside the browser/SPFx today:
 *  - Taxonomy and ContentTypes provision through JSOM (`spfx-jsom` -> `SP.ClientContext`), which
 *    is stubbed out in Node (`debug/stubs/spfx-jsom.ts`).
 *  - Files refuses whenever an SPFx context is set, which the runner always sets.
 *  - PropertyBagEntries rejects in Node.
 * They are skipped unless `--handlers` names them.
 */
export const NODE_UNSUPPORTED = [
  'Taxonomy',
  'ContentTypes',
  'Files',
  'PropertyBagEntries'
]

export type Mode = 'hub' | 'cloud'

export interface IArgs {
  site?: string
  package: string
  lang?: string
  mode: Mode
  handlers?: string[]
  parameters: Record<string, string>
  skipTaxonomy: boolean
  strict: boolean
  twice: boolean
  dryRun: boolean
  help: boolean
}

/**
 * A command-line error, with the exit code the runner ends with (2 for usage errors).
 */
export class UsageError extends Error {
  public exitCode = 2

  constructor(message: string) {
    super(message)
    // Keeps `instanceof` working when `Error` subclasses are compiled to ES5.
    Object.setPrototypeOf(this, UsageError.prototype)
  }
}

/**
 * Parses `k=v,k2=v2` into parameters for `{parameter:k}` tokens. A value may contain `=`.
 *
 * @param value - The `--parameters` value
 */
export function parseParameters(value: string): Record<string, string> {
  const parameters: Record<string, string> = {}
  for (const pair of value.split(',').map((part) => part.trim())) {
    if (!pair) continue
    const at = pair.indexOf('=')
    if (at < 1)
      throw new UsageError(
        `Invalid --parameters entry "${pair}" (expected key=value)`
      )
    parameters[pair.slice(0, at).trim()] = pair.slice(at + 1).trim()
  }
  return parameters
}

/**
 * Parses the runner's arguments. `--site` may be left out when `INTEGRATION_SITE_URL` is set; the
 * caller fills it in.
 *
 * @param argv - The arguments after `npm run provision --`
 */
export function parseArgs(argv: string[]): IArgs {
  const out: IArgs = {
    package: '',
    mode: 'hub',
    parameters: {},
    skipTaxonomy: false,
    strict: false,
    twice: false,
    dryRun: false,
    help: false
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
      case '--package':
      case '--pkg':
        out.package = next()
        break
      case '--lang':
        out.lang = next()
        break
      case '--mode': {
        const mode = next()
        if (mode !== 'hub' && mode !== 'cloud')
          throw new UsageError(
            `Invalid --mode "${mode}" (expected hub or cloud)`
          )
        out.mode = mode
        break
      }
      case '--handlers':
        out.handlers = next()
          .split(',')
          .map((name) => name.trim())
          .filter(Boolean)
        break
      case '--parameters':
        out.parameters = { ...out.parameters, ...parseParameters(next()) }
        break
      case '--skip-taxonomy':
        out.skipTaxonomy = true
        break
      case '--strict':
        out.strict = true
        break
      case '--twice':
        out.twice = true
        break
      case '--dry-run':
        out.dryRun = true
        break
      case '--help':
      case '-h':
        out.help = true
        break
      default:
        throw new UsageError(`Unknown argument: ${arg}`)
    }
  }
  if (!out.help && !out.package)
    throw new UsageError('Missing required --package')
  return out
}

/**
 * Checks `--handlers` against the engine's handlers.
 *
 * @param requested - The handlers asked for
 * @param known - The engine's handlers, in run order
 */
export function validateHandlers(requested: string[], known: string[]): void {
  const unknown = requested.filter((name) => !known.includes(name))
  if (unknown.length)
    throw new UsageError(
      `Unknown handler(s): ${unknown.join(', ')}. Valid handlers: ${known.join(
        ', '
      )}`
    )
}

/**
 * The engine's handlers in the order it runs them.
 *
 * @param sort - `DefaultHandlerSort`
 */
export function handlerOrder(sort: Record<string, number>): string[] {
  return Object.keys(sort).sort((a, b) => sort[a] - sort[b])
}

/**
 * Which handlers to run. An explicit `--handlers` wins as-is; otherwise every handler the template
 * has a section for, except the ones that cannot run in Node.
 *
 * @param schema - The template
 * @param known - The engine's handlers, in run order
 * @param requested - `--handlers`
 */
export function planHandlers(
  schema: Record<string, any>,
  known: string[],
  requested?: string[]
): { run: string[]; skipped: string[] } {
  if (requested) {
    validateHandlers(requested, known)
    return { run: requested, skipped: [] }
  }
  const present = known.filter((name) => schema[name] !== undefined)
  return {
    run: present.filter((name) => !NODE_UNSUPPORTED.includes(name)),
    skipped: present.filter((name) => NODE_UNSUPPORTED.includes(name))
  }
}

/**
 * A package manifest, as far as the runner reads it.
 */
export interface IManifest {
  id?: string
  name?: string
  provisioning?: {
    hubTemplate?: string
    template?: string
    listContent?: Array<{ sourceList: string }>
    localized?: Record<
      string,
      {
        hubTemplate?: string
        template?: string
        listContent?: Array<{ sourceList: string }>
      }
    >
  }
}

/**
 * The manifest with a locale's templates and list content, as PP365's
 * `PackageInstaller._applyLocale` merges them.
 *
 * @param manifest - The manifest
 * @param lang - The locale, e.g. `en-US`
 */
export function applyLocale(manifest: IManifest, lang?: string): IManifest {
  const localized = lang && manifest.provisioning?.localized?.[lang]
  if (!localized) return manifest
  return {
    ...manifest,
    provisioning: {
      ...manifest.provisioning,
      ...(localized.hubTemplate ? { hubTemplate: localized.hubTemplate } : {}),
      ...(localized.template ? { template: localized.template } : {}),
      ...(localized.listContent ? { listContent: localized.listContent } : {})
    }
  }
}

/**
 * The subset of a hub template that PP365 applies when a package is published as a cloud template
 * (`PackageInstaller.provisionCloudTemplateHubDependencies`): site fields, content types, and the
 * lists that carry content type bindings or hub configuration rows, minus the lists whose rows are
 * project content (`listContent` sources). Taxonomy is kept unless skipped.
 *
 * @param schema - The hub template
 * @param manifest - The package manifest (none for a raw template)
 * @param skipTaxonomy - Drop the Taxonomy section
 */
export function cloudHubSubset(
  schema: Record<string, any>,
  manifest?: IManifest,
  skipTaxonomy = false
): Record<string, any> {
  const filtered: Record<string, any> = {}
  if (schema.SiteFields?.length) filtered.SiteFields = schema.SiteFields
  if (schema.ContentTypes?.length) filtered.ContentTypes = schema.ContentTypes
  const listContentSources = new Set(
    (manifest?.provisioning?.listContent ?? []).map((entry) => entry.sourceList)
  )
  const hubLists = (schema.Lists ?? []).filter((list: any) => {
    if (listContentSources.has(list.Title)) return false
    const hasBindings =
      Array.isArray(list.ContentTypeBindings) &&
      list.ContentTypeBindings.length > 0
    return (hasBindings && !list.DataRows) || !!list.DataRows
  })
  if (hubLists.length) filtered.Lists = hubLists
  const hasTaxonomy = !!(
    schema.Taxonomy &&
    ((schema.Taxonomy.TermSets?.length ?? 0) > 0 || schema.Taxonomy.TermGroup)
  )
  if (hasTaxonomy && !skipTaxonomy) filtered.Taxonomy = schema.Taxonomy
  return filtered
}

/**
 * The template to apply after the hub/cloud/taxonomy transforms PP365's PackageInstaller applies.
 *
 * @param schema - The template
 * @param args - The arguments
 * @param manifest - The package manifest (none for a raw template)
 */
export function transformSchema(
  schema: Record<string, any>,
  args: Pick<IArgs, 'mode' | 'skipTaxonomy'>,
  manifest?: IManifest
): Record<string, any> {
  if (args.mode === 'cloud')
    return cloudHubSubset(schema, manifest, args.skipTaxonomy)
  if (args.skipTaxonomy && schema.Taxonomy) {
    const { Taxonomy, ...rest } = schema
    return rest
  }
  return schema
}

/**
 * Resolves the template to apply from `--package`: a package folder (with `manifest.json`), a
 * `.pppkg`/`.zip` (extracted with `unzip`), or a raw template JSON. A package resolves to its hub
 * template, as PP365's PackageInstaller picks `manifest.provisioning.hubTemplate`.
 *
 * @param args - The arguments
 */
export function resolveTemplate(
  args: Pick<IArgs, 'package' | 'lang'>
): { schema: Record<string, any>; label: string; manifest?: IManifest } {
  const stat = fs.statSync(args.package)
  if (stat.isFile() && /\.(pppkg|zip)$/i.test(args.package)) {
    const extracted = fs.mkdtempSync(path.join(os.tmpdir(), 'pp-provision-'))
    try {
      execFileSync('unzip', ['-o', '-q', args.package, '-d', extracted])
      return readPackage(extracted, args.lang)
    } finally {
      rmSync(extracted, { recursive: true, force: true })
    }
  }
  if (stat.isFile())
    return {
      schema: JSON.parse(fs.readFileSync(args.package, 'utf8')),
      label: args.package
    }
  return readPackage(args.package, args.lang)
}

/**
 * Reads a package folder's manifest (with a locale applied) and its hub template.
 *
 * @param dir - The package folder
 * @param lang - The locale, e.g. `en-US`
 */
function readPackage(
  dir: string,
  lang?: string
): { schema: Record<string, any>; label: string; manifest: IManifest } {
  const manifestPath = path.join(dir, 'manifest.json')
  if (!fs.existsSync(manifestPath))
    throw new UsageError(`No manifest.json found in ${dir}`)
  const manifest = applyLocale(
    JSON.parse(fs.readFileSync(manifestPath, 'utf8')),
    lang
  )
  const hubTemplate = manifest.provisioning?.hubTemplate
  if (!hubTemplate)
    throw new UsageError(
      `manifest.provisioning.hubTemplate missing in ${manifestPath}`
    )
  return {
    schema: JSON.parse(fs.readFileSync(path.join(dir, hubTemplate), 'utf8')),
    label: `${manifest.name ?? manifest.id} -> ${hubTemplate}`,
    manifest
  }
}

/**
 * One line per template section, with item counts.
 *
 * @param schema - The template
 */
export function describeSchema(schema: Record<string, any>): string {
  return Object.keys(schema)
    .map((key) => {
      const value = schema[key]
      return `  - ${key}${Array.isArray(value) ? `[${value.length}]` : ''}`
    })
    .join('\n')
}

/**
 * The usage text.
 *
 * @param handlers - The engine's handlers, in run order
 */
export function usage(handlers: string[]): string {
  return [
    'Usage: npm run provision -- --package <path> [--site <url>] [options]',
    '',
    '  --site <url>          Target site (default: INTEGRATION_SITE_URL from .env)',
    '  --package <path>      Package folder, .pppkg/.zip, or template JSON',
    '  --lang <locale>       Localized hub template (e.g. en-US). Default: the package default.',
    '  --mode <hub|cloud>    hub: the full hub template (default); cloud: the subset PP365',
    '                        applies when a package is published as a cloud template',
    '  --handlers <csv>      Only run these handlers (also forces browser-only ones)',
    '  --parameters <k=v,…>  Values for {parameter:k} tokens',
    '  --skip-taxonomy       Drop the Taxonomy section',
    '  --strict              Exit 1 when a handler logged an error, not only when one threw',
    '  --twice               Apply twice and report what the second run logged',
    '  --dry-run             Print the plan and exit, without signing in',
    '',
    `  Handlers, in run order: ${handlers.join(', ')}`,
    `  Skipped by default in Node: ${NODE_UNSUPPORTED.join(', ')}`
  ].join('\n')
}
