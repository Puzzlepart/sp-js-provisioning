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

/**
 * What `--mode` applies: `hub` the hub template, `cloud` the part of it PP365 applies when a package
 * is published as a cloud template, `project` the project template and its extensions as the
 * project setup's `ApplyTemplate` and `Hooks` steps do.
 */
export type Mode = 'hub' | 'cloud' | 'project'

/** The runner's commands. They share the options; `verify` reads and never writes. */
export type Command = 'provision' | 'verify'

export interface IArgs {
  site?: string
  /** `--allow-site`: write to a site outside `INTEGRATION_SITE_URL` (must equal the target). */
  allowSite?: string
  package: string
  lang?: string
  mode: Mode
  /**
   * `--extensions` (`--mode project`): the project extensions to apply, by id; `all` or `none`.
   * Left out, the ones the setup selects to begin with (`defaultSelected` or `locked`).
   */
  extensions?: string[] | 'all' | 'none'
  handlers?: string[]
  parameters: Record<string, string>
  skipTaxonomy: boolean
  strict: boolean
  twice: boolean
  /** `provision --verify`: diff the template against the web after applying it. */
  verify: boolean
  /** `--exact`: also compare what a re-apply leaves as it is (a fresh web). */
  exact: boolean
  /** `verify --json`: print the result as JSON. */
  json: boolean
  dryRun: boolean
  help: boolean
}

/** Options only one command takes. */
const COMMAND_ONLY: Record<string, Command> = {
  '--allow-site': 'provision',
  '--strict': 'provision',
  '--twice': 'provision',
  '--verify': 'provision',
  '--dry-run': 'provision',
  '--json': 'verify'
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
 * Parses `--extensions`: ids, `all` or `none`.
 *
 * @param value - The `--extensions` value
 */
export function parseExtensions(value: string): IArgs['extensions'] {
  const trimmed = value.trim()
  if (trimmed === 'all' || trimmed === 'none') return trimmed
  const ids = trimmed
    .split(',')
    .map((id) => id.trim())
    .filter(Boolean)
  if (ids.length === 0)
    throw new UsageError('Invalid --extensions (expected ids, all or none)')
  return ids
}

/**
 * Parses the arguments of `npm run provision` or `npm run verify`. `--site` may be left out when
 * `INTEGRATION_SITE_URL` is set; the caller fills it in.
 *
 * @param argv - The arguments after `npm run <command> --`
 * @param command - The command they are for
 */
export function parseArgs(
  argv: string[],
  command: Command = 'provision'
): IArgs {
  const out: IArgs = {
    package: '',
    mode: 'hub',
    parameters: {},
    skipTaxonomy: false,
    strict: false,
    twice: false,
    verify: false,
    exact: false,
    json: false,
    dryRun: false,
    help: false
  }
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]
    if (COMMAND_ONLY[arg] && COMMAND_ONLY[arg] !== command)
      throw new UsageError(
        `${arg} is an option of npm run ${COMMAND_ONLY[arg]}`
      )
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
      case '--package':
      case '--pkg':
        out.package = next()
        break
      case '--lang':
        out.lang = next()
        break
      case '--mode': {
        const mode = next()
        if (mode !== 'hub' && mode !== 'cloud' && mode !== 'project')
          throw new UsageError(
            `Invalid --mode "${mode}" (expected hub, cloud or project)`
          )
        out.mode = mode
        break
      }
      case '--extensions':
        out.extensions = parseExtensions(next())
        break
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
      case '--verify':
        out.verify = true
        break
      case '--exact':
        out.exact = true
        break
      case '--json':
        out.json = true
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
  if (out.help) return out
  if (!out.package) throw new UsageError('Missing required --package')
  if (out.extensions && out.mode !== 'project')
    throw new UsageError('--extensions needs --mode project')
  if (out.exact && command === 'provision' && !out.verify)
    throw new UsageError('--exact needs --verify')
  return out
}

/**
 * The arguments of `npm run inspect`.
 */
export interface IInspectArgs {
  site?: string
  prefix?: string
  out?: string
  help: boolean
}

/**
 * Parses the arguments of `npm run inspect`.
 *
 * @param argv - The arguments after `npm run inspect --`
 */
export function parseInspectArgs(argv: string[]): IInspectArgs {
  const out: IInspectArgs = { help: false }
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
      case '--prefix':
        out.prefix = next()
        break
      case '--out':
        out.out = next()
        break
      case '--help':
      case '-h':
        out.help = true
        break
      default:
        throw new UsageError(`Unknown argument: ${arg}`)
    }
  }
  return out
}

/** The usage text of `npm run inspect`. */
export const INSPECT_USAGE = [
  'Usage: npm run inspect -- [--site <url>] [--prefix <text>] [--out <file>]',
  '',
  'Prints what a web holds of what the handlers write, as JSON. Read-only.',
  '',
  '  --site <url>      The web (default: INTEGRATION_SITE_URL from .env)',
  '  --prefix <text>   Only site columns, content types, site groups, lists, custom actions and',
  '                    pages named with this prefix (permissions, features, navigation,',
  '                    settings and the property bag are read whole)',
  '  --out <file>      Write the JSON to a file instead of printing it',
  '',
  "  The JSON holds the tenant's URLs and user names: do not commit it or attach it to an issue."
].join('\n')

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
    /** Project extensions (prosjekttillegg): templates the project setup applies after the template. */
    extensions?: IManifestExtension[]
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
 * A project extension in a package manifest.
 */
export interface IManifestExtension {
  id: string
  name: string
  /** The extension's template, relative to the package. */
  file: string
  defaultSelected?: boolean
  /** Cannot be deselected in the setup. */
  locked?: boolean
}

/**
 * A project extension's template, read from its package.
 */
export interface IExtensionTemplate {
  id: string
  name: string
  schema: Record<string, any>
}

/**
 * A template the runner applies or verifies, and where it came from.
 */
export interface IStep {
  source: string
  schema: Record<string, any>
}

/** The menus a template's `Navigation` can declare. */
const NAVIGATION_COLLECTIONS = ['QuickLaunch', 'TopNavigationBar']

/**
 * A template without some of its `Navigation` menus (and without `Navigation` when none is left).
 *
 * @param schema - The template
 * @param collections - The menus to leave out
 */
function withoutNavigation(
  schema: Record<string, any>,
  collections: string[]
): Record<string, any> {
  if (!schema.Navigation) return schema
  const navigation = { ...schema.Navigation }
  for (const collection of collections) delete navigation[collection]
  const { Navigation, ...rest } = schema
  return Object.keys(navigation).length > 0
    ? { ...rest, Navigation: navigation }
    : rest
}

/**
 * The menus a template declares.
 *
 * @param schema - The template
 */
const declaredMenus = (schema: Record<string, any>): string[] =>
  NAVIGATION_COLLECTIONS.filter((collection) =>
    Array.isArray(schema.Navigation?.[collection])
  )

/**
 * What PP365's project setup applies for a template and the selected extensions, in its order:
 * the `ApplyTemplate` step (the template without `Hooks` and without each menu an extension
 * declares too, then each extension without `Hooks`) and the `Hooks` step (the template's hooks,
 * then each extension's). The rest of the setup (PreTask, SitePermissions, SetTaxonomyFields,
 * CopyListData, Planner) is not mirrored.
 *
 * @param template - The project template
 * @param extensions - The selected extensions, in the order the setup applies them
 */
export function projectSteps(
  template: Record<string, any>,
  extensions: IExtensionTemplate[]
): { apply: IStep[]; hooks: IStep[] } {
  const withoutHooks = (schema: Record<string, any>) => {
    const { Hooks, ...rest } = schema
    return rest
  }
  const extensionMenus = extensions.reduce<string[]>(
    (menus, extension) => menus.concat(declaredMenus(extension.schema)),
    []
  )
  const hooksOf = (source: string, schema: Record<string, any>): IStep[] =>
    Array.isArray(schema.Hooks) && schema.Hooks.length > 0
      ? [{ source, schema: { Hooks: schema.Hooks } }]
      : []
  return {
    apply: [
      {
        source: 'template',
        schema: withoutNavigation(withoutHooks(template), extensionMenus)
      },
      ...extensions.map((extension) => ({
        source: `extension ${extension.id}`,
        schema: withoutHooks(extension.schema)
      }))
    ],
    hooks: [
      ...hooksOf('template hooks', template),
      ...extensions.reduce<IStep[]>(
        (steps, extension) =>
          steps.concat(
            hooksOf(`extension ${extension.id} hooks`, extension.schema)
          ),
        []
      )
    ]
  }
}

/**
 * The templates to verify after applying `steps` in order: a menu a later step declares too is
 * left out of the earlier one, since the later run made the menu its own.
 *
 * @param steps - The templates, in the order they were applied
 */
export function verifyTargets(steps: IStep[]): IStep[] {
  return steps.map((step, index) => ({
    source: step.source,
    schema: withoutNavigation(
      step.schema,
      steps
        .slice(index + 1)
        .reduce<string[]>(
          (menus, later) => menus.concat(declaredMenus(later.schema)),
          []
        )
    )
  }))
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
 * A resolved `--package`: the template, what it came from, the manifest of a package, and in
 * `--mode project` the selected extensions.
 */
export interface IResolvedTemplate {
  schema: Record<string, any>
  label: string
  manifest?: IManifest
  extensions?: IExtensionTemplate[]
}

/**
 * Resolves the template to apply from `--package`: a package folder (with `manifest.json`), a
 * `.pppkg`/`.zip` (extracted with `unzip`), or a raw template JSON. A package resolves to its hub
 * template, as PP365's PackageInstaller picks `manifest.provisioning.hubTemplate`, and in
 * `--mode project` to its project template (`manifest.provisioning.template`) and the selected
 * extensions.
 *
 * @param args - The arguments
 */
export function resolveTemplate(
  args: Pick<IArgs, 'package' | 'lang'> &
    Partial<Pick<IArgs, 'mode' | 'extensions'>>
): IResolvedTemplate {
  const stat = fs.statSync(args.package)
  if (stat.isFile() && /\.(pppkg|zip)$/i.test(args.package)) {
    const extracted = fs.mkdtempSync(path.join(os.tmpdir(), 'pp-provision-'))
    try {
      execFileSync('unzip', ['-o', '-q', args.package, '-d', extracted])
      return readPackage(extracted, args)
    } finally {
      rmSync(extracted, { recursive: true, force: true })
    }
  }
  if (stat.isFile()) {
    if (Array.isArray(args.extensions))
      throw new UsageError('--extensions needs a package (a template has none)')
    return {
      schema: readTemplate(args.package),
      label: args.package,
      ...(args.mode === 'project' ? { extensions: [] } : {})
    }
  }
  return readPackage(args.package, args)
}

/**
 * Reads a template file, without its `$schema`.
 *
 * @param file - The file
 */
function readTemplate(file: string): Record<string, any> {
  const { $schema, ...schema } = JSON.parse(fs.readFileSync(file, 'utf8'))
  return schema
}

/**
 * The extensions `--extensions` selects: by id, all or none; left out, the ones the setup selects
 * to begin with (`defaultSelected` or `locked`).
 *
 * @param available - The manifest's extensions
 * @param selection - `--extensions`
 */
export function selectExtensions(
  available: IManifestExtension[],
  selection?: IArgs['extensions']
): IManifestExtension[] {
  if (selection === 'none') return []
  if (selection === 'all') return available
  if (!selection)
    return available.filter(
      (extension) => extension.defaultSelected || extension.locked
    )
  const unknown = selection.filter(
    (id) => !available.some((extension) => extension.id === id)
  )
  if (unknown.length)
    throw new UsageError(
      `Unknown extension(s): ${unknown.join(', ')}. The package has: ${
        available.map((extension) => extension.id).join(', ') || '(none)'
      }`
    )
  return available.filter((extension) => selection.includes(extension.id))
}

/**
 * Reads a package folder's manifest (with a locale applied) and its hub template, or in
 * `--mode project` its project template and the selected extensions.
 *
 * @param dir - The package folder
 * @param args - The locale, the mode and `--extensions`
 */
function readPackage(
  dir: string,
  args: Pick<IArgs, 'lang'> & Partial<Pick<IArgs, 'mode' | 'extensions'>>
): IResolvedTemplate & { manifest: IManifest } {
  const manifestPath = path.join(dir, 'manifest.json')
  if (!fs.existsSync(manifestPath))
    throw new UsageError(`No manifest.json found in ${dir}`)
  const manifest = applyLocale(
    JSON.parse(fs.readFileSync(manifestPath, 'utf8')),
    args.lang
  )
  const project = args.mode === 'project'
  const key = project ? 'template' : 'hubTemplate'
  const file = manifest.provisioning?.[key]
  if (!file)
    throw new UsageError(
      `manifest.provisioning.${key} missing in ${manifestPath}`
    )
  return {
    schema: readTemplate(path.join(dir, file)),
    label: `${manifest.name ?? manifest.id} -> ${file}`,
    manifest,
    ...(project
      ? {
          extensions: selectExtensions(
            manifest.provisioning?.extensions ?? [],
            args.extensions
          ).map((extension) => ({
            id: extension.id,
            name: extension.name,
            schema: readTemplate(path.join(dir, extension.file))
          }))
        }
      : {})
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
 * @param command - The command
 */
export function usage(
  handlers: string[],
  command: Command = 'provision'
): string {
  const provision = command === 'provision'
  return [
    provision
      ? 'Usage: npm run provision -- --package <path> [--site <url>] [options]'
      : 'Usage: npm run verify -- --package <path> [--site <url>] [options]',
    '',
    provision
      ? 'Applies a template or a package to a site.'
      : 'Compares a template or a package with a site and reports what differs. Read-only.',
    '',
    '  --site <url>          Target site (default: INTEGRATION_SITE_URL from .env)',
    ...(provision
      ? [
          '  --allow-site <url>    Write to a site outside INTEGRATION_SITE_URL (repeat the target)'
        ]
      : []),
    '  --package <path>      Package folder, .pppkg/.zip, or template JSON',
    '  --lang <locale>       Localized template (e.g. en-US). Default: the package default.',
    '  --mode <mode>         hub: the full hub template (default); cloud: the subset PP365',
    '                        applies when a package is published as a cloud template;',
    "                        project: the package's project template and extensions, as the",
    "                        project setup's ApplyTemplate and Hooks steps apply them",
    '  --extensions <ids>    --mode project: extension ids, all or none',
    '                        (default: the ones the setup selects to begin with)',
    provision
      ? '  --handlers <csv>      Only run these handlers (also forces browser-only ones)'
      : '  --handlers <csv>      Only compare these sections (default: every section)',
    '  --parameters <k=v,…>  Values for {parameter:k} tokens',
    '  --skip-taxonomy       Drop the Taxonomy section',
    ...(provision
      ? [
          '  --strict              Exit 1 when a handler logged an error, not only when one threw',
          '  --twice               Apply twice and report what the second run logged',
          '  --verify              Then compare the template with the site (exit 1 on a finding)',
          '  --exact               With --verify: also what a re-apply leaves as it is (a fresh web)',
          '  --dry-run             Print the plan and exit, without signing in'
        ]
      : [
          '  --exact               Also compare what a re-apply leaves as it is (a fresh web)',
          '  --json                Print the result as JSON'
        ]),
    '',
    `  Handlers, in run order: ${handlers.join(', ')}`,
    ...(provision
      ? [`  Skipped by default in Node: ${NODE_UNSUPPORTED.join(', ')}`]
      : [])
  ].join('\n')
}
