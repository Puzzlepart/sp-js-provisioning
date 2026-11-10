import * as fs from 'fs'
import * as path from 'path'
import { fillPlaceholders } from '../../debug/lib/placeholders'

/** The folder with the fixture templates. */
export const TEMPLATES = path.join(__dirname, 'templates')

/**
 * Values for the two test-only placeholders, written where the engine replaces no tokens: `@@user@@`
 * (a User value or a group member) and `@@hookUrl@@` (a Hooks `Url`).
 */
export interface IFixtureValues {
  user?: string
  hookUrl?: string
}

/**
 * A fresh copy of a fixture template, `tests/fixtures/templates/<name>.json`, with the test-only
 * placeholders filled in. Fresh on every call, because handlers mutate their input (WebSettings
 * writes values back). A placeholder without a value is an error, so a
 * fixture never reaches SharePoint half-filled.
 *
 * @param name - The fixture, e.g. `lists`
 * @param values - The placeholder values
 */
export function loadFixture(
  name: string,
  values: IFixtureValues = {}
): Record<string, any> {
  const parsed = JSON.parse(
    fs.readFileSync(path.join(TEMPLATES, `${name}.json`), 'utf8')
  )
  delete parsed.$schema
  return fillPlaceholders(parsed, values, `Fixture ${name}`)
}

/**
 * The fixtures `all.json` is made of, in the order their sections are joined: every handler
 * fixture without the parts that need something only its own suite sets up
 * (docs/specs/integration-testing.md, Fixtures). Left out: `taxonomy.json` (the tenant's term
 * store, docs/specs/jsom-exit.md J5), `files.json` (a local server and a source in the web, through
 * `{parameter:}`), `propertybagentries.json` and `customactions-script.json` (custom script),
 * `hooks.json` (a local listener), `failures.json` and `smoke.json` (runner checks) and the `.b`
 * variants.
 */
export const ALL_PARTS = [
  'sitefields',
  'contenttypes',
  'features',
  'security',
  'lists',
  'lists-datarows',
  'customactions',
  'composedlook',
  'clientsidepages',
  'navigation',
  'websettings'
]

/**
 * The list `sitefields.json`'s Lookup column points at. Its suite creates it before the first run;
 * `all.json` declares it, so WebProvisioner creates it before SiteFields.
 */
export const SITEFIELDS_SOURCE_LIST = {
  Title: 'SPJS Itest Sitefields Source',
  Description: 'SPJS Itest the Lookup site column points here',
  Template: 100,
  ContentTypesEnabled: false
}

/** The section order of a composed template: the engine's run order. */
const SECTION_ORDER = [
  'SiteFields',
  'ContentTypes',
  'Features',
  'Security',
  'Lists',
  'CustomActions',
  'ComposedLook',
  'ClientSidePages',
  'Navigation',
  'WebSettings'
]

/**
 * `all.json` as the handler fixtures make it: array sections joined in `ALL_PARTS` order, object
 * sections taken from the one fixture that has them, the Lookup's source list added. The
 * placeholders stay, for `loadFixture` (or the runner) to fill.
 */
export function composeAll(): Record<string, any> {
  const composed: Record<string, any> = {
    $schema: '../../../schema.json'
  }
  for (const name of ALL_PARTS) {
    const template = JSON.parse(
      fs.readFileSync(path.join(TEMPLATES, `${name}.json`), 'utf8')
    )
    for (const section of Object.keys(template)) {
      if (section === '$schema') continue
      if (!SECTION_ORDER.includes(section))
        throw new Error(`${name}.json: ${section} is not in all.json`)
      const value = template[section]
      if (Array.isArray(value))
        composed[section] = [...(composed[section] || []), ...value]
      else if (composed[section] !== undefined)
        throw new Error(
          `${name}.json: ${section} is already set by another fixture`
        )
      else composed[section] = value
    }
  }
  composed.Lists = [SITEFIELDS_SOURCE_LIST, ...(composed.Lists || [])]
  const ordered: Record<string, any> = { $schema: composed.$schema }
  for (const section of SECTION_ORDER)
    if (composed[section] !== undefined) ordered[section] = composed[section]
  return ordered
}
