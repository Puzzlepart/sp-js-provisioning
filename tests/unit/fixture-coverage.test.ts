import assert from 'node:assert/strict'
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, it } from 'node:test'

/**
 * Every property of the template contract (schema.json) is used by a fixture in
 * tests/fixtures/templates, so the integration suites apply it, or is listed below with the reason
 * it is not. A new property in src/schema.ts (and so in schema.json) fails this test until a
 * fixture uses it or a reason is given.
 */

const JSOM =
  'JSOM handler, browser-only until the JSOM exit (integration-testing.md D3)'
const PASS_THROUGH =
  'passed to SharePoint unchanged with the other settings; the fixtures cover the path with other keys'

const NOT_COVERED: Record<string, string> = {
  'Schema.Version': 'metadata: the engine has no handler for it and skips it',
  'Schema.Parameters':
    'the {parameter:} token is covered by tests/unit (runner, csom); a parameters.json fixture comes with the depth part of Phase 2',
  'IFileObject.Properties': JSOM,
  'IFileObject.RemoveExistingWebParts': JSOM,
  'IFileObject.WebParts': JSOM,
  'IWebPart.Contents': JSOM,
  'IWebPart.ListView': JSOM,
  'IWebPart.ListView.List': JSOM,
  'IWebPart.ListView.View': JSOM,
  'IWebPart.Order': JSOM,
  'IWebPart.PropertyOverrides': JSOM,
  'IWebPart.Title': JSOM,
  'IWebPart.Zone': JSOM,
  'IWebPartContents.FileSrc': JSOM,
  'IWebPartContents.Xml': JSOM,
  'IWebPartPropertyOverride.name': JSOM,
  'IWebPartPropertyOverride.type': JSOM,
  'IWebPartPropertyOverride.value': JSOM,
  'IListInstance.RemoveExistingViews':
    'inline templates in tests/integration/lists.test.ts (it acts only on lists made in the same setup)',
  'IListInstanceFieldReference.AdditionalProperties':
    'inline templates in tests/integration/lists.test.ts',
  'IListSecurity.ClearSubscopes':
    'tests/unit/lists.test.ts (SharePoint ignores it on a list that already has unique permissions)',
  'IWebSecurity.ClearSubscopes': 'tests/unit/security.test.ts',
  'IWebSecurity.CopyRoleAssignments':
    'inline templates in tests/integration/security.test.ts',
  'IListView.PersonalView':
    'inline templates in tests/integration/lists.test.ts (a personal view belongs to the account that runs the setup)',
  'IListView.AdditionalSettings.CustomFormatter': `${PASS_THROUGH}; the re-apply after a view is created is in tests/unit/lists.test.ts`,
  'IListView.AdditionalSettings.ViewType2': `${PASS_THROUGH}; the re-apply after a view is created is in tests/unit/lists.test.ts`,
  'IListView.AdditionalSettings.Hidden': PASS_THROUGH,
  'IListView.AdditionalSettings.JSLink': PASS_THROUGH,
  'IWebSettings.AlternateCssUrl': `${PASS_THROUGH} (tests/unit/websettings.test.ts)`,
  'IWebSettings.CustomMasterUrl': `${PASS_THROUGH} (tests/unit/websettings.test.ts)`,
  'IWebSettings.MasterUrl': `${PASS_THROUGH} (tests/unit/websettings.test.ts)`,
  'IWebSettings.RecycleBinEnabled': `${PASS_THROUGH} (tests/unit/websettings.test.ts)`,
  'IWebSettings.SaveSiteAsTemplateEnabled': `${PASS_THROUGH} (tests/unit/websettings.test.ts)`,
  'IWebSettings.SiteLogoUrl': `${PASS_THROUGH} (tests/unit/websettings.test.ts)`
}

// A list's AdditionalSettings are SharePoint's list properties, sent as they are.
for (const name of [
  'DefaultContentApprovalWorkflowId',
  'DefaultDisplayFormUrl',
  'DefaultEditFormUrl',
  'DefaultNewFormUrl',
  'Description',
  'Direction',
  'DocumentTemplateUrl',
  'DraftVersionVisibility',
  'EnableAttachments',
  'EnableFolderCreation',
  'EnableModeration',
  'ForceCheckout',
  'IrmEnabled',
  'IrmExpire',
  'IrmReject',
  'IsApplicationList',
  'NoCrawl',
  'OnQuickLaunch',
  'Title',
  'ValidationFormula',
  'ValidationMessage'
])
  NOT_COVERED[`IListInstance.AdditionalSettings.${name}`] = PASS_THROUGH

const root = join(__dirname, '..', '..')
const schema = JSON.parse(readFileSync(join(root, 'schema.json'), 'utf8'))
const definitions: Record<string, any> = schema.definitions

const refName = (ref: string) =>
  decodeURIComponent(ref.replace('#/definitions/', ''))

/**
 * Every property path the schema declares: `<Definition>.<property>`, with inline objects as
 * `<Definition>.<property>.<property>`.
 *
 * @param node - A schema node
 * @param label - Its path
 * @param into - Where the paths go
 */
function declared(node: any, label: string, into: Set<string>): Set<string> {
  for (const [key, child] of Object.entries<any>(node.properties || {})) {
    into.add(`${label}.${key}`)
    if (!child.$ref) declared(child, `${label}.${key}`, into)
    if (child.items && !child.items.$ref)
      declared(child.items, `${label}.${key}`, into)
  }
  for (const branch of node.anyOf || [])
    if (!branch.$ref) declared(branch, label, into)
  return into
}

/**
 * Records the property paths a template uses, following the schema.
 *
 * @param node - The schema node for the value
 * @param value - The value
 * @param label - The path of the node
 * @param into - Where the paths go
 */
function used(node: any, value: any, label: string, into: Set<string>): void {
  if (!node || value === null || value === undefined) return
  if (node.$ref) {
    const name = refName(node.$ref)
    return used(definitions[name], value, name, into)
  }
  if (node.anyOf) {
    for (const branch of node.anyOf) used(branch, value, label, into)
    return
  }
  if (Array.isArray(value)) {
    if (node.items)
      for (const item of value) used(node.items, item, label, into)
    return
  }
  if (typeof value !== 'object' || !node.properties) return
  for (const [key, child] of Object.entries(value)) {
    const property = node.properties[key]
    if (!property) continue
    into.add(`${label}.${key}`)
    used(property, child, `${label}.${key}`, into)
  }
}

describe('fixture coverage of the template contract', () => {
  const all = Object.entries(definitions).reduce(
    (paths, [name, definition]) => declared(definition, name, paths),
    declared(schema, 'Schema', new Set<string>())
  )
  const fixtures = join(root, 'tests', 'fixtures', 'templates')
  const covered = new Set<string>()
  for (const name of readdirSync(fixtures).filter((n) => n.endsWith('.json')))
    used(
      schema,
      JSON.parse(readFileSync(join(fixtures, name), 'utf8')),
      'Schema',
      covered
    )

  it('uses every property in a fixture or gives the reason it does not', () => {
    const missing = [...all].filter(
      (path) => !covered.has(path) && !NOT_COVERED[path]
    )
    assert.deepEqual(
      missing,
      [],
      'add these to a fixture, or to NOT_COVERED with a reason'
    )
  })

  it('lists no property as not covered that a fixture covers, or that does not exist', () => {
    const stale = Object.keys(NOT_COVERED).filter(
      (path) => covered.has(path) || !all.has(path)
    )
    assert.deepEqual(stale, [], 'remove these from NOT_COVERED')
  })
})
