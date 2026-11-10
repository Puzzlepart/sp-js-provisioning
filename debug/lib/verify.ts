import { SPInstance } from '@pnp/sp/presets/all'
import { DefaultHandlerSort } from '../../src/handlers/exports'
import { ProvisioningContext } from '../../src/provisioningcontext'
import {
  base64EncodeString,
  makeUrlRelative,
  mapStrings,
  replaceUrlTokens
} from '../../src/util'
import { declaredType, fieldTarget, guidOf } from '../../src/util/fieldxml'
import { normaliseNavigationUrl } from '../../src/util/navigationurl'
import { TokenHelper } from '../../src/util/tokenhelper'
import { handlerOrder } from './cli'
import { ISiteConnection, spfxContext } from './connect'
import { SHADE_NAMES, FONT_SLOTS } from './readback/composedlook'
import {
  bareGuid,
  IAssignmentState,
  IFieldState,
  IItemsState,
  IListState,
  IPrincipalState,
  ISecurableState,
  ISiteQuery,
  ISiteState,
  readSite
} from './readback/site'
import { fieldSchema, IFieldSchema } from './readback/sitefields'

/**
 * `verify`: a read-only diff of a template against a web (docs/specs/integration-testing.md,
 * Harness). `expectations()` turns a template into what the web should hold, replacing tokens only
 * where the engine replaces them (the token matrix: `TokenHelper` tokens, URL tokens or the
 * associated-group tokens, per handler and member); `diff()` compares that with what
 * `readSite()` read and returns findings, `{ handler, path, expected, actual }`. Both are pure, so
 * the rules are unit-tested (tests/unit/verify.test.ts).
 *
 * The rules follow what the handlers do on a re-apply:
 *
 * - lists, content types, views, folders, fields, group members and role assignments compare as a
 *   superset (what the template declares must be there; more may be), except with
 *   `RemoveExistingContentTypes`, `RemoveExistingViews` and `ClearExistingMembers`;
 * - data rows are found by their `KeyColumn` (its tokens replaced); a `Skip` row only has to exist,
 *   an `Overwrite` row is compared column by column;
 * - pages with `Overwrite: false` (or none) and custom actions only have to exist: the handlers
 *   leave existing ones as they are;
 * - navigation titles compare exactly, in order, at every level of a collection the template
 *   declares (the handler owns it); a node's URL only with `IgnoreExisting`, since a kept node keeps
 *   its stored URL;
 * - `exact` (for a fresh scratch web the template was just applied to) also compares the values of
 *   `Skip` rows, `Overwrite: false` pages and custom actions, every navigation URL, and requires site
 *   columns and content types on the web itself, not inherited from a parent web.
 */

/**
 * A difference between the template and the web.
 */
export interface IFinding {
  handler: string
  /** Where, e.g. `SPJS Itest Oppgaver/views/Aktive/fields`. */
  path: string
  expected: any
  actual: any
  /** Why it matters, or what the engine would have said. */
  message?: string
  /** Which template, when more than one was verified (`--mode project`). */
  source?: string
}

/**
 * Something the template declares that `verify` does not compare, and why.
 */
export interface INote {
  handler: string
  path: string
  reason: string
}

/**
 * The values tokens stand for on the web, read by {@link verifyContext}.
 */
export interface IVerifyContext {
  web: {
    Id: string
    Title: string
    ServerRelativeUrl: string
    Url: string
    Language?: number
  }
  site: { Id: string; Url: string; ServerRelativeUrl: string }
  /** `{sitecollectiontermstoreid}`, when the template uses it and the term store could be read. */
  termStoreId?: string
  /** List title to id, for `{listid:}`. */
  lists: Record<string, string>
  /** `List|View` to view id, for `{listviewid:}`. */
  listViews: Record<string, string>
  /** Values for `{parameter:}`. */
  parameters?: Record<string, string>
}

/**
 * How to verify.
 */
export interface IVerifyOptions {
  /** Compare everything, for a fresh web the template was just applied to (see the module). */
  exact?: boolean
  /** Only these sections (handler names); every section the template has by default. */
  handlers?: string[]
}

/**
 * Token replacement as the engine does it, per kind (the matrix in the spec): `t` is
 * `TokenHelper.replaceTokens`, `u` is `replaceUrlTokens`, `tu` is both (ClientSidePages
 * properties); `unresolved` names the known tokens a string still holds after `t`.
 */
export interface ITokens {
  t(text: string): string
  u(text: string): string
  tu(text: string): string
  unresolved(text: string): string[]
}

/**
 * The engine's token replacement for a web, with the values of a {@link IVerifyContext}.
 *
 * @param context - The values tokens stand for
 */
export function tokensOf(context: IVerifyContext): ITokens {
  const provisioning = new ProvisioningContext()
  provisioning.web = {
    Id: context.web.Id,
    Title: context.web.Title,
    ServerRelativeUrl: context.web.ServerRelativeUrl
  }
  provisioning.siteId = context.site.Id
  provisioning.termStoreId = context.termStoreId
  provisioning.lists = { ...context.lists }
  provisioning.listViews = { ...context.listViews }
  const config = {
    parameters: context.parameters || {},
    spfxContext: spfxContext(context.site, {
      Language: 0,
      ...context.web
    } as any)
  }
  const helper = new TokenHelper(provisioning, config)
  const t = (text: string) => helper.replaceTokens(text)
  const u = (text: string) =>
    typeof text === 'string' ? replaceUrlTokens(text, config) : text
  return {
    t,
    u,
    tu: (text: string) => u(t(text)),
    unresolved: (text: string) => helper.unresolvedTokens(text)
  }
}

/** A field of the template, as it should read back. */
export interface IExpectedField {
  /** Internal name (`Name`). */
  name: string
  /** `ID`, lower case without braces, when the XML has one. */
  id?: string
  /** The type SharePoint reports (`declaredType`: `Mult="TRUE"` on a Lookup is a LookupMulti). */
  type: string
  /** The XML with its tokens replaced. */
  xml: string
  schema: IFieldSchema
  /** Known tokens the XML still holds (no value on this web). */
  unresolved: string[]
}

/** A principal of a role assignment or a group owner, resolved as `resolvePrincipal` does. */
export type IExpectedPrincipal =
  | { kind: 'associated'; which: 'owner' | 'member' | 'visitor'; text: string }
  | { kind: 'id'; id: number; text: string }
  | { kind: 'name'; name: string; text: string }

export interface IExpectedAssignment {
  principal: IExpectedPrincipal
  role: string
}

export interface IExpectedSecurable {
  breakRoleInheritance: boolean
  assignments: IExpectedAssignment[]
}

export interface IExpectedGroup {
  title: string
  /** The group settings the template sets, tokens replaced in strings. */
  properties: Record<string, any>
  members: string[]
  clearExisting: boolean
  owner?: IExpectedPrincipal
}

export interface IExpectedFieldRef {
  id: string
  name?: string
  displayName?: string
  required?: boolean
  hidden?: boolean
  additional?: Record<string, string>
}

export interface IExpectedView {
  title: string
  fields?: string[]
  settings: Record<string, any>
}

export interface IExpectedRow {
  /** The row's place in the template. */
  index: number
  /** The `KeyColumn` value, tokens replaced as the handler matches it. */
  key?: any
  /** The values as the template gives them (Lookup, User, MultiChoice, DateTime, Boolean). */
  raw: Record<string, any>
  /** The values with tokens replaced in strings and URL values (Text, Number, URL, ...). */
  resolved: Record<string, any>
}

export interface IExpectedList {
  title: string
  template?: number
  description?: string
  contentTypesEnabled?: boolean
  settings: Record<string, any>
  bindings: string[]
  removeExistingContentTypes: boolean
  fields: IExpectedField[]
  fieldRefs: IExpectedFieldRef[]
  views: IExpectedView[]
  removeExistingViews: boolean
  /** Folder paths below the root folder (`Økonomi/Fakturaer`). */
  folders: string[]
  rows?: { keyColumn?: string; skip: boolean; rows: IExpectedRow[] }
  security?: IExpectedSecurable
}

export type IExpectedControl =
  | { type: 'text'; text: string }
  | {
      type: 'webpart'
      /** The template's control id, lower case: the handler matches a component id containing it. */
      id: string
      properties: Record<string, any>
      serverProcessedContent?: Record<string, any>
    }

export interface IExpectedPage {
  name: string
  overwrite: boolean
  title: string
  layout: string
  commentsDisabled: boolean
  sections: Array<{
    columns: Array<{ factor: number; controls: IExpectedControl[] }>
  }>
  verticalSection: IExpectedControl[] | null
}

export interface IExpectedNode {
  title: string
  /** The URL with the URL tokens replaced. */
  url: string
  /** Compare the URL: `IgnoreExisting`, or `exact`. */
  checkUrl: boolean
  children: IExpectedNode[]
}

export interface IExpectedContentType {
  id: string
  name: string
  description?: string
  group?: string
  fieldRefs: Array<{
    id: string
    name?: string
    required?: boolean
    hidden?: boolean
  }>
}

/**
 * What a template should leave on a web, per handler, with tokens replaced as the engine does.
 */
export interface IExpected {
  exact: boolean
  /** The sections compared, in the engine's run order. */
  handlers: string[]
  taxonomy?: {
    group?: { id: string; name: string }
    sets: Array<{
      id: string
      name: string
      terms: Array<{ id: string; name: string }>
    }>
  }
  siteFields?: IExpectedField[]
  contentTypes?: IExpectedContentType[]
  features?: Array<{ id: string; active: boolean }>
  security?: { groups: IExpectedGroup[]; web: IExpectedSecurable }
  lists?: IExpectedList[]
  files?: Array<{ url: string }>
  customActions?: Array<{ title: string; action: Record<string, any> }>
  composedLook?: { palette: string; fontScheme?: string; background?: string }
  pages?: IExpectedPage[]
  propertyBag?: Array<{
    key: string
    value: string
    overwrite: boolean
    indexed: boolean
  }>
  navigation?: {
    QuickLaunch?: IExpectedNode[]
    TopNavigationBar?: IExpectedNode[]
  }
  webSettings?: Record<string, any>
  /** What is declared but not compared, and why. */
  notes: INote[]
}

/** The engine's associated-group tokens (`resolvePrincipal`). */
const ASSOCIATED: Record<string, 'owner' | 'member' | 'visitor'> = {
  associatedownergroupid: 'owner',
  associatedmembergroupid: 'member',
  associatedvisitorgroupid: 'visitor'
}

/**
 * A role assignment principal or a group owner as `resolvePrincipal` resolves it, after
 * `TokenHelper` (associated-group tokens, an id, a site group title or a user login).
 *
 * @param principal - The template's principal
 * @param tokens - Token replacement
 */
function principalOf(
  principal: string | number,
  tokens: ITokens
): IExpectedPrincipal {
  if (typeof principal === 'number')
    return { kind: 'id', id: principal, text: String(principal) }
  const text = tokens.t(String(principal)).trim()
  if (/^\d+$/.test(text)) return { kind: 'id', id: Number(text), text }
  const which = ASSOCIATED[text.replace(/[{}]/g, '').toLowerCase()]
  if (which) return { kind: 'associated', which, text }
  return { kind: 'name', name: text, text }
}

/**
 * A securable's expectation (`Security` of the web or a list).
 *
 * @param security - The template's security
 * @param tokens - Token replacement
 */
function securableOf(security: any, tokens: ITokens): IExpectedSecurable {
  return {
    breakRoleInheritance: !!security.BreakRoleInheritance,
    assignments: (security.RoleAssignments || []).map((assignment: any) => ({
      principal: principalOf(assignment.Principal, tokens),
      role: String(assignment.RoleDefinition)
    }))
  }
}

/**
 * A field's expectation from its template XML (tokens replaced with `TokenHelper`).
 *
 * @param xml - The template's field XML
 * @param tokens - Token replacement
 */
function fieldOf(xml: string, tokens: ITokens): IExpectedField {
  const resolved = tokens.t(xml)
  const schema = fieldSchema(resolved)
  return {
    name: schema.attributes.Name,
    ...(schema.attributes.ID ? { id: bareGuid(schema.attributes.ID) } : {}),
    type: declaredType(schema.attributes),
    xml: resolved,
    schema,
    unresolved: tokens.unresolved(resolved)
  }
}

/**
 * The folder paths of a template's `Folders`, parents first.
 *
 * @param folders - The folders
 * @param parent - The parent's path
 */
function folderPaths(folders: any[], parent = ''): string[] {
  const paths: string[] = []
  for (const folder of folders || []) {
    if (!folder || !folder.Name) continue
    const path = parent ? `${parent}/${folder.Name}` : folder.Name
    paths.push(path, ...folderPaths(folder.Folders, path))
  }
  return paths
}

/**
 * A data row's values with tokens replaced where `Lists._buildItemValues` replaces them: in
 * strings and in a URL value's `Url` and `Description`.
 *
 * @param row - The row
 * @param tokens - Token replacement
 */
function resolvedRow(
  row: Record<string, any>,
  tokens: ITokens
): Record<string, any> {
  const resolved: Record<string, any> = {}
  for (const name of Object.keys(row)) {
    const value = row[name]
    if (typeof value === 'string') resolved[name] = tokens.t(value)
    else if (
      value &&
      typeof value === 'object' &&
      typeof value.Url === 'string'
    )
      resolved[name] = {
        Url: tokens.t(value.Url),
        ...(value.Description
          ? { Description: tokens.t(value.Description) }
          : {})
      }
    else resolved[name] = value
  }
  return resolved
}

/** The ids the ClientSidePages handler turns into a text control. */
const TEXT_IDS = ['Text', 'PageText']

/**
 * A page control's expectation: tokens (`TokenHelper` then URL) in each string of `Properties`,
 * `TokenHelper` only in `ServerProcessedContent`, none in the id (the handler's rules).
 *
 * @param control - The template's control
 * @param tokens - Token replacement
 */
function controlOf(control: any, tokens: ITokens): IExpectedControl {
  if (TEXT_IDS.includes(control.Id))
    return { type: 'text', text: control.Text ?? '' }
  const content = control.ServerProcessedContent || {}
  const filled = Object.keys(content).filter(
    (key) => content[key] && Object.keys(content[key]).length > 0
  )
  const expected: IExpectedControl = {
    type: 'webpart',
    id: String(control.Id).toLowerCase(),
    properties: mapStrings(control.Properties || {}, tokens.tu)
  }
  if (filled.length > 0) {
    expected.serverProcessedContent = {}
    for (const key of filled)
      expected.serverProcessedContent[key] = mapStrings(content[key], tokens.t)
  }
  return expected
}

/**
 * A navigation node's expectation: the URL tokens in `Url` only.
 *
 * @param node - The template's node
 * @param tokens - Token replacement
 * @param exact - Compare every URL
 */
function nodeOf(node: any, tokens: ITokens, exact: boolean): IExpectedNode {
  return {
    title: node.Title,
    url: tokens.u(node.Url ?? ''),
    checkUrl: exact || node.IgnoreExisting === true,
    children: (node.Children || []).map((child: any) =>
      nodeOf(child, tokens, exact)
    )
  }
}

/**
 * What a template should leave on a web. Pure: tokens are replaced with the values in `context`,
 * per the engine's matrix (see {@link ITokens}).
 *
 * @param schema - The template
 * @param context - The values tokens stand for on the web
 * @param options - `exact`, and which sections
 */
export function expectations(
  schema: Record<string, any>,
  context: IVerifyContext,
  options: IVerifyOptions = {}
): IExpected {
  const tokens = tokensOf(context)
  const exact = !!options.exact
  const handlers = handlerOrder(DefaultHandlerSort).filter(
    (name) =>
      schema[name] !== undefined &&
      (!options.handlers || options.handlers.includes(name))
  )
  const expected: IExpected = { exact, handlers, notes: [] }
  const note = (handler: string, path: string, reason: string) =>
    expected.notes.push({ handler, path, reason })
  const unresolvedNote = (handler: string, path: string, text: string) => {
    const left = tokens.unresolved(tokens.t(text))
    if (left.length > 0)
      note(
        handler,
        path,
        `${left.join(
          ', '
        )} has no value on this web (pass --parameters for {parameter:}); compared as it stands`
      )
  }

  for (const handler of handlers) {
    const section = schema[handler]
    switch (handler) {
      case 'Taxonomy':
        expected.taxonomy = {
          ...(section.TermGroup && section.TermGroup.Id
            ? {
                group: {
                  id: bareGuid(section.TermGroup.Id),
                  name: section.TermGroup.Name
                }
              }
            : {}),
          sets: (section.TermSets || []).map((set: any) => ({
            id: bareGuid(set.Id),
            name: set.Name,
            terms: (set.Terms || []).map((term: any) => ({
              id: bareGuid(term.Id),
              name: term.Name
            }))
          }))
        }
        break
      case 'SiteFields':
        expected.siteFields = []
        for (const xml of section as string[])
          try {
            expected.siteFields.push(fieldOf(xml, tokens))
          } catch (error) {
            note(
              handler,
              '?',
              `the field XML could not be read: ${error && error.message}`
            )
          }
        break
      case 'ContentTypes':
        expected.contentTypes = []
        for (const ct of section) {
          if (!ct.ID) {
            note(
              handler,
              ct.Name,
              'a content type without ID (matched by name by the handler) is not compared'
            )
            continue
          }
          expected.contentTypes.push({
            id: String(ct.ID).toUpperCase(),
            name: ct.Name,
            ...(ct.Description ? { description: ct.Description } : {}),
            ...(ct.Group ? { group: ct.Group } : {}),
            fieldRefs: (ct.FieldRefs || []).map((ref: any) => ({
              id: bareGuid(ref.ID),
              ...(ref.Name ? { name: ref.Name } : {}),
              ...(ref.Required !== undefined ? { required: ref.Required } : {}),
              ...(ref.Hidden !== undefined ? { hidden: ref.Hidden } : {})
            }))
          })
        }
        break
      case 'Features':
        expected.features = section.map((feature: any) => ({
          id: bareGuid(feature.id),
          active: !feature.deactivate
        }))
        break
      case 'Security':
        expected.security = {
          groups: (section.SiteGroups || []).map((group: any) => {
            const properties: Record<string, any> = {}
            for (const key of [
              'Description',
              'AllowMembersEditMembership',
              'AllowRequestToJoinLeave',
              'AutoAcceptRequestToJoinLeave',
              'OnlyAllowMembersViewMembership',
              'RequestToJoinLeaveEmailSetting'
            ])
              if (group[key] !== undefined)
                properties[key] =
                  typeof group[key] === 'string'
                    ? tokens.t(group[key])
                    : group[key]
            unresolvedNote(
              handler,
              `groups/${group.Title}`,
              JSON.stringify(group)
            )
            return {
              title: tokens.t(group.Title),
              properties,
              members: (group.Members || []).map((member: string) =>
                tokens.t(member)
              ),
              clearExisting: !!group.ClearExistingMembers,
              ...(group.Owner
                ? { owner: principalOf(group.Owner, tokens) }
                : {})
            }
          }),
          web: securableOf(section, tokens)
        }
        break
      case 'Lists':
        expected.lists = (section as any[]).map((lc) => {
          const views: IExpectedView[] = (lc.Views || []).map((view: any) => ({
            title: view.Title,
            ...(view.ViewFields ? { fields: view.ViewFields } : {}),
            settings: { ...(view.AdditionalSettings || {}) }
          }))
          const list: IExpectedList = {
            title: lc.Title,
            ...(lc.Template !== undefined ? { template: lc.Template } : {}),
            ...(lc.Description !== undefined
              ? { description: lc.Description }
              : {}),
            ...(lc.ContentTypesEnabled !== undefined
              ? { contentTypesEnabled: lc.ContentTypesEnabled }
              : {}),
            settings: { ...(lc.AdditionalSettings || {}) },
            bindings: (lc.ContentTypeBindings || []).map((binding: any) =>
              String(binding.ContentTypeID).toUpperCase()
            ),
            removeExistingContentTypes:
              !!lc.RemoveExistingContentTypes && !!lc.ContentTypeBindings,
            fields: [],
            // A field ref without an ID is noted below, not compared.
            fieldRefs: (lc.FieldRefs || [])
              .filter((ref: any) => !!ref.ID)
              .map((ref: any) => ({
                id: bareGuid(ref.ID),
                ...(ref.Name ? { name: ref.Name } : {}),
                ...(ref.DisplayName !== undefined
                  ? { displayName: ref.DisplayName }
                  : {}),
                ...(ref.Required !== undefined
                  ? { required: ref.Required }
                  : {}),
                ...(ref.Hidden !== undefined ? { hidden: ref.Hidden } : {}),
                ...(ref.AdditionalProperties
                  ? { additional: ref.AdditionalProperties }
                  : {})
              })),
            views,
            removeExistingViews: !!lc.RemoveExistingViews && !!lc.Views,
            folders: folderPaths(lc.Folders)
          }
          for (const ref of lc.FieldRefs || [])
            if (!ref.ID)
              note(
                handler,
                `${lc.Title}/FieldRefs/${ref.Name || '?'}`,
                'a field ref without an ID is never applied: the handler finds a list field ref by its ID, and logs an error for this one on every run'
              )
          for (const xml of lc.Fields || [])
            try {
              list.fields.push(fieldOf(xml, tokens))
            } catch (error) {
              note(
                handler,
                `${lc.Title}/Fields`,
                `a field XML could not be read: ${error && error.message}`
              )
            }
          const rows = lc.DataRows
          if (rows && rows.Rows && rows.Rows.length > 0) {
            list.rows = {
              ...(rows.KeyColumn ? { keyColumn: rows.KeyColumn } : {}),
              skip: rows.UpdateBehavior === 'Skip',
              rows: rows.Rows.map((row: any, index: number) => {
                const raw = row[rows.KeyColumn]
                return {
                  index,
                  ...(rows.KeyColumn
                    ? { key: typeof raw === 'string' ? tokens.t(raw) : raw }
                    : {}),
                  raw: row,
                  resolved: resolvedRow(row, tokens)
                }
              })
            }
            if (!rows.KeyColumn)
              note(
                handler,
                `${lc.Title}/DataRows`,
                'without a KeyColumn the handler adds every row on every run: each row must match an item, duplicates are not reported'
              )
            unresolvedNote(
              handler,
              `${lc.Title}/DataRows`,
              JSON.stringify(rows.Rows)
            )
          }
          if (lc.Security) list.security = securableOf(lc.Security, tokens)
          return list
        })
        break
      case 'Files':
        expected.files = (section as any[]).map((file) => ({
          url: `/${[context.web.ServerRelativeUrl, file.Folder, file.Url]
            .map((part) => String(part || '').replace(/^\/+|\/+$/g, ''))
            .filter(Boolean)
            .join('/')}`
        }))
        note(
          handler,
          '*',
          'only that each file exists: its contents, properties and web parts are not compared'
        )
        break
      case 'CustomActions':
        expected.customActions = (section as any[]).map((action) => ({
          title: action.Title,
          action
        }))
        break
      case 'ComposedLook':
        expected.composedLook = {
          palette: makeUrlRelative(tokens.u(section.ColorPaletteUrl)),
          ...(section.FontSchemeUrl
            ? { fontScheme: makeUrlRelative(tokens.u(section.FontSchemeUrl)) }
            : {}),
          ...(section.BackgroundImageUrl
            ? {
                background: makeUrlRelative(
                  tokens.u(section.BackgroundImageUrl)
                )
              }
            : {})
        }
        break
      case 'ClientSidePages':
        expected.pages = (section as any[]).map((page) => ({
          name: page.Name,
          overwrite: !!page.Overwrite,
          title: page.Title,
          layout: page.PageLayoutType,
          commentsDisabled: !!page.CommentsDisabled,
          sections: (page.Sections || []).map((s: any) => ({
            columns: (s.Columns || []).map((column: any) => ({
              factor: column.Factor,
              controls: (column.Controls || []).map((control: any) =>
                controlOf(control, tokens)
              )
            }))
          })),
          verticalSection: page.VerticalSection
            ? page.VerticalSection.map((control: any) =>
                controlOf(control, tokens)
              )
            : null
        }))
        break
      case 'PropertyBagEntries':
        expected.propertyBag = (section as any[]).map((entry) => ({
          key: entry.Key,
          value: entry.Value,
          overwrite: !!entry.Overwrite,
          indexed: !!entry.Indexed
        }))
        break
      case 'Navigation':
        expected.navigation = {}
        for (const collection of ['QuickLaunch', 'TopNavigationBar'] as const)
          if (Array.isArray(section[collection]))
            expected.navigation[collection] = section[
              collection
            ].map((node: any) => nodeOf(node, tokens, exact))
        break
      case 'WebSettings':
        expected.webSettings = {}
        for (const key of Object.keys(section))
          expected.webSettings[key] =
            typeof section[key] === 'string'
              ? tokens.u(section[key])
              : section[key]
        break
      case 'Hooks':
        note(
          handler,
          '*',
          'hooks call a URL and leave nothing on the web to read'
        )
        break
    }
  }
  return expected
}

/**
 * What to read from the web for an expectation.
 *
 * @param expected - The expectation
 */
export function queryOf(expected: IExpected): ISiteQuery {
  const query: ISiteQuery = {}
  if (expected.taxonomy)
    query.taxonomy = {
      ...(expected.taxonomy.group
        ? { groupId: expected.taxonomy.group.id }
        : {}),
      sets: expected.taxonomy.sets.map((set) => ({
        id: set.id,
        terms: set.terms.map((term) => term.id)
      }))
    }
  if (expected.siteFields)
    query.siteFields = expected.siteFields.map((f) => f.name)
  if (expected.contentTypes)
    query.contentTypes = expected.contentTypes.map((ct) => ct.id)
  if (expected.features) query.features = true
  if (expected.security)
    query.security = {
      groups: expected.security.groups.map((group) => group.title)
    }
  else if (
    (expected.lists || []).some(
      (list) =>
        list.security &&
        list.security.assignments.some(
          (assignment) => assignment.principal.kind === 'associated'
        )
    )
  )
    query.associatedGroups = true
  if (expected.lists)
    query.lists = expected.lists.map((list) => {
      const viewSettings: Record<string, string[]> = {}
      for (const view of list.views)
        viewSettings[view.title] = Object.keys(view.settings)
      const columns: string[] = []
      for (const row of list.rows ? list.rows.rows : [])
        for (const name of Object.keys(row.raw))
          if (!columns.includes(name)) columns.push(name)
      return {
        title: list.title,
        settings: Object.keys(list.settings),
        viewSettings,
        ...(list.rows ? { rowColumns: columns } : {})
      }
    })
  if (expected.files) query.files = expected.files.map((file) => file.url)
  if (expected.customActions) query.customActions = true
  if (expected.composedLook)
    query.composedLook = {
      palette: expected.composedLook.palette,
      ...(expected.composedLook.fontScheme
        ? { fontScheme: expected.composedLook.fontScheme }
        : {})
    }
  if (expected.pages) query.pages = expected.pages.map((page) => page.name)
  if (expected.propertyBag) query.propertyBag = true
  if (expected.navigation) query.navigation = true
  if (expected.webSettings)
    query.webSettings = Object.keys(expected.webSettings)
  return query
}

/** Field attributes SharePoint sets or rewrites itself: not compared. */
const MANAGED_ATTRIBUTES = [
  'ID',
  'Type',
  'Version',
  'SourceID',
  'ColName',
  'RowOrdinal',
  'Customization',
  'DisplayName'
]

/** What SharePoint manages on a taxonomy field (`SiteFields`' `TAXONOMY_MANAGED`). */
const TAXONOMY_MANAGED = ['List', 'WebId', 'ShowField']

/** The English role names `resolveRoleDefinitionId` resolves by role type. */
const WELL_KNOWN_ROLES: Record<string, number> = {
  'full control': 5,
  administrator: 5,
  design: 4,
  edit: 6,
  contribute: 3,
  read: 2
}

/**
 * Deep equality of plain data.
 *
 * @param a - A value
 * @param b - Another
 */
export function deepEqual(a: any, b: any): boolean {
  if (a === b) return true
  if (typeof a !== typeof b || a === null || b === null) return false
  if (Array.isArray(a) !== Array.isArray(b)) return false
  if (Array.isArray(a))
    return (
      a.length === b.length &&
      a.every((item, index) => deepEqual(item, b[index]))
    )
  if (typeof a === 'object') {
    const keys = Object.keys(a)
    return (
      keys.length === Object.keys(b).length &&
      keys.every(
        (key) =>
          Object.prototype.hasOwnProperty.call(b, key) &&
          deepEqual(a[key], b[key])
      )
    )
  }
  return false
}

/**
 * The part of `actual` that `expected` names: an object keeps only the expectation's keys (SharePoint
 * adds defaults), an array is reduced item by item (extra items kept whole), anything else as is.
 *
 * @param actual - What the web holds
 * @param expected - What the template says
 */
export function projected(actual: any, expected: any): any {
  if (Array.isArray(actual) && Array.isArray(expected))
    return actual.map((item, index) =>
      index < expected.length ? projected(item, expected[index]) : item
    )
  if (
    actual &&
    expected &&
    typeof actual === 'object' &&
    typeof expected === 'object' &&
    !Array.isArray(expected)
  ) {
    const copy: Record<string, any> = {}
    for (const key of Object.keys(expected))
      if (key in actual) copy[key] = projected(actual[key], expected[key])
    return copy
  }
  return actual
}

/**
 * Whether a principal is a user with a login (UPN, e-mail or claims) or e-mail address.
 *
 * @param principal - The user
 * @param login - The template's login
 */
export function loginMatches(
  principal: IPrincipalState | null,
  login: string
): boolean {
  if (!principal) return false
  const wanted = String(login).trim().toLowerCase()
  const name = (principal.LoginName || '').toLowerCase()
  return (
    name === wanted ||
    name.endsWith(`|${wanted}`) ||
    (principal.Email || '').toLowerCase() === wanted
  )
}

/**
 * Whether a principal on the web is the one the template names.
 *
 * @param expected - The template's principal
 * @param actual - A principal on the web
 * @param associated - The web's associated groups
 */
export function principalMatches(
  expected: IExpectedPrincipal,
  actual: IPrincipalState | null,
  associated: Record<string, IPrincipalState | null>
): boolean {
  if (!actual) return false
  if (expected.kind === 'id') return actual.Id === expected.id
  if (expected.kind === 'associated') {
    const group = associated[expected.which]
    return !!group && group.Id === actual.Id
  }
  return actual.Title === expected.name || loginMatches(actual, expected.name)
}

/**
 * Whether a role assignment grants a role definition, by name or, for the English names, by role
 * type (as `resolveRoleDefinitionId` resolves it).
 *
 * @param assignment - The assignment
 * @param role - The template's role definition
 */
export function grantsRole(
  assignment: IAssignmentState,
  role: string
): boolean {
  const wanted = role.trim().toLowerCase()
  return assignment.roles.some(
    (binding) =>
      String(binding.Name).toLowerCase() === wanted ||
      binding.RoleTypeKind === WELL_KNOWN_ROLES[wanted]
  )
}

/**
 * CAML or HTML as SharePoint stores it, for comparing: without the whitespace between tags, around
 * an attribute's `=` and before a tag's end (a template's `<FieldRef Name= "X"/>` is stored as
 * `<FieldRef Name="X" />`).
 */
const compact = (text: any) =>
  String(text ?? '')
    .replace(/>\s+</g, '><')
    .replace(/<[^<>]+>/g, (tag) =>
      tag
        .replace(/\s*=\s*(["'])/g, '=$1')
        .replace(/\s+/g, ' ')
        .replace(/\s*(\/?)>$/, (end, slash) => (slash ? ' />' : '>'))
    )
    .trim()

/** Text of HTML: tags dropped, whitespace collapsed (a Note column wraps its value). */
const textOfHtml = (html: any) =>
  String(html ?? '')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

/** A value that is not given (the engine leaves the column out). */
const absent = (value: any) => value === undefined || value === null

/**
 * Compares a template value with a setting SharePoint returns: equal, or equal as text (a number
 * the template wrote as a string).
 *
 * @param expected - The template's value
 * @param actual - The web's
 */
function sameSetting(expected: any, actual: any): boolean {
  return expected === actual || String(expected) === String(actual)
}

/**
 * The findings of a diff, with a helper to add one.
 */
class Findings {
  public list: IFinding[] = []
  constructor(public notes: INote[]) {}
  public add(
    handler: string,
    path: string,
    expected: any,
    actual: any,
    message?: string
  ) {
    this.list.push({
      handler,
      path,
      expected,
      actual,
      ...(message ? { message } : {})
    })
  }
  public note(handler: string, path: string, reason: string) {
    this.notes.push({ handler, path, reason })
  }
}

/**
 * Compares a field with what the template declares: id, type, title, every declared attribute
 * (but those SharePoint manages) and the `Default`, `Formula`, `CHOICES` and `FieldRefs` children;
 * a taxonomy field's term set and term store.
 *
 * @param f - The findings
 * @param handler - `SiteFields` or `Lists`
 * @param path - Where
 * @param expected - The template's field
 * @param actual - The web's
 */
function diffField(
  f: Findings,
  handler: string,
  path: string,
  expected: IExpectedField,
  actual: IFieldState
): void {
  const message = expected.unresolved.length
    ? `the template's ${expected.unresolved.join(
        ', '
      )} has no value on this web`
    : undefined
  if (expected.id && actual.Id !== expected.id)
    f.add(handler, `${path}/ID`, expected.id, actual.Id)
  if (expected.type && actual.TypeAsString !== expected.type)
    f.add(
      handler,
      `${path}/Type`,
      expected.type,
      actual.TypeAsString,
      'the handler applies only a widening type change'
    )
  const displayName = expected.schema.attributes.DisplayName
  if (displayName !== undefined && actual.Title !== displayName)
    f.add(handler, `${path}/DisplayName`, displayName, actual.Title)
  let stored: IFieldSchema
  try {
    stored = fieldSchema(actual.SchemaXml)
  } catch {
    f.note(handler, path, 'the stored SchemaXml could not be read')
    return
  }
  const taxonomy = /^TaxonomyFieldType/.test(expected.type || '')
  for (const name of Object.keys(expected.schema.attributes)) {
    if (MANAGED_ATTRIBUTES.includes(name)) continue
    if (taxonomy && TAXONOMY_MANAGED.includes(name)) continue
    const want = expected.schema.attributes[name]
    const have = stored.attributes[name]
    if (have === undefined && want === 'FALSE') continue
    if (want !== have) f.add(handler, `${path}/@${name}`, want, have, message)
  }
  for (const part of ['Default', 'Formula', 'Choices', 'FieldRefs'] as const)
    if (
      expected.schema[part] !== undefined &&
      !deepEqual(expected.schema[part], stored[part])
    )
      f.add(
        handler,
        `${path}/${part}`,
        expected.schema[part],
        stored[part],
        message
      )
  if (taxonomy)
    for (const target of ['TermSetId', 'SspId']) {
      const want = fieldTarget(expected.xml, target)
      if (want === undefined) continue
      const have = fieldTarget(actual.SchemaXml, target)
      if ((guidOf(want) || want) !== (guidOf(have) || have))
        f.add(
          handler,
          `${path}/${target}`,
          guidOf(want) || want,
          guidOf(have) || have,
          message
        )
    }
}

/**
 * Compares role inheritance and role assignments (a superset: the template's assignments must be
 * there, others may be).
 *
 * @param f - The findings
 * @param handler - `Security` or `Lists`
 * @param path - Where
 * @param expected - The template's security
 * @param actual - The web's or list's permissions
 * @param associated - The web's associated groups
 */
function diffSecurable(
  f: Findings,
  handler: string,
  path: string,
  expected: IExpectedSecurable,
  actual: ISecurableState,
  associated: Record<string, IPrincipalState | null>
): void {
  if (expected.breakRoleInheritance && !actual.hasUniqueRoleAssignments)
    f.add(handler, `${path}/HasUniqueRoleAssignments`, true, false)
  for (const assignment of expected.assignments) {
    const where = `${path}/RoleAssignments/${assignment.principal.text}`
    if (
      assignment.principal.kind === 'associated' &&
      !associated[assignment.principal.which]
    ) {
      f.add(
        handler,
        where,
        assignment.role,
        null,
        `the web has no associated ${assignment.principal.which} group of its own`
      )
      continue
    }
    const held = actual.roleAssignments.filter((candidate) =>
      principalMatches(assignment.principal, candidate.member, associated)
    )
    if (held.length === 0)
      f.add(
        handler,
        where,
        assignment.role,
        null,
        'the principal has no role assignment'
      )
    else if (!held.some((candidate) => grantsRole(candidate, assignment.role)))
      f.add(
        handler,
        where,
        assignment.role,
        held.reduce(
          (roles: string[], candidate) =>
            roles.concat(candidate.roles.map((r) => r.Name)),
          []
        )
      )
  }
}

/**
 * Compares one data row value with a list value, by the column's type (see `readItems`). Returns
 * `undefined` when they are the same, `{ unchecked }` when the value is not compared, or the two
 * sides.
 *
 * @param type - The column's `TypeAsString`
 * @param raw - The template's value as given
 * @param resolved - The template's value with tokens replaced (strings, URL values)
 * @param actual - The list value
 * @param expanded - Whether lookup and user values were expanded
 */
export function compareValue(
  type: string,
  raw: any,
  resolved: any,
  actual: any,
  expanded: boolean
): undefined | { unchecked: string } | { expected: any; actual: any } {
  const many = (value: any): any[] => (Array.isArray(value) ? value : [value])
  const differs = (expected: any, have: any) => ({ expected, actual: have })
  if (absent(raw)) return undefined
  if (/^Lookup/.test(type)) {
    const ids: number[] = []
    const texts: string[] = []
    for (const value of many(raw)) {
      if (typeof value === 'number') ids.push(value)
      else if (
        value &&
        typeof value === 'object' &&
        value.lookupId !== undefined
      )
        ids.push(Number(value.lookupId))
      else if (typeof value === 'string') texts.push(value)
      else if (
        value &&
        typeof value === 'object' &&
        value.lookupValue !== undefined
      )
        texts.push(String(value.lookupValue))
    }
    if (texts.length > 0 && !expanded)
      return {
        unchecked:
          'lookup values by text need the lookup expanded, which SharePoint refused'
      }
    const have = actual || { ids: [], texts: [] }
    const ok =
      have.ids.length === ids.length + texts.length &&
      ids.every((id) => have.ids.includes(id)) &&
      texts.every((text) => (have.texts || []).includes(text))
    return ok
      ? undefined
      : differs([...ids, ...texts], expanded ? have.texts : have.ids)
  }
  if (/^User/.test(type)) {
    const ids: number[] = []
    const logins: string[] = []
    for (const value of many(raw)) {
      if (typeof value === 'number') ids.push(value)
      else if (typeof value === 'string') logins.push(value)
      else if (value && (value.login || value.email))
        logins.push(value.login || value.email)
    }
    if (logins.length > 0 && !expanded)
      return {
        unchecked:
          'users by login need the column expanded, which SharePoint refused'
      }
    const have = actual || { ids: [], logins: [], emails: [] }
    const users: IPrincipalState[] = (have.ids || []).map(
      (id: number, index: number) => ({
        Id: id,
        Title: '',
        LoginName: (have.logins || [])[index] || '',
        Email: (have.emails || [])[index] || ''
      })
    )
    const ok =
      users.length === ids.length + logins.length &&
      ids.every((id) => users.some((user) => user.Id === id)) &&
      logins.every((login) => users.some((user) => loginMatches(user, login)))
    return ok
      ? undefined
      : differs(
          [...ids, ...logins],
          users.map((user) => user.LoginName || user.Id)
        )
  }
  if (/^TaxonomyFieldType/.test(type)) {
    const ids = many(raw)
      .filter((term) => term && term.termId)
      .map((term) => bareGuid(term.termId))
    const have = (actual && actual.termIds) || []
    return ids.length === have.length && ids.every((id) => have.includes(id))
      ? undefined
      : differs(ids, have)
  }
  if (type === 'URL') {
    const want =
      typeof resolved === 'string'
        ? { Url: resolved }
        : {
            Url: resolved.Url,
            Description: resolved.Description ?? resolved.Url
          }
    const have = actual || {}
    const ok =
      have.Url === want.Url &&
      (want.Description === undefined || have.Description === want.Description)
    return ok
      ? undefined
      : differs(
          want,
          actual && { Url: have.Url, Description: have.Description }
        )
  }
  if (type === 'MultiChoice') {
    const want = many(raw).map(String).sort()
    const have = (actual || []).map(String).sort()
    return deepEqual(want, have) ? undefined : differs(want, have)
  }
  if (type === 'Boolean') {
    const want = raw === true || raw === 'true' || raw === 1 || raw === '1'
    return want === actual ? undefined : differs(want, actual)
  }
  if (type === 'DateTime') {
    if (typeof raw !== 'string' || !/(Z|[+-]\d\d:?\d\d)$/i.test(raw))
      return {
        unchecked:
          "a DateTime without a time zone is stored in the web's time zone; not compared"
      }
    return actual && Date.parse(actual) === Date.parse(raw)
      ? undefined
      : differs(raw, actual)
  }
  if (['Number', 'Currency', 'Integer', 'Counter'].includes(type))
    return Number(resolved) === Number(actual)
      ? undefined
      : differs(resolved, actual)
  if (type === 'Note') {
    const want = textOfHtml(resolved)
    return want === textOfHtml(actual) ? undefined : differs(resolved, actual)
  }
  const want =
    typeof resolved === 'object' ? JSON.stringify(resolved) : String(resolved)
  return want === String(absent(actual) ? '' : actual)
    ? undefined
    : differs(resolved, actual)
}

/**
 * Compares a list's data rows with its items: a row with a key is the item whose key column holds
 * the key (one, not more); a `Skip` row only has to exist (unless `exact`); without a key, some item
 * must hold every value of the row.
 *
 * @param f - The findings
 * @param title - The list
 * @param rows - The template's rows
 * @param items - The list's items
 * @param exact - Compare the values of `Skip` rows too
 */
function diffRows(
  f: Findings,
  title: string,
  rows: NonNullable<IExpectedList['rows']>,
  items: IItemsState,
  exact: boolean
): void {
  for (const name of items.missing)
    f.add(
      'Lists',
      `${title}/DataRows/${name}`,
      'a column of the list',
      'no such column',
      'a row names it'
    )
  const rowDifferences = (
    row: IExpectedRow,
    item: { values: Record<string, any> },
    path: string
  ) => {
    const differences: IFinding[] = []
    for (const name of Object.keys(row.raw)) {
      if (!items.types[name]) continue
      const result = compareValue(
        items.types[name],
        row.raw[name],
        row.resolved[name],
        item.values[name],
        items.expanded
      )
      if (!result) continue
      if ('unchecked' in result)
        f.note('Lists', `${path}/${name}`, result.unchecked)
      else
        differences.push({
          handler: 'Lists',
          path: `${path}/${name}`,
          expected: result.expected,
          actual: result.actual
        })
    }
    return differences
  }
  for (const row of rows.rows) {
    const path = `${title}/DataRows/${row.index}`
    if (rows.keyColumn && !absent(row.key) && row.key !== '') {
      const type = items.types[rows.keyColumn]
      const matches = items.items.filter((item) => {
        const value = item.values[rows.keyColumn]
        return ['Number', 'Currency', 'Integer', 'Counter'].includes(type)
          ? Number(value) === Number(row.key)
          : String(absent(value) ? '' : value) === String(row.key)
      })
      const where = `${path} (${rows.keyColumn}=${row.key})`
      if (matches.length === 0) {
        f.add('Lists', where, 'an item', 'none')
        continue
      }
      if (matches.length > 1)
        f.add(
          'Lists',
          where,
          '1 item',
          `${matches.length} items`,
          'the handler updates only the first'
        )
      if (!rows.skip || exact)
        for (const difference of rowDifferences(row, matches[0], where))
          f.list.push(difference)
      continue
    }
    const differences = items.items.map((item) =>
      rowDifferences(row, item, path)
    )
    if (!differences.some((list) => list.length === 0))
      f.add('Lists', path, row.resolved, 'no item holds these values')
  }
}

/**
 * Compares a list with the template's: type, settings, content type bindings, fields, field refs,
 * views, folders, data rows and security.
 *
 * @param f - The findings
 * @param expected - The template's list
 * @param actual - The web's
 * @param associated - The web's associated groups
 * @param exact - See {@link IVerifyOptions}
 */
function diffList(
  f: Findings,
  expected: IExpectedList,
  actual: IListState,
  associated: Record<string, IPrincipalState | null>,
  exact: boolean
): void {
  const title = expected.title
  const properties = actual.properties
  if (
    expected.template !== undefined &&
    properties.BaseTemplate !== expected.template
  )
    f.add(
      'Lists',
      `${title}/Template`,
      expected.template,
      properties.BaseTemplate
    )
  if (
    expected.description !== undefined &&
    (properties.Description ?? '') !== expected.description
  )
    f.add(
      'Lists',
      `${title}/Description`,
      expected.description,
      properties.Description
    )
  if (
    expected.contentTypesEnabled !== undefined &&
    properties.ContentTypesEnabled !== expected.contentTypesEnabled
  )
    f.add(
      'Lists',
      `${title}/ContentTypesEnabled`,
      expected.contentTypesEnabled,
      properties.ContentTypesEnabled
    )
  for (const key of Object.keys(expected.settings)) {
    if (properties[key] === undefined) continue
    if (!sameSetting(expected.settings[key], properties[key]))
      f.add(
        'Lists',
        `${title}/AdditionalSettings/${key}`,
        expected.settings[key],
        properties[key]
      )
  }
  const bound = (id: string, binding: string) =>
    id.toUpperCase().indexOf(binding) === 0
  for (const binding of expected.bindings)
    if (!actual.contentTypes.some((ct) => bound(ct.Id, binding)))
      f.add(
        'Lists',
        `${title}/ContentTypeBindings/${binding}`,
        'bound',
        'not bound'
      )
  if (expected.removeExistingContentTypes) {
    const extra = actual.contentTypes.filter(
      (ct) =>
        !expected.bindings.some((binding) => bound(ct.Id, binding)) &&
        !bound(ct.Id, '0X0120')
    )
    if (extra.length > 0)
      f.add(
        'Lists',
        `${title}/ContentTypes`,
        'only the bound content types (RemoveExistingContentTypes)',
        extra.map((ct) => ct.Name)
      )
  }
  for (const field of expected.fields) {
    const path = `${title}/Fields/${field.name}`
    const stored = field.id
      ? actual.fields.find((candidate) => candidate.Id === field.id)
      : actual.fields.find((candidate) => candidate.InternalName === field.name)
    if (!stored) {
      const sameName = actual.fields.find(
        (candidate) => candidate.InternalName === field.name
      )
      f.add(
        'Lists',
        path,
        'a field',
        sameName ? `a field with another id (${sameName.Id})` : 'none',
        field.unresolved.length
          ? `the template's ${field.unresolved.join(
              ', '
            )} has no value on this web`
          : undefined
      )
      continue
    }
    diffField(f, 'Lists', path, field, stored)
  }
  for (const ref of expected.fieldRefs) {
    const path = `${title}/FieldRefs/${ref.name || ref.id}`
    const stored = actual.fields.find((candidate) => candidate.Id === ref.id)
    if (!stored) {
      f.add('Lists', path, 'a field', 'none')
      continue
    }
    if (ref.displayName !== undefined && stored.Title !== ref.displayName)
      f.add('Lists', `${path}/DisplayName`, ref.displayName, stored.Title)
    if (ref.required !== undefined && stored.Required !== ref.required)
      f.add('Lists', `${path}/Required`, ref.required, stored.Required)
    if (ref.hidden !== undefined && stored.Hidden !== ref.hidden)
      f.add('Lists', `${path}/Hidden`, ref.hidden, stored.Hidden)
    if (ref.additional) {
      let attributes: Record<string, string> = {}
      try {
        attributes = fieldSchema(stored.SchemaXml).attributes
      } catch {
        // compared as missing below
      }
      const wanted = fieldSchema(
        `<Field ${Object.keys(ref.additional)
          .map(
            (key) =>
              `${key}="${String(ref.additional[key])
                .replace(/&/g, '&amp;')
                .replace(/"/g, '&quot;')
                .replace(/</g, '&lt;')}"`
          )
          .join(' ')} />`
      ).attributes
      for (const key of Object.keys(wanted))
        if (attributes[key] !== wanted[key])
          f.add('Lists', `${path}/@${key}`, wanted[key], attributes[key])
    }
  }
  for (const view of expected.views) {
    const path = `${title}/views/${view.title}`
    const stored = actual.views.find(
      (candidate) => candidate.Title === view.title
    )
    if (!stored) {
      f.add('Lists', path, 'a view', 'none')
      continue
    }
    if (view.fields && !deepEqual(view.fields, stored.fields))
      f.add('Lists', `${path}/fields`, view.fields, stored.fields)
    for (const key of Object.keys(view.settings)) {
      const have = stored.properties[key]
      if (have === undefined) continue
      const want = view.settings[key]
      const same =
        key === 'ViewQuery'
          ? compact(want) === compact(have)
          : sameSetting(want, have)
      if (!same) f.add('Lists', `${path}/${key}`, want, have)
    }
  }
  if (expected.removeExistingViews) {
    const named = new Set(expected.views.map((view) => view.title))
    const root = `${actual.rootFolder}/`.toLowerCase()
    const extra = actual.views.filter(
      (view) =>
        !named.has(view.Title) &&
        !view.PersonalView &&
        !view.DefaultView &&
        !(
          view.Hidden &&
          view.ServerRelativeUrl.toLowerCase().indexOf(root) !== 0
        )
    )
    if (extra.length > 0)
      f.add(
        'Lists',
        `${title}/views`,
        "only the template's views (RemoveExistingViews)",
        extra.map((view) => view.Title)
      )
  }
  for (const folder of expected.folders)
    if (
      !actual.folders.some(
        (stored) => stored.toLowerCase() === folder.toLowerCase()
      )
    )
      f.add('Lists', `${title}/Folders/${folder}`, 'a folder', 'none')
  if (expected.rows && actual.items)
    diffRows(f, title, expected.rows, actual.items, exact)
  if (expected.security)
    diffSecurable(
      f,
      'Lists',
      `${title}/Security`,
      expected.security,
      actual,
      associated
    )
}

/**
 * Compares a page control with the template's.
 *
 * @param f - The findings
 * @param path - Where
 * @param expected - The template's control
 * @param actual - The page's
 */
function diffControl(
  f: Findings,
  path: string,
  expected: IExpectedControl,
  actual: any
): void {
  if (!actual) {
    f.add(
      'ClientSidePages',
      path,
      expected.type === 'text' ? 'a text control' : `web part ${expected.id}`,
      'none'
    )
    return
  }
  if (expected.type === 'text') {
    if (
      actual.type !== 'text' ||
      compact(actual.text) !== compact(expected.text)
    )
      f.add(
        'ClientSidePages',
        path,
        expected.text,
        actual.type === 'text' ? actual.text : actual.webPartId
      )
    return
  }
  if (
    actual.type !== 'webpart' ||
    !String(actual.webPartId || '').includes(expected.id)
  ) {
    f.add(
      'ClientSidePages',
      path,
      `web part ${expected.id}`,
      actual.type === 'webpart' ? `web part ${actual.webPartId}` : actual.type
    )
    return
  }
  const properties = projected(actual.properties || {}, expected.properties)
  if (!deepEqual(properties, expected.properties))
    f.add(
      'ClientSidePages',
      `${path}/properties`,
      expected.properties,
      properties
    )
  if (expected.serverProcessedContent) {
    const content = projected(
      actual.serverProcessedContent || {},
      expected.serverProcessedContent
    )
    if (!deepEqual(content, expected.serverProcessedContent))
      f.add(
        'ClientSidePages',
        `${path}/serverProcessedContent`,
        expected.serverProcessedContent,
        content
      )
  }
}

/**
 * Compares the nodes of one navigation level: the titles exactly and in order (the handler owns
 * the collection), then each node's URL where it is compared and its children.
 *
 * @param f - The findings
 * @param path - Where
 * @param expected - The template's nodes
 * @param actual - The stored nodes
 * @param base - The web URLs are resolved against
 */
function diffNodes(
  f: Findings,
  path: string,
  expected: IExpectedNode[],
  actual: Array<{ Title: string; Url: string; children: any[] }>,
  base: { serverRelativeUrl: string; origin: string }
): void {
  const want = expected.map((node) => node.title)
  const have = actual.map((node) => node.Title)
  if (!deepEqual(want, have)) f.add('Navigation', path, want, have)
  const used = actual.map(() => false)
  for (const node of expected) {
    const index = actual.findIndex(
      (stored, at) => !used[at] && stored.Title === node.title
    )
    if (index === -1) continue
    used[index] = true
    const stored = actual[index]
    if (
      node.checkUrl &&
      normaliseNavigationUrl(node.url, base) !==
        normaliseNavigationUrl(stored.Url, base)
    )
      f.add('Navigation', `${path}/${node.title}/Url`, node.url, stored.Url)
    diffNodes(
      f,
      `${path}/${node.title}`,
      node.children,
      stored.children || [],
      base
    )
  }
}

/**
 * Compares what a template should leave on a web with what the web holds. Pure.
 *
 * @param expected - From {@link expectations}
 * @param actual - From `readSite(conn, queryOf(expected))`
 * @param notes - Where values that are not compared are recorded
 */
export function diff(
  expected: IExpected,
  actual: ISiteState,
  notes: INote[] = []
): IFinding[] {
  const f = new Findings(notes)
  const exact = expected.exact
  const associated =
    (actual.security && actual.security.associated) ||
    actual.associatedGroups ||
    {}

  if (expected.taxonomy && actual.taxonomy) {
    const store = actual.taxonomy
    if (store.unavailable) f.note('Taxonomy', '*', store.unavailable)
    else {
      const group = expected.taxonomy.group
      if (group && !store.group)
        f.add('Taxonomy', `TermGroup/${group.name}`, 'a term group', 'none')
      else if (group && store.group.Name !== group.name)
        f.add(
          'Taxonomy',
          `TermGroup/${group.id}/Name`,
          group.name,
          store.group.Name
        )
      for (const set of expected.taxonomy.sets) {
        const stored = store.sets[set.id]
        if (!stored) {
          f.add('Taxonomy', `TermSets/${set.name}`, 'a term set', 'none')
          continue
        }
        if (!stored.names.includes(set.name))
          f.add('Taxonomy', `TermSets/${set.id}/Name`, set.name, stored.names)
        for (const term of set.terms) {
          const storedTerm = stored.terms[term.id]
          if (!storedTerm)
            f.add(
              'Taxonomy',
              `TermSets/${set.name}/Terms/${term.name}`,
              'a term',
              'none'
            )
          else if (!storedTerm.labels.includes(term.name))
            f.add(
              'Taxonomy',
              `TermSets/${set.name}/Terms/${term.id}/Name`,
              term.name,
              storedTerm.labels
            )
        }
      }
    }
  }

  for (const field of expected.siteFields || []) {
    const stored = (actual.siteFields || {})[field.name]
    if (!stored) {
      f.add(
        'SiteFields',
        field.name,
        'a site column',
        'none',
        field.unresolved.length
          ? `the template's ${field.unresolved.join(
              ', '
            )} has no value on this web`
          : undefined
      )
      continue
    }
    if (exact && !stored.onWeb)
      f.add('SiteFields', field.name, 'on this web', 'only in a parent web')
    diffField(f, 'SiteFields', field.name, field, stored)
  }

  for (const ct of expected.contentTypes || []) {
    const stored = (actual.contentTypes || {})[ct.id]
    if (!stored) {
      f.add('ContentTypes', `${ct.name} (${ct.id})`, 'a content type', 'none')
      continue
    }
    const path = `${ct.name} (${ct.id})`
    if (exact && !stored.onWeb)
      f.add('ContentTypes', path, 'on this web', 'only in a parent web')
    if (stored.Name !== ct.name)
      f.add('ContentTypes', `${path}/Name`, ct.name, stored.Name)
    if (ct.description !== undefined && stored.Description !== ct.description)
      f.add(
        'ContentTypes',
        `${path}/Description`,
        ct.description,
        stored.Description
      )
    if (ct.group !== undefined && stored.Group !== ct.group)
      f.add('ContentTypes', `${path}/Group`, ct.group, stored.Group)
    const positions: number[] = []
    for (const ref of ct.fieldRefs) {
      const index = stored.fieldLinks.findIndex(
        (link) => link.Id === ref.id || (!!ref.name && link.Name === ref.name)
      )
      const where = `${path}/FieldRefs/${ref.name || ref.id}`
      if (index === -1) {
        f.add('ContentTypes', where, 'a field link', 'none')
        continue
      }
      positions.push(index)
      const link = stored.fieldLinks[index]
      if (ref.required !== undefined && link.Required !== ref.required)
        f.add('ContentTypes', `${where}/Required`, ref.required, link.Required)
      if (ref.hidden !== undefined && link.Hidden !== ref.hidden)
        f.add('ContentTypes', `${where}/Hidden`, ref.hidden, link.Hidden)
    }
    if (
      exact &&
      positions.some(
        (position, index) => index > 0 && position < positions[index - 1]
      )
    )
      f.add(
        'ContentTypes',
        `${path}/FieldRefs`,
        ct.fieldRefs.map((ref) => ref.name || ref.id),
        stored.fieldLinks.map((link) => link.Name),
        'the handler orders the field links as the template does'
      )
  }

  for (const feature of expected.features || []) {
    const active = (actual.features || []).includes(feature.id)
    if (active !== feature.active)
      f.add(
        'Features',
        feature.id,
        feature.active ? 'active' : 'inactive',
        active ? 'active' : 'inactive'
      )
  }

  if (expected.security && actual.security) {
    for (const group of expected.security.groups) {
      const stored = actual.security.groups[group.title]
      const path = `groups/${group.title}`
      if (!stored) {
        f.add('Security', path, 'a site group', 'none')
        continue
      }
      for (const key of Object.keys(group.properties))
        if (!sameSetting(group.properties[key], (stored as any)[key]))
          f.add(
            'Security',
            `${path}/${key}`,
            group.properties[key],
            (stored as any)[key]
          )
      for (const member of group.members)
        if (!stored.members.some((user) => loginMatches(user, member)))
          f.add(
            'Security',
            `${path}/members/${member}`,
            'a member',
            'not a member'
          )
      if (group.clearExisting) {
        const extra = stored.members.filter(
          (user) => !group.members.some((member) => loginMatches(user, member))
        )
        if (extra.length > 0)
          f.add(
            'Security',
            `${path}/members`,
            group.members,
            stored.members.map((user) => user.LoginName),
            "ClearExistingMembers: only the template's members"
          )
      }
      if (
        group.owner &&
        !principalMatches(group.owner, stored.Owner, associated)
      )
        f.add(
          'Security',
          `${path}/Owner`,
          group.owner.text,
          stored.Owner ? stored.Owner.Title : null,
          group.owner.kind === 'associated' && !associated[group.owner.which]
            ? `the web has no associated ${group.owner.which} group of its own`
            : undefined
        )
    }
    diffSecurable(
      f,
      'Security',
      'web',
      expected.security.web,
      actual.security.web,
      associated
    )
  }

  for (const list of expected.lists || []) {
    const stored = (actual.lists || {})[list.title]
    if (!stored) {
      f.add('Lists', list.title, 'a list', 'none')
      continue
    }
    diffList(f, list, stored, associated, exact)
  }

  for (const file of expected.files || [])
    if (!(actual.files || {})[file.url])
      f.add('Files', file.url, 'a file', 'none')

  for (const action of expected.customActions || []) {
    const stored = (actual.customActions || []).filter(
      (candidate) => candidate.Title === action.title
    )
    if (stored.length === 0) {
      f.add('CustomActions', action.title, 'a custom action', 'none')
      continue
    }
    if (stored.length > 1)
      f.add(
        'CustomActions',
        action.title,
        '1 custom action',
        `${stored.length}`,
        'the handler adds an action whose title is not there'
      )
    if (!exact) continue
    const text = (value: any) => (absent(value) ? '' : String(value))
    for (const key of [
      'Name',
      'Description',
      'Location',
      'Sequence',
      'Url',
      'ClientSideComponentId',
      'ClientSideComponentProperties',
      'RegistrationId',
      'RegistrationType',
      'ScriptBlock',
      'ScriptSrc'
    ]) {
      const want = action.action[key]
      if (want === undefined) continue
      const have = (stored[0] as any)[key]
      let same =
        key === 'ClientSideComponentId'
          ? text(want).toLowerCase() === text(have)
          : sameSetting(text(want), text(have))
      if (!same && key === 'ClientSideComponentProperties')
        try {
          same = deepEqual(
            JSON.parse(text(want) || '{}'),
            JSON.parse(text(have) || '{}')
          )
        } catch {
          // not JSON: compared as text
        }
      if (!same) f.add('CustomActions', `${action.title}/${key}`, want, have)
    }
  }

  if (expected.composedLook && actual.composedLook) {
    const look = actual.composedLook
    if (!look.theme.themedCssFolderUrl)
      f.add('ComposedLook', 'ThemedCssFolderUrl', 'a theme', 'none')
    if (look.unreadable)
      f.note(
        'ComposedLook',
        'files',
        `the palette or font scheme could not be read: ${look.unreadable}`
      )
    for (const name of SHADE_NAMES)
      if (
        look.palette &&
        look.palette[name] !== undefined &&
        look.theme.shades[name] !== look.palette[name]
      )
        f.add(
          'ComposedLook',
          `shades/${name}`,
          look.palette[name],
          look.theme.shades[name]
        )
    for (const slot of FONT_SLOTS)
      if (
        look.fonts &&
        look.fonts[slot] !== undefined &&
        look.theme.fonts[slot] !== look.fonts[slot]
      )
        f.add(
          'ComposedLook',
          `fonts/${slot}`,
          look.fonts[slot],
          look.theme.fonts[slot]
        )
    const background = !!look.theme.backgroundImageUri
    if (background !== !!expected.composedLook.background)
      f.add(
        'ComposedLook',
        'BackgroundImageUrl',
        expected.composedLook.background || null,
        look.theme.backgroundImageUri
      )
  }

  for (const page of expected.pages || []) {
    const stored = (actual.pages || {})[page.name]
    if (!stored) {
      f.add('ClientSidePages', page.name, 'a page', 'none')
      continue
    }
    if (!page.overwrite && !exact) continue
    if (stored.title !== page.title)
      f.add('ClientSidePages', `${page.name}/Title`, page.title, stored.title)
    if (stored.layout !== page.layout)
      f.add(
        'ClientSidePages',
        `${page.name}/PageLayoutType`,
        page.layout,
        stored.layout
      )
    if (stored.commentsDisabled !== page.commentsDisabled)
      f.add(
        'ClientSidePages',
        `${page.name}/CommentsDisabled`,
        page.commentsDisabled,
        stored.commentsDisabled
      )
    if (stored.sections.length !== page.sections.length)
      f.add(
        'ClientSidePages',
        `${page.name}/Sections`,
        page.sections.length,
        stored.sections.length,
        'sections'
      )
    page.sections.forEach((section, s) => {
      const storedSection = stored.sections[s]
      if (!storedSection) return
      const path = `${page.name}/Sections/${s}`
      const factors = section.columns.map((column) => column.factor)
      const storedFactors = storedSection.columns.map((column) => column.factor)
      if (!deepEqual(factors, storedFactors))
        f.add('ClientSidePages', `${path}/Columns`, factors, storedFactors)
      section.columns.forEach((column, c) => {
        const storedColumn = storedSection.columns[c]
        if (!storedColumn) return
        if (storedColumn.controls.length !== column.controls.length)
          f.add(
            'ClientSidePages',
            `${path}/Columns/${c}/Controls`,
            column.controls.length,
            storedColumn.controls.length,
            'controls'
          )
        column.controls.forEach((control, k) =>
          diffControl(
            f,
            `${path}/Columns/${c}/Controls/${k}`,
            control,
            storedColumn.controls[k]
          )
        )
      })
    })
    if (page.verticalSection || stored.verticalSection) {
      const want = page.verticalSection || []
      const have = stored.verticalSection || []
      if (
        !page.verticalSection ||
        !stored.verticalSection ||
        want.length !== have.length
      )
        f.add(
          'ClientSidePages',
          `${page.name}/VerticalSection`,
          page.verticalSection ? want.length : null,
          stored.verticalSection ? have.length : null,
          'controls'
        )
      want.forEach((control, k) =>
        diffControl(f, `${page.name}/VerticalSection/${k}`, control, have[k])
      )
    }
  }

  for (const entry of expected.propertyBag || []) {
    const bag = actual.propertyBag || {}
    const have = bag[entry.key]
    if (absent(have)) {
      f.add(
        'PropertyBagEntries',
        entry.key,
        entry.overwrite ? entry.value : 'a value',
        null,
        entry.overwrite
          ? undefined
          : 'the handler writes only Overwrite: true entries'
      )
      continue
    }
    if (entry.overwrite && String(have) !== entry.value)
      f.add('PropertyBagEntries', entry.key, entry.value, have)
    if (entry.overwrite && entry.indexed) {
      const indexed = String(bag.vti_indexedpropertykeys || '').split('|')
      if (!indexed.includes(base64EncodeString(entry.key)))
        f.add('PropertyBagEntries', `${entry.key}/Indexed`, true, false)
    }
  }

  if (expected.navigation && actual.navigation)
    for (const collection of ['QuickLaunch', 'TopNavigationBar'] as const)
      if (expected.navigation[collection])
        diffNodes(
          f,
          collection,
          expected.navigation[collection],
          actual.navigation[collection],
          actual.navigation.base
        )

  if (expected.webSettings && actual.webSettings)
    for (const key of Object.keys(expected.webSettings)) {
      const have = actual.webSettings[key]
      if (have === undefined) continue
      const want = expected.webSettings[key]
      const same =
        key === 'WelcomePage'
          ? String(want).replace(/^\/+/, '').toLowerCase() ===
            String(have).replace(/^\/+/, '').toLowerCase()
          : /Url$/.test(key) && typeof want === 'string'
          ? makeUrlRelative(want).toLowerCase() ===
            makeUrlRelative(String(have ?? '')).toLowerCase()
          : sameSetting(want, have)
      if (!same) f.add('WebSettings', key, want, have)
    }

  return f.list
}

/**
 * The values tokens stand for on a web: its lists (`{listid:}`), the views a `{listviewid:}` names,
 * the term store id when the template uses it, the web and the site collection.
 *
 * @param conn - The web
 * @param schema - The template
 * @param parameters - Values for `{parameter:}`
 */
export async function verifyContext(
  conn: ISiteConnection,
  schema: Record<string, any>,
  parameters: Record<string, string> = {}
): Promise<IVerifyContext> {
  const text = JSON.stringify(schema)
  const lists: Record<string, string> = {}
  for (const list of await conn.web.lists.select('Id', 'Title').top(5000)<
    Array<{ Id: string; Title: string }>
  >())
    lists[list.Title] = list.Id
  let termStoreId: string | undefined
  if (/{(sitecollectiontermstoreid|termstoreid)}/.test(text))
    try {
      termStoreId = (
        await SPInstance([conn.web, conn.url], '_api/v2.1/termstore').select(
          'id'
        )<{
          id: string
        }>()
      ).id
    } catch {
      // Left unresolved, as in the engine (a note names it).
    }
  const listViews: Record<string, string> = {}
  const pattern = /{listviewid:([^{}|]*)\|([^{}]*)}/g
  for (let match = pattern.exec(text); match; match = pattern.exec(text)) {
    if (listViews[`${match[1]}|${match[2]}`] || !lists[match[1]]) continue
    for (const view of await conn.web.lists
      .getByTitle(match[1])
      .views.select('Id', 'Title')<Array<{ Id: string; Title: string }>>())
      listViews[`${match[1]}|${view.Title}`] = view.Id
  }
  return {
    web: conn.webInfo,
    site: conn.siteInfo,
    ...(termStoreId ? { termStoreId } : {}),
    lists,
    listViews,
    parameters
  }
}

/**
 * The outcome of verifying a template against a web.
 */
export interface IVerifyResult {
  /** The sections compared, in run order. */
  handlers: string[]
  findings: IFinding[]
  /** What was not compared, and why. */
  notes: INote[]
}

/**
 * Verifies a template against a web: reads the token values, builds the expectations, reads the
 * web and diffs. Read-only.
 *
 * @param conn - The web
 * @param schema - The template (not changed)
 * @param options - `exact`, the sections, values for `{parameter:}`
 */
export async function verify(
  conn: ISiteConnection,
  schema: Record<string, any>,
  options: IVerifyOptions & { parameters?: Record<string, string> } = {}
): Promise<IVerifyResult> {
  const context = await verifyContext(conn, schema, options.parameters)
  const expected = expectations(schema, context, options)
  const actual = await readSite(conn, queryOf(expected))
  const notes: INote[] = [...expected.notes, ...actual.unreadable]
  const findings = diff(expected, actual, notes)
  return { handlers: expected.handlers, findings, notes }
}

/**
 * A value for a finding line: JSON, shortened.
 *
 * @param value - The value
 */
function shown(value: any): string {
  const text = value === undefined ? 'undefined' : JSON.stringify(value)
  return text.length > 200 ? `${text.slice(0, 197)}...` : text
}

/**
 * A verify result as text lines: one per finding, one per note, and a summary.
 *
 * @param result - The result
 */
export function formatResult(result: IVerifyResult): string[] {
  const lines: string[] = []
  for (const finding of result.findings)
    lines.push(
      `FINDING [${finding.handler}]${
        finding.source ? ` (${finding.source})` : ''
      } ${finding.path}: expected ${shown(finding.expected)}, found ${shown(
        finding.actual
      )}${finding.message ? ` (${finding.message})` : ''}`
    )
  for (const note of result.notes)
    lines.push(`not checked [${note.handler}] ${note.path}: ${note.reason}`)
  lines.push(
    `${result.findings.length} finding(s) in ${
      result.handlers.join(', ') || '(no sections)'
    }${result.notes.length ? `; ${result.notes.length} not checked` : ''}.`
  )
  return lines
}
