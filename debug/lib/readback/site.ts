import { ContentTypes, SPInstance } from '@pnp/sp/presets/all'
import { ISiteConnection } from '../connect'
import {
  ICustomActionState,
  CUSTOM_ACTION_MEMBERS
} from './customactions'
import {
  IComposedLookState,
  readComposedLook,
  readFontScheme,
  readPalette
} from './composedlook'
import { IPageState, readClientSidePage } from './clientsidepages'
import { INavigationNodeState, readNavigationTree } from './navigation'

/**
 * Generic readers for what a template's handlers leave on a web, for `verify` (debug/lib/verify.ts)
 * and `inspect`: the same REST reads the suites' readers make, but for any template, not only the
 * fixtures (no prefix filters), and returning what SharePoint stores with as little shaping as
 * comparing needs. Read-only: nothing here writes to the site.
 */

/**
 * A field (site column or list field) as SharePoint stores it.
 */
export interface IFieldState {
  /** Lower case, without braces. */
  Id: string
  InternalName: string
  TypeAsString: string
  Title: string
  Required: boolean
  Hidden: boolean
  /** `SchemaXml`, as stored. */
  SchemaXml: string
  /** Site columns: the web's own (`web.fields`), not only available from a parent web. */
  onWeb?: boolean
  /** Lookups: the source list's show field. */
  LookupField?: string
}

/**
 * A user or group as a role assignment, a group's owner or a member names it.
 */
export interface IPrincipalState {
  Id: number
  Title: string
  LoginName: string
  Email?: string
  /** 1 user, 4 security group, 8 SharePoint group. */
  PrincipalType?: number
}

/**
 * A role assignment: the principal and its role definitions.
 */
export interface IAssignmentState {
  member: IPrincipalState
  roles: Array<{ Name: string; RoleTypeKind: number }>
}

/**
 * Permissions of a web or a list.
 */
export interface ISecurableState {
  hasUniqueRoleAssignments: boolean
  roleAssignments: IAssignmentState[]
}

/**
 * A site group with its owner and members.
 */
export interface IGroupState {
  Id: number
  Title: string
  Description: string
  Owner: IPrincipalState | null
  AllowMembersEditMembership: boolean
  AllowRequestToJoinLeave: boolean
  AutoAcceptRequestToJoinLeave: boolean
  OnlyAllowMembersViewMembership: boolean
  RequestToJoinLeaveEmailSetting: string
  members: IPrincipalState[]
}

/**
 * A list view with its fields and every property REST returns for it.
 */
export interface IViewState {
  Id: string
  Title: string
  PersonalView: boolean
  DefaultView: boolean
  Hidden: boolean
  ServerRelativeUrl: string
  /** The view as `views()` returns it (all its simple properties). */
  properties: Record<string, any>
  /** The view fields, in order. */
  fields: string[]
}

/**
 * A list value read for a data row, by the column's type:
 *
 * - Lookup and User columns: the ids, and the show field's values (lookups) or the users' login
 *   names and e-mail addresses, when they could be expanded;
 * - taxonomy columns: the term ids and labels;
 * - URL columns: `{ Url, Description }`;
 * - anything else: the value REST returns (MultiChoice as an array).
 */
export interface IItemState {
  Id: number
  values: Record<string, any>
}

/**
 * A list's items, with the columns the template's rows name.
 */
export interface IItemsState {
  /** `TypeAsString` by internal name, for the columns read. */
  types: Record<string, string>
  /** Columns a row names that the list does not have. */
  missing: string[]
  /** Whether lookup and user columns were expanded (their show field, login and e-mail). */
  expanded: boolean
  items: IItemState[]
}

/**
 * A list as `verify` and `inspect` read it.
 */
export interface IListState extends ISecurableState {
  /** The list as REST returns it without `$select`, plus the AdditionalSettings asked for. */
  properties: Record<string, any>
  /** List content types, in the list's order; ids upper case. */
  contentTypes: Array<{ Id: string; Name: string }>
  fields: IFieldState[]
  views: IViewState[]
  /** The root folder's server-relative URL. */
  rootFolder: string
  /** Every folder, as a path below the root folder. */
  folders: string[]
  items?: IItemsState
}

/**
 * A content type with its field links.
 */
export interface IContentTypeState {
  /** Upper case. */
  Id: string
  Name: string
  Description: string
  Group: string
  /** Whether the web itself has it (`web.contentTypes`), not only a parent web. */
  onWeb?: boolean
  fieldLinks: Array<{
    /** Lower case, without braces. */
    Id: string
    Name: string
    Required: boolean
    Hidden: boolean
  }>
}

/**
 * A term set and the terms asked for (`null` for one that is not there).
 */
export interface ITermSetState {
  Id: string
  names: string[]
  terms: Record<string, { Id: string; labels: string[] } | null>
}

/**
 * What the Taxonomy handler leaves in the term store, or why it could not be read.
 */
export interface ITaxonomyState {
  /** Why the term store could not be read (a 403 without `TermStore.Read.All`). */
  unavailable?: string
  group?: { Id: string; Name: string } | null
  sets: Record<string, ITermSetState | null>
}

/**
 * What a web holds of what a template asks about. Each member is read only when the query names it.
 */
export interface ISiteState {
  /** Available site columns by internal name (`onWeb` for the web's own). */
  siteFields?: Record<string, IFieldState>
  /** By id (upper case), `null` when not available to the web. */
  contentTypes?: Record<string, IContentTypeState | null>
  /** Ids of the active web features, lower case. */
  features?: string[]
  security?: {
    /** By title, `null` when the site collection has no such group. */
    groups: Record<string, IGroupState | null>
    web: ISecurableState
    /** The web's associated groups, `null` where it has none. */
    associated: Record<'owner' | 'member' | 'visitor', IPrincipalState | null>
  }
  /**
   * The web's associated groups, read without `security` when only a list's permissions name them
   * (`{associatedownergroupid}` and the others).
   */
  associatedGroups?: Record<'owner' | 'member' | 'visitor', IPrincipalState | null>
  /** By title, `null` when the web has no such list. */
  lists?: Record<string, IListState | null>
  customActions?: ICustomActionState[]
  composedLook?: {
    theme: IComposedLookState
    /** The palette file's slots, as `ThemeInfo` reports them once applied. */
    palette?: Record<string, string>
    /** The font scheme file's Latin typefaces. */
    fonts?: Record<string, string>
    /** Why a file could not be read. */
    unreadable?: string
  }
  /** By file name in `SitePages`, `null` when there is no such page. */
  pages?: Record<string, IPageState | null>
  navigation?: {
    /** The web the URLs are resolved against. */
    base: { serverRelativeUrl: string; origin: string }
    QuickLaunch: Array<INavigationNodeState<number>>
    TopNavigationBar: Array<INavigationNodeState<number>>
  }
  /** The web's settings by key (`WelcomePage` from the root folder). */
  webSettings?: Record<string, any>
  /** The web's property bag, keys decoded. */
  propertyBag?: Record<string, any>
  /** Whether each file exists, by server-relative URL. */
  files?: Record<string, boolean>
  taxonomy?: ITaxonomyState
  /** What could not be read, by `<handler>/<path>`, with the reason. */
  unreadable: Array<{ handler: string; path: string; reason: string }>
}

/**
 * What to read from a web.
 */
export interface ISiteQuery {
  /** Internal names; every available column is read once. */
  siteFields?: string[]
  /** Content type ids. */
  contentTypes?: string[]
  features?: boolean
  security?: { groups: string[] }
  /** The web's associated groups (`security` reads them too). */
  associatedGroups?: boolean
  lists?: Array<{
    title: string
    /** AdditionalSettings keys to read when the list's default JSON leaves them out. */
    settings?: string[]
    /** View AdditionalSettings keys, by view title. */
    viewSettings?: Record<string, string[]>
    /** Columns the data rows name: the items are read with these. */
    rowColumns?: string[]
  }>
  customActions?: boolean
  composedLook?: { palette?: string; fontScheme?: string }
  /** File names in `SitePages`. */
  pages?: string[]
  navigation?: boolean
  /** Web setting keys (`WelcomePage` included). */
  webSettings?: string[]
  propertyBag?: boolean
  /** Server-relative file URLs. */
  files?: string[]
  taxonomy?: { groupId?: string; sets: Array<{ id: string; terms: string[] }> }
}

/** A GUID without braces, lower case. */
export const bareGuid = (guid?: string): string =>
  String(guid ?? '')
    .replace(/[{}]/g, '')
    .toLowerCase()

/**
 * Whether a request failed because the object is not there.
 *
 * @param error - What the request rejected with
 */
const notFound = (error: any): boolean =>
  !!error && (error.status === 404 || error.status === 400)

/**
 * The message of what a request rejected with.
 *
 * @param error - The error
 */
const messageOf = (error: any): string =>
  `${(error && error.message) || error}`.split('\n')[0]

/**
 * A property bag key as REST encodes it (`vti_x005f_indexedpropertykeys`), decoded.
 *
 * @param key - The encoded key
 */
export function decodePropertyKey(key: string): string {
  return key.replace(/_x([\da-f]{4})_/gi, (match, code: string) =>
    String.fromCharCode(parseInt(code, 16))
  )
}

/**
 * A principal as REST returns it (`Member`, `Owner`, a group's `users`).
 *
 * @param principal - The principal
 */
export function principalOf(principal: any): IPrincipalState | null {
  if (!principal || typeof principal.Id !== 'number') return null
  return {
    Id: principal.Id,
    Title: principal.Title ?? '',
    LoginName: principal.LoginName ?? '',
    ...(principal.Email ? { Email: principal.Email } : {}),
    ...(principal.PrincipalType !== undefined
      ? { PrincipalType: principal.PrincipalType }
      : {})
  }
}

/**
 * Role assignments as REST returns them, expanded with `Member` and `RoleDefinitionBindings`.
 *
 * @param assignments - The assignments
 */
export function assignmentsOf(assignments: any[]): IAssignmentState[] {
  return (assignments || []).map((assignment) => ({
    member: principalOf(assignment.Member),
    roles: (assignment.RoleDefinitionBindings || []).map((binding: any) => ({
      Name: binding.Name,
      RoleTypeKind: binding.RoleTypeKind
    }))
  }))
}

/**
 * A field as REST returns it.
 *
 * @param field - The field
 */
export function fieldStateOf(field: any): IFieldState {
  return {
    Id: bareGuid(field.Id),
    InternalName: field.InternalName,
    TypeAsString: field.TypeAsString,
    Title: field.Title,
    Required: !!field.Required,
    Hidden: !!field.Hidden,
    SchemaXml: field.SchemaXml || '',
    ...(field.LookupField ? { LookupField: field.LookupField } : {})
  }
}

/** The members of a field the readers select. */
export const FIELD_MEMBERS = [
  'Id',
  'InternalName',
  'TypeAsString',
  'Title',
  'Required',
  'Hidden',
  'SchemaXml'
]

/**
 * Reads the site columns available to the web (its own and its parents'), by internal name.
 *
 * @param conn - The web
 */
export async function readSiteColumns(
  conn: ISiteConnection
): Promise<Record<string, IFieldState>> {
  const [available, own] = await Promise.all([
    conn.web.availablefields.select(...FIELD_MEMBERS)<any[]>(),
    conn.web.fields.select('InternalName')<Array<{ InternalName: string }>>()
  ])
  const ownNames = new Set(own.map((field) => field.InternalName))
  const columns: Record<string, IFieldState> = {}
  for (const field of available)
    columns[field.InternalName] = {
      ...fieldStateOf(field),
      onWeb: ownNames.has(field.InternalName)
    }
  return columns
}

/**
 * Reads the permissions of a web or list.
 *
 * @param securable - Returns the web or the list (a fresh queryable on every call)
 */
async function readSecurable(securable: () => any): Promise<ISecurableState> {
  const [{ HasUniqueRoleAssignments }, assignments] = await Promise.all([
    securable().select('HasUniqueRoleAssignments')() as Promise<{
      HasUniqueRoleAssignments: boolean
    }>,
    securable().roleAssignments.expand('Member', 'RoleDefinitionBindings')()
  ])
  return {
    hasUniqueRoleAssignments: HasUniqueRoleAssignments,
    roleAssignments: assignmentsOf(assignments)
  }
}

/**
 * Reads a site group with its owner and members, or `null` when there is none with the title.
 *
 * @param conn - The web
 * @param title - The group's title
 */
export async function readGroup(
  conn: ISiteConnection,
  title: string
): Promise<IGroupState | null> {
  const group = () => conn.web.siteGroups.getByName(title.replace(/'/g, "''"))
  let info: any
  try {
    info = await group()
      .expand('Owner')
      .select(
        'Id',
        'Title',
        'Description',
        'Owner/Id',
        'Owner/Title',
        'Owner/LoginName',
        'Owner/PrincipalType',
        'AllowMembersEditMembership',
        'AllowRequestToJoinLeave',
        'AutoAcceptRequestToJoinLeave',
        'OnlyAllowMembersViewMembership',
        'RequestToJoinLeaveEmailSetting'
      )()
  } catch (error) {
    if (notFound(error)) return null
    throw error
  }
  const users = await group().users.select(
    'Id',
    'Title',
    'LoginName',
    'Email',
    'PrincipalType'
  )<any[]>()
  return {
    Id: info.Id,
    Title: info.Title,
    Description: info.Description ?? '',
    Owner: principalOf(info.Owner),
    AllowMembersEditMembership: info.AllowMembersEditMembership,
    AllowRequestToJoinLeave: info.AllowRequestToJoinLeave,
    AutoAcceptRequestToJoinLeave: info.AutoAcceptRequestToJoinLeave,
    OnlyAllowMembersViewMembership: info.OnlyAllowMembersViewMembership,
    RequestToJoinLeaveEmailSetting: info.RequestToJoinLeaveEmailSetting ?? '',
    members: users.map(principalOf).filter(Boolean)
  }
}

/**
 * Reads the web's associated owner, member and visitor groups (`null` where it has none: a web
 * that inherits permissions answers `odata.null`).
 *
 * @param conn - The web
 */
export async function readAssociatedGroups(
  conn: ISiteConnection
): Promise<Record<'owner' | 'member' | 'visitor', IPrincipalState | null>> {
  const read = (group: any) =>
    group
      .select('Id', 'Title', 'LoginName')()
      .then((info: any) => principalOf(info))
      .catch(() => null)
  const [owner, member, visitor] = await Promise.all([
    read(conn.web.associatedOwnerGroup),
    read(conn.web.associatedMemberGroup),
    read(conn.web.associatedVisitorGroup)
  ])
  return { owner, member, visitor }
}

/**
 * Reads a list's views with their fields.
 *
 * @param list - Returns the list (a fresh queryable on every call)
 */
async function readViews(list: () => any): Promise<IViewState[]> {
  const views: any[] = await list().views()
  const states: IViewState[] = []
  for (const view of views) {
    const fields: { Items?: string[] } = await list()
      .views.getById(view.Id)
      .fields()
    states.push({
      Id: bareGuid(view.Id),
      Title: view.Title,
      PersonalView: !!view.PersonalView,
      DefaultView: !!view.DefaultView,
      Hidden: !!view.Hidden,
      ServerRelativeUrl: view.ServerRelativeUrl || '',
      properties: view,
      fields: fields.Items || []
    })
  }
  return states
}

/**
 * Reads the items of a list with the columns data rows name, each value shaped by the column's
 * type (see {@link IItemState}). Lookup and user columns are expanded for their show field, login
 * and e-mail; when SharePoint refuses that (too many lookups), they are read as ids only.
 *
 * @param list - Returns the list (a fresh queryable on every call)
 * @param fields - The list's fields
 * @param columns - The columns the rows name
 */
export async function readItems(
  list: () => any,
  fields: IFieldState[],
  columns: string[]
): Promise<IItemsState> {
  const byName = new Map(fields.map((field) => [field.InternalName, field]))
  const types: Record<string, string> = {}
  const missing: string[] = []
  for (const name of columns) {
    const field = byName.get(name)
    if (field) types[name] = field.TypeAsString
    else missing.push(name)
  }
  const names = Object.keys(types)
  const kind = (name: string) =>
    /^Lookup/.test(types[name])
      ? 'lookup'
      : /^User/.test(types[name])
      ? 'user'
      : 'plain'
  const query = (expanded: boolean) => {
    const select = ['Id']
    const expand: string[] = []
    for (const name of names) {
      if (kind(name) === 'plain') {
        select.push(name)
        continue
      }
      select.push(`${name}Id`)
      if (!expanded) continue
      expand.push(name)
      if (kind(name) === 'lookup')
        select.push(`${name}/${byName.get(name).LookupField || 'Title'}`)
      else select.push(`${name}/Name`, `${name}/EMail`)
    }
    let items = list()
      .items.select(...select)
      .top(5000)
    if (expand.length > 0) items = items.expand(...expand)
    return items() as Promise<any[]>
  }
  let expanded = names.some((name) => kind(name) !== 'plain')
  let rows: any[]
  try {
    rows = await query(expanded)
  } catch (error) {
    if (!expanded) throw error
    expanded = false
    rows = await query(false)
  }
  const many = (value: any): any[] =>
    value === null || value === undefined
      ? []
      : Array.isArray(value)
      ? value
      : value.results || [value]
  const items = rows.map((row) => {
    const values: Record<string, any> = {}
    for (const name of names) {
      const type = types[name]
      if (kind(name) === 'lookup') {
        const show = byName.get(name).LookupField || 'Title'
        values[name] = {
          ids: many(row[`${name}Id`]),
          ...(expanded
            ? {
                texts: many(row[name]).map((value: any) =>
                  String(value && value[show])
                )
              }
            : {})
        }
      } else if (kind(name) === 'user') {
        values[name] = {
          ids: many(row[`${name}Id`]),
          ...(expanded
            ? {
                logins: many(row[name]).map((user: any) => user && user.Name),
                emails: many(row[name]).map((user: any) => user && user.EMail)
              }
            : {})
        }
      } else if (/^TaxonomyFieldType/.test(type)) {
        const terms = many(row[name])
        values[name] = {
          termIds: terms.map((term: any) => bareGuid(term && term.TermGuid)),
          labels: terms.map((term: any) => term && term.Label)
        }
      } else if (type === 'MultiChoice') {
        values[name] = many(row[name])
      } else {
        values[name] = row[name] === undefined ? null : row[name]
      }
    }
    return { Id: row.Id, values }
  })
  return { types, missing, expanded, items }
}

/**
 * Reads a list as `verify` compares it, or `null` when the web has no list with the title.
 *
 * @param conn - The web
 * @param title - The list's title
 * @param options - Settings to read when the default JSON leaves them out, view settings, and the
 * columns data rows name
 * @param unreadable - Where what cannot be read is recorded
 */
export async function readListState(
  conn: ISiteConnection,
  title: string,
  options: {
    settings?: string[]
    viewSettings?: Record<string, string[]>
    rowColumns?: string[]
  } = {},
  unreadable: ISiteState['unreadable'] = []
): Promise<IListState | null> {
  const list = () => conn.web.lists.getByTitle(title)
  let properties: Record<string, any>
  try {
    properties = await list()()
  } catch (error) {
    if (notFound(error)) return null
    throw error
  }
  const absent = (options.settings || []).filter(
    (key) => properties[key] === undefined
  )
  for (const key of absent)
    try {
      Object.assign(properties, await list().select(key)())
    } catch (error) {
      unreadable.push({
        handler: 'Lists',
        path: `${title}/AdditionalSettings/${key}`,
        reason: `not readable through REST (${messageOf(error)})`
      })
    }
  const [contentTypes, fields, root, folderItems, securable] = await Promise.all([
    list().contentTypes.select('Id', 'Name')<any[]>(),
    list().fields.select(...FIELD_MEMBERS, 'LookupField')<any[]>(),
    list().rootFolder.select('ServerRelativeUrl')<{ ServerRelativeUrl: string }>(),
    list()
      .items.filter('FSObjType eq 1')
      .select('FileRef')
      .top(5000)<Array<{ FileRef: string }>>(),
    readSecurable(list)
  ])
  const views = await readViews(list)
  for (const view of views) {
    const keys = (options.viewSettings || {})[view.Title] || []
    for (const key of keys.filter((k) => view.properties[k] === undefined))
      try {
        Object.assign(
          view.properties,
          await list().views.getById(view.Id).select(key)()
        )
      } catch (error) {
        unreadable.push({
          handler: 'Lists',
          path: `${title}/views/${view.Title}/${key}`,
          reason: `not readable through REST (${messageOf(error)})`
        })
      }
  }
  const rootUrl = root.ServerRelativeUrl.replace(/\/+$/, '')
  const fieldStates = fields.map(fieldStateOf)
  return {
    properties,
    contentTypes: contentTypes.map((ct) => ({
      Id: String(ct.Id.StringValue).toUpperCase(),
      Name: ct.Name
    })),
    fields: fieldStates,
    views,
    rootFolder: rootUrl,
    folders: folderItems
      .map((item) =>
        item.FileRef.startsWith(`${rootUrl}/`)
          ? item.FileRef.slice(rootUrl.length + 1)
          : item.FileRef
      )
      .sort(),
    ...securable,
    ...(options.rowColumns
      ? { items: await readItems(list, fieldStates, options.rowColumns) }
      : {})
  }
}

/**
 * Reads a content type available to the web, with its field links, or `null` when there is none.
 *
 * @param conn - The web
 * @param id - The content type id
 */
export async function readContentType(
  conn: ISiteConnection,
  id: string
): Promise<IContentTypeState | null> {
  let info: any
  try {
    info = await ContentTypes(conn.web, 'availablecontenttypes')
      .getById(id)
      .select('Id', 'Name', 'Description', 'Group')
      .expand('FieldLinks')()
  } catch (error) {
    if (notFound(error)) return null
    throw error
  }
  if (!info || !info.Id) return null
  const onWeb = await conn.web.contentTypes
    .getById(id)
    .select('Id')()
    .then((own: any) => !!(own && own.Id))
    .catch(() => false)
  return {
    Id: String(info.Id.StringValue).toUpperCase(),
    Name: info.Name,
    Description: info.Description ?? '',
    Group: info.Group ?? '',
    onWeb,
    fieldLinks: (info.FieldLinks || []).map((link: any) => ({
      Id: bareGuid(link.Id),
      Name: link.Name,
      Required: !!link.Required,
      Hidden: !!link.Hidden
    }))
  }
}

/**
 * Reads term sets, their terms and a term group from the default term store
 * (`_api/v2.1/termstore`). A refused read (no `TermStore.Read.All`) is reported, not thrown.
 *
 * @param conn - The web
 * @param query - The group and the term sets with their terms
 */
export async function readTaxonomy(
  conn: ISiteConnection,
  query: NonNullable<ISiteQuery['taxonomy']>
): Promise<ITaxonomyState> {
  const get = (path: string) =>
    SPInstance([conn.web, conn.url], `_api/v2.1/termstore/${path}`)()
  const state: ITaxonomyState = { sets: {} }
  try {
    if (query.groupId) {
      try {
        const group: any = await get(`groups/${bareGuid(query.groupId)}`)
        state.group = { Id: bareGuid(group.id), Name: group.displayName }
      } catch (error) {
        if (error && error.status === 404) state.group = null
        else throw error
      }
    }
    for (const set of query.sets) {
      let info: any
      try {
        info = await get(`sets/${bareGuid(set.id)}`)
      } catch (error) {
        if (error && error.status === 404) {
          state.sets[set.id] = null
          continue
        }
        throw error
      }
      const terms: ITermSetState['terms'] = {}
      for (const termId of set.terms)
        try {
          const term: any = await get(
            `sets/${bareGuid(set.id)}/terms/${bareGuid(termId)}`
          )
          terms[termId] = {
            Id: bareGuid(term.id),
            labels: (term.labels || []).map((label: any) => label.name)
          }
        } catch (error) {
          if (error && error.status === 404) terms[termId] = null
          else throw error
        }
      state.sets[set.id] = {
        Id: bareGuid(info.id),
        names: (info.localizedNames || []).map((name: any) => name.name),
        terms
      }
    }
  } catch (error) {
    return {
      sets: {},
      unavailable: `the term store could not be read (${messageOf(error)}${
        error && error.status ? `, ${error.status}` : ''
      })`
    }
  }
  return state
}

/**
 * Reads what a query names from a web. Nothing is written.
 *
 * @param conn - The web
 * @param query - What to read
 */
export async function readSite(
  conn: ISiteConnection,
  query: ISiteQuery
): Promise<ISiteState> {
  const state: ISiteState = { unreadable: [] }
  if (query.siteFields) state.siteFields = await readSiteColumns(conn)
  if (query.contentTypes) {
    state.contentTypes = {}
    for (const id of query.contentTypes)
      state.contentTypes[id.toUpperCase()] = await readContentType(conn, id)
  }
  if (query.features)
    state.features = (
      await conn.web.features.select('DefinitionId')<
        Array<{ DefinitionId: string }>
      >()
    ).map((feature) => bareGuid(feature.DefinitionId))
  if (query.security) {
    const groups: Record<string, IGroupState | null> = {}
    for (const title of query.security.groups)
      groups[title] = await readGroup(conn, title)
    state.security = {
      groups,
      web: await readSecurable(() => conn.web),
      associated: await readAssociatedGroups(conn)
    }
  }
  if (query.associatedGroups && !query.security)
    state.associatedGroups = await readAssociatedGroups(conn)
  if (query.lists) {
    state.lists = {}
    for (const list of query.lists)
      state.lists[list.title] = await readListState(
        conn,
        list.title,
        list,
        state.unreadable
      )
  }
  if (query.customActions) {
    const actions = await conn.web.userCustomActions.select(
      ...CUSTOM_ACTION_MEMBERS
    )<Array<Record<string, any>>>()
    const text = (value: any): string =>
      value === null || value === undefined ? '' : String(value)
    state.customActions = actions.map((action) => ({
      Id: text(action.Id).toLowerCase(),
      Title: text(action.Title),
      Name: text(action.Name),
      Description: text(action.Description),
      Location: text(action.Location),
      Sequence: Number(action.Sequence),
      ClientSideComponentId: text(action.ClientSideComponentId).toLowerCase(),
      ClientSideComponentProperties: text(action.ClientSideComponentProperties),
      Url: text(action.Url),
      RegistrationId: text(action.RegistrationId),
      RegistrationType: Number(action.RegistrationType),
      ScriptBlock: text(action.ScriptBlock),
      ScriptSrc: text(action.ScriptSrc)
    }))
  }
  if (query.composedLook) {
    const look: ISiteState['composedLook'] = {
      theme: await readComposedLook(conn)
    }
    try {
      if (query.composedLook.palette)
        look.palette = await readPalette(conn, query.composedLook.palette)
      if (query.composedLook.fontScheme)
        look.fonts = await readFontScheme(conn, query.composedLook.fontScheme)
    } catch (error) {
      look.unreadable = messageOf(error)
    }
    state.composedLook = look
  }
  if (query.pages) {
    const library = `${conn.webInfo.ServerRelativeUrl.replace(
      /\/$/,
      ''
    )}/SitePages`
    state.pages = {}
    for (const name of query.pages)
      state.pages[name] = await readClientSidePage(conn, `${library}/${name}`)
  }
  if (query.navigation) {
    const tree = await readNavigationTree(conn)
    state.navigation = {
      base: {
        serverRelativeUrl: conn.webInfo.ServerRelativeUrl,
        origin: new URL(conn.webInfo.Url).origin
      },
      QuickLaunch: tree.quickLaunch,
      TopNavigationBar: tree.topNavigationBar
    }
  }
  if (query.webSettings) {
    const keys = query.webSettings.filter((key) => key !== 'WelcomePage')
    const settings: Record<string, any> = {}
    if (keys.length > 0) {
      let web: Record<string, any>
      try {
        web = await conn.web.select(...keys)()
      } catch {
        // An unknown key fails the whole `$select`: read the default JSON and say what is missing.
        web = await conn.web()
      }
      for (const key of keys)
        if (web[key] === undefined)
          state.unreadable.push({
            handler: 'WebSettings',
            path: key,
            reason: 'not readable through REST'
          })
        else settings[key] = web[key]
    }
    if (query.webSettings.includes('WelcomePage'))
      settings.WelcomePage =
        (
          await conn.web.rootFolder.select('WelcomePage')<{
            WelcomePage?: string
          }>()
        ).WelcomePage ?? ''
    state.webSettings = settings
  }
  if (query.propertyBag) {
    const raw = await conn.web.allProperties<Record<string, any>>()
    const bag: Record<string, any> = {}
    for (const key of Object.keys(raw))
      if (!/^odata\./.test(key)) bag[decodePropertyKey(key)] = raw[key]
    state.propertyBag = bag
  }
  if (query.files) {
    state.files = {}
    for (const url of query.files)
      state.files[url] = await conn.web
        .getFileByServerRelativePath(url)
        .select('Exists')<{ Exists: boolean }>()
        .then((file) => file.Exists !== false)
        .catch((error) => {
          if (notFound(error)) return false
          throw error
        })
  }
  if (query.taxonomy) state.taxonomy = await readTaxonomy(conn, query.taxonomy)
  return state
}
