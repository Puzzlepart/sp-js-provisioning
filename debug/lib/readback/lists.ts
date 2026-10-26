import { ISiteConnection } from '../connect'

/**
 * A list field as SharePoint stores it. The lookup members are there for lookups only.
 */
export interface IListFieldState {
  Id: string
  InternalName: string
  Title: string
  TypeAsString: string
  Required: boolean
  Hidden: boolean
  /** The source list of a lookup, a bare lower-case GUID. */
  LookupList?: string
  /** The web of a lookup's source list, a bare lower-case GUID. */
  LookupWebId?: string
  LookupField?: string
  AllowMultipleValues?: boolean
}

/**
 * A list view with its fields, without its id.
 */
export interface IViewState {
  Title: string
  DefaultView: boolean
  Hidden: boolean
  PersonalView: boolean
  RowLimit: number
  Paged: boolean
  Scope: number
  ViewQuery: string
  /** The view fields, in order. */
  fields: string[]
}

/**
 * What the Lists handler provisions on a list (structure, not items), without volatile members.
 */
export interface IListState {
  Title: string
  Description: string
  BaseTemplate: number
  ContentTypesEnabled: boolean
  EnableVersioning: boolean
  EnableMinorVersions: boolean
  Hidden: boolean
  HasUniqueRoleAssignments: boolean
  /** The list content types, in the list's order. `Id` is the list content type's id. */
  contentTypes: Array<{ Id: string; Name: string }>
  /** The fields whose internal name starts with the prefix, by internal name (sorted). */
  fields: Record<string, IListFieldState>
  /** Every view, by title. */
  views: IViewState[]
  /** Every folder item, as a path relative to the list's root folder, sorted. */
  folders: string[]
  /** Role assignments: member title to role names, sorted. */
  roleAssignments: Record<string, string[]>
}

/**
 * Lists by title.
 */
export type ListsState = Record<string, IListState>

/** A GUID without braces, lower case. */
const bare = (guid?: string) =>
  guid === undefined || guid === null
    ? guid
    : String(guid).replace(/[{}]/g, '').toLowerCase()

/**
 * The role assignments of a securable object: member title to role names, sorted.
 *
 * @param assignments - `roleAssignments.expand('Member','RoleDefinitionBindings')()`
 */
export function roleAssignmentsOf(
  assignments: Array<{
    Member: { Title: string }
    RoleDefinitionBindings: Array<{ Name: string }>
  }>
): Record<string, string[]> {
  const result: Record<string, string[]> = {}
  for (const assignment of assignments)
    result[assignment.Member.Title] = assignment.RoleDefinitionBindings.map(
      (binding) => binding.Name
    ).sort()
  return result
}

/**
 * Reads back one list: its properties, content types, prefixed fields, views with their fields,
 * folders and role assignments.
 *
 * @param conn - The web
 * @param title - The list
 * @param fieldPrefix - Only fields whose internal name starts with this are read
 */
export async function readList(
  conn: ISiteConnection,
  title: string,
  fieldPrefix = 'SPJSItest'
): Promise<IListState> {
  const list = () => conn.web.lists.getByTitle(title)
  const info: any = await list().select(
    'Title',
    'Description',
    'BaseTemplate',
    'ContentTypesEnabled',
    'EnableVersioning',
    'EnableMinorVersions',
    'Hidden',
    'HasUniqueRoleAssignments'
  )()
  const contentTypes: any[] = await list().contentTypes.select('Id', 'Name')()
  const fields: any[] = await list().fields.select(
    'Id',
    'InternalName',
    'Title',
    'TypeAsString',
    'Required',
    'Hidden',
    'LookupList',
    'LookupWebId',
    'LookupField',
    'AllowMultipleValues'
  )()
  const views: any[] = await list().views.select(
    'Id',
    'Title',
    'DefaultView',
    'Hidden',
    'PersonalView',
    'RowLimit',
    'Paged',
    'Scope',
    'ViewQuery'
  )()
  const root: { ServerRelativeUrl: string } = await list().rootFolder.select(
    'ServerRelativeUrl'
  )()
  const folderItems: Array<{ FileRef: string }> = await list()
    .items.filter('FSObjType eq 1')
    .select('FileRef')
    .top(5000)()
  const assignments: any[] = await list().roleAssignments.expand(
    'Member',
    'RoleDefinitionBindings'
  )()

  // By internal name, sorted: a recreated field moves to the end of the collection.
  const fieldStates: Record<string, IListFieldState> = {}
  for (const field of fields
    .filter((f) => String(f.InternalName).startsWith(fieldPrefix))
    .sort((a, b) => String(a.InternalName).localeCompare(b.InternalName))) {
    const isLookup = String(field.TypeAsString).startsWith('Lookup')
    fieldStates[field.InternalName] = {
      Id: bare(field.Id),
      InternalName: field.InternalName,
      Title: field.Title,
      TypeAsString: field.TypeAsString,
      Required: field.Required,
      Hidden: field.Hidden,
      ...(isLookup
        ? {
            LookupList: bare(field.LookupList),
            LookupWebId: bare(field.LookupWebId),
            LookupField: field.LookupField,
            AllowMultipleValues: field.AllowMultipleValues
          }
        : {})
    }
  }

  const viewStates: IViewState[] = []
  for (const view of views) {
    const viewFields: { Items: string[] } = await list()
      .views.getById(view.Id)
      .fields()
    viewStates.push({
      Title: view.Title,
      DefaultView: view.DefaultView,
      Hidden: view.Hidden,
      PersonalView: view.PersonalView,
      RowLimit: view.RowLimit,
      Paged: view.Paged,
      Scope: view.Scope,
      ViewQuery: view.ViewQuery,
      fields: viewFields.Items
    })
  }
  viewStates.sort((a, b) => a.Title.localeCompare(b.Title))

  const rootUrl = root.ServerRelativeUrl.replace(/\/+$/, '')
  return {
    Title: info.Title,
    Description: info.Description,
    BaseTemplate: info.BaseTemplate,
    ContentTypesEnabled: info.ContentTypesEnabled,
    EnableVersioning: info.EnableVersioning,
    EnableMinorVersions: info.EnableMinorVersions,
    Hidden: info.Hidden,
    HasUniqueRoleAssignments: info.HasUniqueRoleAssignments,
    contentTypes: contentTypes.map((ct) => ({
      Id: ct.Id.StringValue,
      Name: ct.Name
    })),
    fields: fieldStates,
    views: viewStates,
    folders: folderItems
      .map((item) =>
        item.FileRef.startsWith(`${rootUrl}/`)
          ? item.FileRef.slice(rootUrl.length + 1)
          : item.FileRef
      )
      .sort(),
    roleAssignments: roleAssignmentsOf(assignments)
  }
}

/**
 * Reads back the lists with the given titles.
 *
 * @param conn - The web
 * @param titles - The lists
 * @param fieldPrefix - Only fields whose internal name starts with this are read
 */
export async function readLists(
  conn: ISiteConnection,
  titles: string[],
  fieldPrefix = 'SPJSItest'
): Promise<ListsState> {
  const state: ListsState = {}
  for (const title of titles)
    state[title] = await readList(conn, title, fieldPrefix)
  return state
}
