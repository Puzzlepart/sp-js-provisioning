import { ISiteConnection } from './connect'
import {
  FIELD_MEMBERS,
  fieldStateOf,
  IContentTypeState,
  IFieldState,
  ISiteState,
  readContentType,
  readSite
} from './readback/site'

/**
 * `inspect`: what a web holds of what the handlers write, as JSON (docs/specs/integration-testing.md,
 * Commands). Read-only. It reads with the same readers as `verify` (`readback/site.ts`), but names
 * what to read from the web instead of from a template.
 */

/**
 * What to inspect.
 */
export interface IInspectOptions {
  /**
   * Only what is named with this prefix, ignoring case: site columns (internal name, title or
   * group), content types (name or group), site groups, lists, custom actions and pages. The web's
   * permissions, features, navigation, settings and property bag are read whole.
   */
  prefix?: string
}

/** The web settings `inspect` reads (the WebSettings handler sets any key it is given). */
export const INSPECTED_WEB_SETTINGS = [
  'Title',
  'Description',
  'SiteLogoUrl',
  'AlternateCssUrl',
  'MasterUrl',
  'CustomMasterUrl',
  'QuickLaunchEnabled',
  'TreeViewEnabled',
  'WelcomePage'
]

/**
 * What a web holds, as `inspect` prints it.
 */
export interface IInspection {
  web: {
    Title: string
    Url: string
    ServerRelativeUrl: string
    Language?: number
  }
  prefix?: string
  /** The web's own site columns (not those of a parent web), by internal name. */
  siteFields: Record<string, IFieldState & { Group: string }>
  /** The web's own content types, by id. */
  contentTypes: Record<string, IContentTypeState>
  features: ISiteState['features']
  security: ISiteState['security']
  lists: ISiteState['lists']
  customActions: ISiteState['customActions']
  composedLook: ISiteState['composedLook']
  pages: ISiteState['pages']
  navigation: ISiteState['navigation']
  webSettings: ISiteState['webSettings']
  propertyBag: ISiteState['propertyBag']
  unreadable: ISiteState['unreadable']
}

/**
 * Whether a name starts with the prefix, ignoring case (always, without a prefix).
 *
 * @param prefix - The prefix
 * @param names - The names to try
 */
export function matchesPrefix(
  prefix: string | undefined,
  ...names: any[]
): boolean {
  if (!prefix) return true
  const wanted = prefix.toLowerCase()
  return names.some((name) =>
    String(name ?? '')
      .toLowerCase()
      .startsWith(wanted)
  )
}

/**
 * Reads what the handlers write to a web. Nothing is written.
 *
 * @param conn - The web
 * @param options - The prefix
 */
export async function inspectWeb(
  conn: ISiteConnection,
  options: IInspectOptions = {}
): Promise<IInspection> {
  const { prefix } = options
  // `conn.web` is a fresh web on every access, so each read gets its own query.
  const web = () => conn.web
  const [fields, contentTypes, groups, lists, files] = await Promise.all([
    web().fields.select(...FIELD_MEMBERS, 'Group')<any[]>(),
    web().contentTypes.select('Id', 'Name', 'Group')<any[]>(),
    web().siteGroups.select('Title')<Array<{ Title: string }>>(),
    web().lists.filter('Hidden eq false').select('Title')<
      Array<{ Title: string }>
    >(),
    web()
      .getFolderByServerRelativePath(
        `${conn.webInfo.ServerRelativeUrl.replace(/\/$/, '')}/SitePages`
      )
      .files.select('Name')<Array<{ Name: string }>>()
      .catch((): Array<{ Name: string }> => [])
  ])
  const siteFields: IInspection['siteFields'] = {}
  for (const field of fields)
    if (
      (prefix || !field.Hidden) &&
      matchesPrefix(prefix, field.InternalName, field.Title, field.Group)
    )
      siteFields[field.InternalName] = {
        ...fieldStateOf(field),
        Group: field.Group ?? ''
      }
  const ownContentTypes: IInspection['contentTypes'] = {}
  for (const ct of contentTypes.filter((candidate) =>
    matchesPrefix(prefix, candidate.Name, candidate.Group)
  )) {
    const id = String(ct.Id.StringValue).toUpperCase()
    const state = await readContentType(conn, id)
    if (state) ownContentTypes[id] = state
  }
  const state = await readSite(conn, {
    features: true,
    security: {
      groups: groups
        .map((group) => group.Title)
        .filter((title) => matchesPrefix(prefix, title))
    },
    lists: lists
      .map((list) => list.Title)
      .filter((title) => matchesPrefix(prefix, title))
      .map((title) => ({ title })),
    customActions: true,
    composedLook: {},
    pages: files
      .map((file) => file.Name)
      .filter((name) => /\.aspx$/i.test(name) && matchesPrefix(prefix, name)),
    navigation: true,
    webSettings: INSPECTED_WEB_SETTINGS,
    propertyBag: true
  })
  return {
    web: {
      Title: conn.webInfo.Title,
      Url: conn.webInfo.Url,
      ServerRelativeUrl: conn.webInfo.ServerRelativeUrl,
      ...(conn.webInfo.Language !== undefined
        ? { Language: conn.webInfo.Language }
        : {})
    },
    ...(prefix ? { prefix } : {}),
    siteFields,
    contentTypes: ownContentTypes,
    features: state.features,
    security: state.security,
    lists: state.lists,
    customActions: (state.customActions || []).filter((action) =>
      matchesPrefix(prefix, action.Title)
    ),
    composedLook: state.composedLook,
    pages: state.pages,
    navigation: state.navigation,
    webSettings: state.webSettings,
    propertyBag: state.propertyBag,
    unreadable: state.unreadable
  }
}
