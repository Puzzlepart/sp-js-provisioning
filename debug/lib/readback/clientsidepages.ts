import { ISiteConnection } from '../connect'

/**
 * A control on a modern page as SharePoint stores it, without instance ids or positions (the
 * position is the control's place in {@link IColumnState.controls}).
 */
export interface IControlState {
  /** `text` (control type 4) or `webpart` (control type 3); the number for anything else. */
  type: 'text' | 'webpart' | number
  /** A text control's HTML. */
  text?: string
  /** A web part's component id (`webPartId`), lower case. */
  webPartId?: string
  /** A web part's definition id (`webPartData.id`), lower case. */
  definitionId?: string
  /** A web part's title (`webPartData.title`, from the component's preconfigured entry). */
  title?: string
  /** A web part's properties (`webPartData.properties`). */
  properties?: Record<string, any>
  /** A web part's server processed content (`webPartData.serverProcessedContent`). */
  serverProcessedContent?: Record<string, any>
}

/**
 * A column of a section: its width (factor of 12) and its controls in order.
 */
export interface IColumnState {
  factor: number
  controls: IControlState[]
}

/**
 * A section of a page: its columns in order.
 */
export interface ISectionState {
  columns: IColumnState[]
}

/**
 * A modern page as `web.loadClientsidePage()` reads it, plus its file's identity.
 */
export interface IPageState {
  title: string
  layout: string
  /** From the file's list item (see {@link readClientSidePage}). */
  commentsDisabled: boolean
  /** The file's `Level` is 1 (published, not a draft). */
  published: boolean
  /** The file is checked out (`CheckOutType` is not 2, None). */
  checkedOut: boolean
  /** The sections (layout index 1), in order. */
  sections: ISectionState[]
  /** The vertical section's controls (layout index 2), or `null` when the page has none. */
  verticalSection: IControlState[] | null
  /** The file's `UniqueId`: a page that was deleted and created again has a new one. */
  uniqueId: string
  /** The file's `TimeLastModified`. */
  modified: string
}

/**
 * What the ClientSidePages handler provisions on a web: the named pages (`null` when a page does
 * not exist), and the names of the files in `SitePages` that start with the prefix (a page created
 * twice shows up as a second file, e.g. `Name(1).aspx`).
 */
export interface IClientSidePagesState {
  pages: Record<string, IPageState | null>
  files: string[]
}

/** The members of a page that change when the page is deleted and created again. */
export const VOLATILE_PAGE_MEMBERS: Array<keyof IPageState> = [
  'uniqueId',
  'modified'
]

/**
 * A page without its volatile members ({@link VOLATILE_PAGE_MEMBERS}), for comparing a page that
 * `Overwrite: true` recreates.
 *
 * @param page - The page
 */
export function withoutVolatile(
  page: IPageState | null
): Partial<IPageState> | null {
  if (!page) return page
  const copy: Partial<IPageState> = { ...page }
  for (const member of VOLATILE_PAGE_MEMBERS) delete copy[member]
  return copy
}

/**
 * A control as plain data.
 *
 * @param data - The control's JSON (`CanvasContent1` entry)
 */
function controlState(data: any): IControlState {
  if (data.controlType === 4) return { type: 'text', text: data.innerHTML }
  if (data.controlType !== 3) return { type: data.controlType }
  const part = data.webPartData || {}
  return {
    type: 'webpart',
    webPartId: data.webPartId && String(data.webPartId).toLowerCase(),
    definitionId: part.id && String(part.id).toLowerCase(),
    title: part.title,
    properties: part.properties || {},
    serverProcessedContent: part.serverProcessedContent || {}
  }
}

/**
 * The columns of a section as plain data, in order.
 *
 * @param section - A PnPjs `CanvasSection`
 */
function columnStates(section: any): IColumnState[] {
  return [...section.columns]
    .sort((a: any, b: any) => a.order - b.order)
    .map((column: any) => ({
      factor: column.factor,
      controls: [...column.controls]
        .sort((a: any, b: any) => a.order - b.order)
        .map((control: any) => controlState(control.data))
    }))
}

/**
 * What `_api/web/getFileByServerRelativePath(...)` says about a page's file.
 */
interface IFileInfo {
  UniqueId: string
  TimeLastModified: string
  /** 1 published, 2 draft, 255 checked out. */
  Level: number
  /** 0 online, 1 offline, 2 none. */
  CheckOutType: number
}

/**
 * Reads back one page, or `null` when there is no such file.
 *
 * @param conn - The web
 * @param serverRelativeUrl - The page's file
 */
export async function readClientSidePage(
  conn: ISiteConnection,
  serverRelativeUrl: string
): Promise<IPageState | null> {
  let file: IFileInfo
  try {
    file = await conn.web
      .getFileByServerRelativePath(serverRelativeUrl)
      .select(
        'UniqueId',
        'TimeLastModified',
        'Level',
        'CheckOutType'
      )<IFileInfo>()
  } catch (error) {
    if (error && error.status === 404) return null
    throw error
  }
  const page = await conn.web.loadClientsidePage(serverRelativeUrl)
  // Not `page.commentsDisabled`: PnPjs reads it from `lists(id)/items(id)?$select=CommentsDisabled`,
  // which answers `false` (scope 0) for a page whose comments are disabled, while the file's item
  // answers `true` (scope 1) and SharePoint refuses a comment with `CommentsDisabledOnItem`
  // (observed 2026-10-21).
  const { CommentsDisabled } = await conn.web
    .getFileByServerRelativePath(serverRelativeUrl)
    .listItemAllFields.select('CommentsDisabled')<{
    CommentsDisabled: boolean
  }>()
  const sections: any[] = page.sections
  const vertical = sections.find((section) => section.layoutIndex === 2)
  return {
    title: page.title,
    layout: page.pageLayout,
    commentsDisabled: CommentsDisabled,
    published: file.Level === 1,
    checkedOut: file.CheckOutType !== 2,
    sections: sections
      .filter((section) => section.layoutIndex !== 2)
      .map((section) => ({ columns: columnStates(section) })),
    verticalSection: vertical
      ? columnStates(vertical).reduce(
          (controls, column) => controls.concat(column.controls),
          [] as IControlState[]
        )
      : null,
    uniqueId: file.UniqueId.toLowerCase(),
    modified: file.TimeLastModified
  }
}

/**
 * Reads back the named pages in the web's `SitePages` library, and the files there that start with
 * the prefix.
 *
 * @param conn - The web
 * @param names - The pages' file names (`Name` in the template, e.g. `Home.aspx`)
 * @param prefix - The file name prefix the template's pages share
 */
export async function readClientSidePages(
  conn: ISiteConnection,
  names: string[],
  prefix: string
): Promise<IClientSidePagesState> {
  const library = `${conn.webInfo.ServerRelativeUrl.replace(
    /\/$/,
    ''
  )}/SitePages`
  const pages: Record<string, IPageState | null> = {}
  for (const name of names)
    pages[name] = await readClientSidePage(conn, `${library}/${name}`)
  const files = await conn.web
    .getFolderByServerRelativePath(library)
    .files.select('Name')<Array<{ Name: string }>>()
  return {
    pages,
    files: files
      .map((file) => file.Name)
      .filter((name) => name.toLowerCase().startsWith(prefix.toLowerCase()))
      .sort()
  }
}
