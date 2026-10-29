import { ISiteConnection } from '../connect'

/**
 * What the WebSettings handler provisions on a web, as far as the integration fixtures set it:
 * three web properties and the root folder's welcome page (relative to the web, as SharePoint
 * stores it).
 */
export interface IWebSettingsState {
  Description: string
  TreeViewEnabled: boolean
  QuickLaunchEnabled: boolean
  WelcomePage: string
}

/**
 * Reads back the web's settings and its root folder's welcome page.
 *
 * @param conn - The web
 */
export async function readWebSettings(
  conn: ISiteConnection
): Promise<IWebSettingsState> {
  const [web, rootFolder] = await Promise.all([
    conn.web.select('Description', 'TreeViewEnabled', 'QuickLaunchEnabled')<{
      Description?: string
      TreeViewEnabled: boolean
      QuickLaunchEnabled: boolean
    }>(),
    conn.web.rootFolder.select('WelcomePage')<{ WelcomePage?: string }>()
  ])
  return {
    Description: web.Description ?? '',
    TreeViewEnabled: web.TreeViewEnabled,
    QuickLaunchEnabled: web.QuickLaunchEnabled,
    WelcomePage: rootFolder.WelcomePage ?? ''
  }
}
