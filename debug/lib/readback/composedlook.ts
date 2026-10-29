import { spPost } from '@pnp/sp'
import { Web } from '@pnp/sp/presets/all'
import { ISiteConnection } from '../connect'

/**
 * The palette slots read back: text, title, accent, link, a translucent and an opaque background,
 * navigation and a content accent. `ThemeInfo` answers one slot per request (a POST), so not all
 * 90-odd slots of a palette are read.
 */
export const SHADE_NAMES = [
  'BodyText',
  'SiteTitle',
  'AccentText',
  'Hyperlink',
  'BackgroundOverlay',
  'PageBackground',
  'Navigation',
  'ContentAccent1'
]

/** The font slots read back. */
export const FONT_SLOTS = ['title', 'navigation', 'heading', 'body']

/**
 * What a composed look leaves on a web, as `_api/web/ThemeInfo` and the web report it.
 */
export interface IComposedLookState {
  /** The palette in words, in the web's language (`ThemeInfo.AccessibleDescription`). */
  accessibleDescription: string
  /**
   * `ThemeInfo.ThemeBackgroundImageUri`: not the image the look names but a recoloured copy
   * SharePoint writes into the themed folder (`<web>/_themes/0/<guid>bkimage-<hash>.themedjpg`
   * for a `.jpg`), so generation, id and hash are replaced (`<n>`, `<id>`, `<hash>`); `null`
   * without a background image.
   */
  backgroundImageUri: string | null
  /**
   * `ThemedCssFolderUrl`, with the generation SharePoint counts up on every `applyTheme`
   * (`<web>/_themes/0`, then `/_themes/1`) replaced by `<n>`; `null` on a web never themed.
   */
  themedCssFolderUrl: string | null
  /** Palette slot to ARGB, upper case (`ThemeInfo.GetThemeShadeByName`). */
  shades: Record<string, string>
  /** Font slot to Latin typeface in the web's language (`ThemeInfo.GetThemeFontByName`). */
  fonts: Record<string, string>
}

/**
 * Reads back the theme of a web: what `ThemeInfo` says about its palette, font scheme and
 * background image, and where its themed CSS lives.
 *
 * @param conn - The web
 */
export async function readComposedLook(
  conn: ISiteConnection
): Promise<IComposedLookState> {
  const info = await Web(conn.web, 'ThemeInfo').select(
    'AccessibleDescription',
    'ThemeBackgroundImageUri'
  )<{ AccessibleDescription: string; ThemeBackgroundImageUri: string | null }>()
  const { ThemedCssFolderUrl } = await conn.web.select('ThemedCssFolderUrl')<{
    ThemedCssFolderUrl: string | null
  }>()
  const shades: Record<string, string> = {}
  for (const name of SHADE_NAMES)
    shades[name] = (
      await spPost<string>(
        Web(conn.web, `ThemeInfo/GetThemeShadeByName(name='${name}')`)
      )
    ).toUpperCase()
  const fonts: Record<string, string> = {}
  for (const slot of FONT_SLOTS)
    fonts[slot] = await spPost<string>(
      Web(
        conn.web,
        `ThemeInfo/GetThemeFontByName(name='${slot}',lcid=${conn.webInfo.Language})`
      )
    )
  return {
    accessibleDescription: info.AccessibleDescription ?? '',
    backgroundImageUri: info.ThemeBackgroundImageUri
      ? info.ThemeBackgroundImageUri.replace(
          /\/_themes\/\d+\//,
          '/_themes/<n>/'
        ).replace(/[\da-f-]{36}bkimage-[\da-f]+\./i, '<id>bkimage-<hash>.')
      : null,
    themedCssFolderUrl: ThemedCssFolderUrl
      ? ThemedCssFolderUrl.replace(/\/_themes\/\d+$/, '/_themes/<n>')
      : null,
    shades,
    fonts
  }
}

/**
 * The text of a file of the site collection's theme gallery. A subweb refuses a server-relative
 * URL outside itself ("must start with SPWeb.ServerRelativeUrl"), also through
 * `_api/site/rootweb`, so the file is read from the root web's own `_api`.
 *
 * @param conn - A web of the site collection
 * @param serverRelativeUrl - The file
 */
function galleryFile(
  conn: ISiteConnection,
  serverRelativeUrl: string
): Promise<string> {
  return Web([conn.web, conn.siteInfo.Url])
    .getFileByServerRelativePath(serverRelativeUrl)
    .getText()
}

/**
 * The slots of a color palette file (`.spcolor`) as `ThemeInfo` reports them once applied: ARGB,
 * upper case, opaque (`FF`) where the file gives only RGB.
 *
 * @param conn - A web of the site collection (the file is read through its root web)
 * @param serverRelativeUrl - The palette, e.g. `/sites/x/_catalogs/theme/15/Palette005.spcolor`
 * @param names - The slots to return
 */
export async function readPalette(
  conn: ISiteConnection,
  serverRelativeUrl: string,
  names: string[] = SHADE_NAMES
): Promise<Record<string, string>> {
  const xml = await galleryFile(conn, serverRelativeUrl)
  const all: Record<string, string> = {}
  const color = /<(?:\w+:)?color\s+name="([^"]+)"\s+value="([\da-f]{6}|[\da-f]{8})"/gi
  for (let match = color.exec(xml); match; match = color.exec(xml))
    all[match[1]] = (match[2].length === 6
      ? `FF${match[2]}`
      : match[2]
    ).toUpperCase()
  const picked: Record<string, string> = {}
  for (const name of names) picked[name] = all[name]
  return picked
}

/**
 * The Latin typeface of each slot of a font scheme file (`.spfont`).
 *
 * @param conn - A web of the site collection (the file is read through its root web)
 * @param serverRelativeUrl - The font scheme, e.g. `/sites/x/_catalogs/theme/15/fontscheme003.spfont`
 * @param slots - The slots to return
 */
export async function readFontScheme(
  conn: ISiteConnection,
  serverRelativeUrl: string,
  slots: string[] = FONT_SLOTS
): Promise<Record<string, string>> {
  const xml = await galleryFile(conn, serverRelativeUrl)
  const picked: Record<string, string> = {}
  for (const slot of slots) {
    const match = new RegExp(
      `<(?:\\w+:)?fontSlot\\s+name="${slot}"[^>]*>[\\s\\S]*?<(?:\\w+:)?latin\\s+typeface="([^"]*)"`
    ).exec(xml)
    picked[slot] = match ? match[1] : undefined
  }
  return picked
}
