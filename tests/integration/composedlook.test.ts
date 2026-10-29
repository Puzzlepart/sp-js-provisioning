import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import {
  readComposedLook,
  readFontScheme,
  readPalette
} from '../../debug/lib/readback/composedlook'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

/**
 * The last segment of a URL or path.
 *
 * @param url - The URL
 */
const lastSegment = (url: string) => url.split('/').pop()

/**
 * What the fixture's look should leave on the web: the slots of the gallery files it names (read
 * from the site collection's theme gallery, where the tokens point), and the themed copy of its
 * background image SharePoint writes into the web (`.jpg` becomes `.themedjpg`).
 *
 * @param conn - The scratch web
 */
async function expected(conn: ISiteConnection) {
  const { ComposedLook } = loadFixture('composedlook')
  const gallery = `${conn.siteInfo.ServerRelativeUrl.replace(
    /\/$/,
    ''
  )}/_catalogs/theme/15`
  const extension = lastSegment(ComposedLook.BackgroundImageUrl)
    .split('.')
    .pop()
  return {
    palette: await readPalette(
      conn,
      `${gallery}/${lastSegment(ComposedLook.ColorPaletteUrl)}`
    ),
    fonts: await readFontScheme(
      conn,
      `${gallery}/${lastSegment(ComposedLook.FontSchemeUrl)}`
    ),
    backgroundImage: `${conn.webInfo.ServerRelativeUrl}/_themes/<n>/<id>bkimage-<hash>.themed${extension}`
  }
}

/**
 * The fixture with members of its look removed: `''` (what a template states as "none") or left
 * out. Fresh on every call, like `loadFixture`.
 *
 * @param empty - Members set to `''`
 * @param omitted - Members deleted
 */
const variant = (empty: string[], omitted: string[] = []) => () => {
  const fixture = loadFixture('composedlook')
  for (const key of empty) fixture.ComposedLook[key] = ''
  for (const key of omitted) delete fixture.ComposedLook[key]
  return fixture
}

/**
 * The ComposedLook handler on a scratch web: `web.applyTheme` with an OOTB palette, font scheme and
 * background image of the site collection (the "Sea Monster" look), reached through the three URL
 * tokens `{sitecollection}`, `{themegallery}` and `{site}`. The proof is `_api/web/ThemeInfo`: its
 * shades and fonts equal what the gallery's palette and font scheme files hold (read in the test),
 * and differ from the web's before. Custom script being blocked (NoScript) does not stop
 * `applyTheme` on a subweb (observed 2026-10-21), so nothing here skips for it.
 */
describe('ComposedLook', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('composedlook')

  it('applies the palette, font scheme and background image, and changes nothing on a second apply', async () => {
    const { conn } = scratch()
    const look = await expected(conn)
    // The gallery files were read: every slot has a value.
    for (const value of [
      ...Object.keys(look.palette).map((name) => look.palette[name]),
      ...Object.keys(look.fonts).map((slot) => look.fonts[slot])
    ])
      assert.ok(value, 'a slot of the gallery files has no value')
    const before = await readComposedLook(conn)
    // (`notDeepEqual` is not in the tests' `node:assert` typings.)
    assert.notEqual(
      JSON.stringify(before.shades),
      JSON.stringify(look.palette),
      'the scratch web already has the palette: the test proves nothing'
    )

    const { first, afterFirst } = await applyTwice(
      conn,
      () => loadFixture('composedlook'),
      readComposedLook
    )
    assert.deepEqual(first.order, ['ComposedLook'])
    assert.deepEqual(afterFirst.shades, look.palette)
    assert.deepEqual(afterFirst.fonts, look.fonts)
    assert.notEqual(
      afterFirst.accessibleDescription,
      before.accessibleDescription
    )
    // The theme is generated into this web (`shareGenerated: false`), not the root, with a
    // recoloured copy of the background image: REST shows that a copy is there, not its source.
    assert.equal(afterFirst.backgroundImageUri, look.backgroundImage)
    assert.equal(
      afterFirst.themedCssFolderUrl,
      `${conn.webInfo.ServerRelativeUrl}/_themes/<n>`
    )
  })

  it('applies a look without a background image, removing the one before', async () => {
    const { conn } = scratch()
    const look = await expected(conn)
    assertClean(await apply(conn, loadFixture('composedlook')), 'the full look')
    assert.equal(
      (await readComposedLook(conn)).backgroundImageUri,
      look.backgroundImage
    )
    const { afterFirst } = await applyTwice(
      conn,
      variant(['BackgroundImageUrl']),
      readComposedLook
    )
    assert.equal(afterFirst.backgroundImageUri, null)
    assert.deepEqual(afterFirst.shades, look.palette)
    assert.deepEqual(afterFirst.fonts, look.fonts)
  })

  it('applies a look with a palette only', async () => {
    const { conn } = scratch()
    const look = await expected(conn)
    // Another palette first, so the read-back below can only pass if the palette-only look is
    // applied (the test before left Palette005 on the web).
    const other = loadFixture('composedlook')
    other.ComposedLook.ColorPaletteUrl = other.ComposedLook.ColorPaletteUrl.replace(
      /Palette005\.spcolor$/,
      'Palette001.spcolor'
    )
    assertClean(await apply(conn, other), 'another palette first')
    assert.notDeepEqual(
      (await readComposedLook(conn)).shades,
      look.palette,
      'Palette001 and Palette005 have different shades'
    )
    const { afterFirst } = await applyTwice(
      conn,
      variant([], ['FontSchemeUrl', 'BackgroundImageUrl']),
      readComposedLook
    )
    assert.deepEqual(afterFirst.shades, look.palette)
    assert.equal(afterFirst.backgroundImageUri, null)
  })
})
