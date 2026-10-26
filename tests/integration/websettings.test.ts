import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { ISiteConnection } from '../../debug/lib/connect'
import {
  IWebSettingsState,
  readWebSettings
} from '../../debug/lib/readback/websettings'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, skipReason, withScratchWeb } from './helpers'

/**
 * The fixture's settings as the engine applies them: URL tokens replaced in string values, with
 * `{site}` the site collection's server-relative URL (the SPFx context's `pageContext.site`).
 *
 * @param conn - The scratch web
 * @param name - The fixture
 */
function expected(
  conn: ISiteConnection,
  name: string
): Partial<IWebSettingsState> {
  const settings: Record<string, any> = loadFixture(name).WebSettings
  const out: Record<string, any> = {}
  for (const key of Object.keys(settings))
    out[key] =
      typeof settings[key] === 'string'
        ? settings[key].replace(/{site}/g, conn.siteInfo.ServerRelativeUrl)
        : settings[key]
  return out
}

/**
 * Puts the web into a known state through REST, so a read-back can tell the handler's values from
 * the web's defaults.
 *
 * @param conn - The scratch web
 * @param state - The state
 */
async function arrange(
  conn: ISiteConnection,
  state: IWebSettingsState
): Promise<void> {
  const { WelcomePage, ...properties } = state
  await conn.web.rootFolder.update({ WelcomePage })
  await conn.web.update(properties)
  assert.deepEqual(await readWebSettings(conn), state, 'the arranged state')
}

/**
 * Asserts that every value the template sets differs from the arranged one, so the read-back
 * proves the apply rather than a default.
 *
 * @param before - The arranged state
 * @param settings - What the template sets
 */
function assertChanges(
  before: IWebSettingsState,
  settings: Partial<IWebSettingsState>
): void {
  for (const key of Object.keys(settings))
    assert.notEqual(
      before[key as keyof IWebSettingsState],
      settings[key as keyof IWebSettingsState],
      `the arranged ${key} already has the template's value`
    )
}

/**
 * The WebSettings handler on a scratch web: the welcome page on the root folder, the other
 * settings on the web. Master page and CSS URLs are left out: custom script is blocked on the site
 * (NoScript) and those properties need it.
 */
describe('WebSettings', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('websettings')
  let otherPage: string

  before(async () => {
    const { conn } = scratch()
    // A second page that exists on every web, whatever its language: the default view of the
    // default document library (`Delte dokumenter/Forms/AllItems.aspx` on a Norwegian web),
    // relative to the web as `WelcomePage` is.
    const view = await conn.web.defaultDocumentLibrary.defaultView.select(
      'ServerRelativeUrl'
    )<{ ServerRelativeUrl: string }>()
    otherPage = view.ServerRelativeUrl.slice(
      conn.webInfo.ServerRelativeUrl.length + 1
    )
  })

  it('sets the welcome page, description, tree view and quick launch, and changes nothing on a second apply', async () => {
    const { conn } = scratch()
    const settings = expected(conn, 'websettings')
    // The scratch web's welcome page is `SitePages/Home.aspx` (also on a Norwegian web): move it
    // away first, so the fixture's value is visible in the read-back.
    const arranged: IWebSettingsState = {
      Description: 'SPJS Itest before',
      TreeViewEnabled: false,
      QuickLaunchEnabled: true,
      WelcomePage: otherPage
    }
    await arrange(conn, arranged)
    assertChanges(arranged, settings)

    const { first, afterFirst } = await applyTwice(
      conn,
      () => loadFixture('websettings'),
      readWebSettings
    )

    assert.deepEqual(first.order, ['WebSettings'])
    assert.deepEqual(afterFirst, settings)
    assert.equal(
      afterFirst.Description,
      `SPJS Itest web settings in ${conn.siteInfo.ServerRelativeUrl}`
    )
  })

  it('applies the settings without a WelcomePage, leaving the welcome page and quick launch as they are (D4)', async () => {
    const { conn } = scratch()
    const settings = expected(conn, 'websettings.b')
    assert.equal(settings.WelcomePage, undefined)
    const arranged: IWebSettingsState = {
      Description: 'SPJS Itest before',
      TreeViewEnabled: true,
      QuickLaunchEnabled: false,
      WelcomePage: otherPage
    }
    await arrange(conn, arranged)
    assertChanges(arranged, settings)

    const { first, afterFirst } = await applyTwice(
      conn,
      () => loadFixture('websettings.b'),
      readWebSettings
    )

    assert.deepEqual(first.order, ['WebSettings'])
    // Until 1.5.0 the handler updated the web only inside `if (WelcomePage)`: this applied nothing.
    assert.deepEqual(afterFirst, { ...arranged, ...settings })
  })
})
