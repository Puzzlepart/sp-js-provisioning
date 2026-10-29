import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { ComposedLook } from '../../src/handlers/composedlook'

const spfxContext = {
  pageContext: {
    site: {
      absoluteUrl: 'https://contoso.sharepoint.com/sites/hub',
      serverRelativeUrl: '/sites/hub'
    },
    web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/hub' }
  }
}

describe('ComposedLook', () => {
  it('applies a palette without a font scheme or a background image', async () => {
    // A look without FontSchemeUrl used to throw a TypeError in replaceUrlTokens.
    const calls: any[] = []
    const web: any = {
      applyTheme: async (...args: any[]) => {
        calls.push(args)
      }
    }

    await new ComposedLook({ spfxContext }).ProvisionObjects(web, {
      ColorPaletteUrl: '{sitecollection}/_catalogs/theme/15/palette011.spcolor'
    })

    assert.deepEqual(calls, [
      ['/sites/hub/_catalogs/theme/15/palette011.spcolor', null, null, false]
    ])
  })
})
