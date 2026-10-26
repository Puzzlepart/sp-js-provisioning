import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { WebSettings } from '../../src/handlers/websettings'

const spfxContext = {
  pageContext: {
    site: {
      absoluteUrl: 'https://contoso.sharepoint.com/sites/hub',
      serverRelativeUrl: '/sites/hub'
    },
    web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/hub/sub' }
  }
}

/**
 * A fake web that records the bodies of `web.update` and `web.rootFolder.update`.
 *
 * @param fail - The error `web.update` rejects with, if any
 */
function fakeWeb(fail?: Error) {
  const calls = { web: [] as any[], rootFolder: [] as any[] }
  const web: any = {
    update: (properties: any) => {
      calls.web.push({ ...properties })
      return fail ? Promise.reject(fail) : Promise.resolve()
    },
    rootFolder: {
      update: (properties: any) => {
        calls.rootFolder.push({ ...properties })
        return Promise.resolve()
      }
    }
  }
  return { web, calls }
}

describe('WebSettings', () => {
  it('updates the web without a WelcomePage', async () => {
    const { web, calls } = fakeWeb()
    await new WebSettings({}).ProvisionObjects(web, {
      Description: 'SPJS Itest',
      TreeViewEnabled: true
    })
    assert.deepEqual(calls.web, [
      { Description: 'SPJS Itest', TreeViewEnabled: true }
    ])
    assert.deepEqual(calls.rootFolder, [])
  })

  it('sets the WelcomePage on the root folder and the other settings on the web', async () => {
    const { web, calls } = fakeWeb()
    await new WebSettings({}).ProvisionObjects(web, {
      WelcomePage: 'SitePages/Home.aspx',
      QuickLaunchEnabled: false
    })
    assert.deepEqual(calls.rootFolder, [{ WelcomePage: 'SitePages/Home.aspx' }])
    assert.deepEqual(calls.web, [{ QuickLaunchEnabled: false }])
  })

  it('sends no empty update', async () => {
    const empty = fakeWeb()
    await new WebSettings({}).ProvisionObjects(empty.web, {})
    assert.deepEqual(empty.calls, { web: [], rootFolder: [] })

    const welcomeOnly = fakeWeb()
    await new WebSettings({}).ProvisionObjects(welcomeOnly.web, {
      WelcomePage: 'SitePages/Home.aspx'
    })
    assert.deepEqual(welcomeOnly.calls, {
      web: [],
      rootFolder: [{ WelcomePage: 'SitePages/Home.aspx' }]
    })
  })

  it('never sends WelcomePage to the web, and leaves the root folder alone when it is empty', async () => {
    const { web, calls } = fakeWeb()
    await new WebSettings({}).ProvisionObjects(web, {
      WelcomePage: '',
      Description: 'SPJS Itest'
    })
    assert.deepEqual(calls.rootFolder, [])
    assert.deepEqual(calls.web, [{ Description: 'SPJS Itest' }])
  })

  it('replaces URL tokens in string values and writes them back into the settings', async () => {
    const { web, calls } = fakeWeb()
    const settings = {
      SiteLogoUrl: '{site}/SiteAssets/logo.png',
      RecycleBinEnabled: true
    }
    await new WebSettings({ spfxContext }).ProvisionObjects(web, settings)
    assert.deepEqual(calls.web, [
      { SiteLogoUrl: '/sites/hub/SiteAssets/logo.png', RecycleBinEnabled: true }
    ])
    assert.equal(settings.SiteLogoUrl, '/sites/hub/SiteAssets/logo.png')
  })

  it('rejects with the error of a failed update', async () => {
    const refused = new Error('Access denied')
    const { web } = fakeWeb(refused)
    await assert.rejects(
      new WebSettings({}).ProvisionObjects(web, { TreeViewEnabled: true }),
      (error) => error === refused
    )
  })
})
