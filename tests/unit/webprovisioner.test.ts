import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
// Importing the real `@pnp/sp` here is deliberate: the term store lookup is a
// hand-written url, and only the real queryable proves what url is sent.
import { Web } from '@pnp/sp/presets/all'
import { RejectOnError, ResolveOnData } from '@pnp/queryable'
import { WebProvisioner } from '../../src/webprovisioner'

/**
 * Records every request url the provisioner produces and short-circuits the
 * request by supplying an empty result in the `pre` moment, so no network call
 * is made. Observers are inherited by the child queryables the provisioner
 * builds from this web.
 *
 * @param urls - Array every requested url is pushed to
 */
function createRecordingWeb(urls: string[]): any {
  const recording = (instance: any) => {
    instance.on.pre(async (url: string, init: any) => {
      urls.push(url)
      return [url, init, {}]
    })
    return instance
  }
  // `ResolveOnData`/`RejectOnError` are what settle the request promise - they
  // normally come from `spfi`, which is not used here.
  return Web('https://contoso.sharepoint.com/sites/demo').using(
    ResolveOnData(),
    RejectOnError(),
    recording
  )
}

/**
 * Runs `loadContext` against a recording web and returns the urls it requested.
 */
async function loadContextUrls(): Promise<string[]> {
  const urls: string[] = []
  const provisioner: any = new WebProvisioner(createRecordingWeb(urls))
  provisioner.setup({})
  await provisioner.loadContext()
  return urls
}

describe('WebProvisioner: loadContext urls', () => {
  it('reads the default term store from the v2.1 termstore endpoint', async () => {
    // v4 dropped `@pnp/sp-taxonomy`, so the term store id is read through a
    // hand-written `SPInstance` url. The lookup fails soft, which means a wrong
    // url degrades silently to an unresolved {sitecollectiontermstoreid}.
    const urls = await loadContextUrls()
    const termStoreUrl = urls.find((url) => url.includes('termstore'))

    assert.ok(termStoreUrl, `no term store request in ${JSON.stringify(urls)}`)
    assert.equal(
      termStoreUrl.split('?')[0],
      'https://contoso.sharepoint.com/sites/demo/_api/v2.1/termstore'
    )
    assert.match(decodeURIComponent(termStoreUrl), /[?&]\$select=id(&|$)/)
  })

  it('asks for the site collection id', async () => {
    const urls = await loadContextUrls()
    const siteUrl = urls.find((url) => url.includes('/_api/site'))

    assert.ok(siteUrl, `no site request in ${JSON.stringify(urls)}`)
    assert.equal(
      siteUrl.split('?')[0],
      'https://contoso.sharepoint.com/sites/demo/_api/site'
    )
  })

  it('pages the list lookup so {listid:...} tokens cannot miss lists', async () => {
    // Without an explicit $top SharePoint returns only the first 100 lists.
    const urls = await loadContextUrls()
    const listsUrl = urls.find((url) => url.includes('/_api/web/lists'))

    assert.ok(listsUrl, `no lists request in ${JSON.stringify(urls)}`)
    assert.match(decodeURIComponent(listsUrl), /[?&]\$top=5000(&|$)/)
  })
})
