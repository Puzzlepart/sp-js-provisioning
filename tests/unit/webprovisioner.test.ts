import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
// Importing the real `@pnp/sp` here is deliberate: the term store lookup is a
// hand-written url, and only the real queryable proves what url is sent.
import { Web } from '@pnp/sp/presets/all'
import { RejectOnError, ResolveOnData } from '@pnp/queryable'
import { WebProvisioner } from '../../src/webprovisioner'
import { Lists } from '../../src/handlers/lists'
import { ProvisioningError } from '../../src/provisioningerror'
import { Logger, LogLevel } from '@pnp/logging'

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

describe('WebProvisioner: handler failures', () => {
  /**
   * Applies a template to a recording web and returns the error and the
   * handlers that ran.
   *
   * @param template - The template
   * @param config - The provisioning config
   */
  async function apply(
    template: any,
    config: any
  ): Promise<{ error: any; ran: string[] }> {
    const ran: string[] = []
    const provisioner = new WebProvisioner(createRecordingWeb([])).setup(config)
    try {
      await provisioner.applyTemplate(template, undefined, (handler) =>
        ran.push(handler)
      )
      return { error: undefined, ran }
    } catch (error) {
      return { error, ran }
    }
  }

  it('stops at Files with an SPFx context, naming the handler and the reason', async () => {
    const { error, ran } = await apply(
      {
        Files: [{ Folder: '', Src: '', Url: 'a.txt' }],
        Navigation: { QuickLaunch: [] }
      },
      { spfxContext: { pageContext: { site: {}, web: {} } } }
    )

    assert.ok(error instanceof ProvisioningError)
    assert.equal(error.handler, 'Files')
    assert.match(
      error.message,
      /Files handler is not supported with an SPFx context/
    )
    // The handlers after it do not run.
    assert.deepEqual(ran, ['Files'])
  })

  it('stops at PropertyBagEntries in Node with a readable message', async () => {
    const { error } = await apply(
      {
        PropertyBagEntries: [
          { Key: 'k', Value: 'v', Indexed: false, Overwrite: true }
        ]
      },
      {}
    )

    assert.ok(error instanceof ProvisioningError)
    assert.equal(error.handler, 'PropertyBagEntries')
    assert.match(
      error.message,
      /PropertyBagEntries handler is not supported in Node/
    )
  })

  it('applies a template without logging configured, or without setup', async () => {
    // The first log line read `config.logging.prefix` unguarded: a TypeError.
    assert.equal((await apply({}, {})).error, undefined)
    await new WebProvisioner(createRecordingWeb([])).applyTemplate({})
  })
})

describe('WebProvisioner: console logging', () => {
  it('subscribes its console listener once, however often a template is applied', async () => {
    // PP365's setup wizard applies a template, then each extension: a listener
    // per call printed every line once per earlier call.
    const printed: string[] = []
    const original = { ...console }
    for (const method of ['log', 'info', 'warn', 'error'] as const)
      console[method] = (...parts: any[]) => printed.push(parts.join(' '))
    try {
      const config = {
        logging: { prefix: '(test)', activeLogLevel: LogLevel.Error }
      }
      const first = new WebProvisioner(createRecordingWeb([])).setup(config)
      await first.applyTemplate({})
      await first.applyTemplate({})
      await new WebProvisioner(createRecordingWeb([]))
        .setup(config)
        .applyTemplate({})
      Logger.write('spjs-console-marker', LogLevel.Error)
    } finally {
      Object.assign(console, original)
    }

    assert.equal(
      printed.filter((line) => line.includes('spjs-console-marker')).length,
      1
    )
  })
})

describe("WebProvisioner: the caller's web", () => {
  // PnPjs 4's `select()` sets the query on the queryable itself. A `$select`
  // left on the web made the next applyTemplate read only those properties:
  // `{sitename}` and `{webid}` stayed unresolved (found by the Security suite).
  it('reads the web without a $select the caller left on it', async () => {
    const urls: string[] = []
    const web = createRecordingWeb(urls)
    web.select('HasUniqueRoleAssignments')

    await new WebProvisioner(web).setup({}).applyTemplate({})

    assert.equal(urls[0], 'https://contoso.sharepoint.com/sites/demo/_api/web')
  })

  it('leaves no $select on the web after a handler read it', async () => {
    const urls: string[] = []
    const web = createRecordingWeb(urls)

    await new WebProvisioner(web)
      .setup({})
      .applyTemplate({ Security: { BreakRoleInheritance: true } })

    assert.ok(
      urls.some((url) =>
        decodeURIComponent(url).includes('$select=HasUniqueRoleAssignments')
      ),
      'the Security handler read the web'
    )
    assert.equal(
      web.toRequestUrl(),
      'https://contoso.sharepoint.com/sites/demo/_api/web'
    )
  })
})

describe('WebProvisioner: lists made in this setup', () => {
  it('keeps them across applyTemplate calls, as PP365 applies a template and then its extensions', async () => {
    // Forskningsmal and Veimal remove Fasesjekkliste's standard content type, a list the main
    // template creates through the same provisioner a call earlier.
    const seen: any[] = []
    const original = Lists.prototype.ProvisionObjects
    Lists.prototype.ProvisionObjects = async function (
      _web: any,
      _lists: any,
      context: any
    ) {
      seen.push({ ...context.createdLists })
      context.createdLists['Fasesjekkliste'] = true
    }
    try {
      const provisioner = new WebProvisioner(createRecordingWeb([])).setup({})
      await provisioner.applyTemplate({ Lists: [] })
      await provisioner.applyTemplate({ Lists: [] })
    } finally {
      Lists.prototype.ProvisionObjects = original
    }

    assert.deepEqual(seen, [{}, { Fasesjekkliste: true }])
  })
})
