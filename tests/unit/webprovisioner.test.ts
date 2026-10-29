import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
// Importing the real `@pnp/sp` here is deliberate: the term store lookup is a
// hand-written url, and only the real queryable proves what url is sent.
import { Web } from '@pnp/sp/presets/all'
import { RejectOnError, ResolveOnData } from '@pnp/queryable'
import { WebProvisioner } from '../../src/webprovisioner'
import { DefaultHandlerSort } from '../../src/handlers/exports'
import { Lists } from '../../src/handlers/lists'
import { Features } from '../../src/handlers/features'
import { Taxonomy } from '../../src/handlers/taxonomy'
import { ProvisioningError } from '../../src/provisioningerror'
import { Logger, LogLevel } from '@pnp/logging'
import { captureLogs, levelOf } from '../helpers/logs'

/**
 * Records every request url the provisioner produces and short-circuits the
 * request by supplying a result in the `pre` moment (an empty one unless
 * `respond` gives another), so no network call is made. Observers are
 * inherited by the child queryables the provisioner builds from this web.
 *
 * @param urls - Array every requested url is pushed to
 * @param respond - The result for a request, `undefined` for an empty one; it may throw
 */
function createRecordingWeb(
  urls: string[],
  respond?: (url: string, init: any) => any
): any {
  const recording = (instance: any) => {
    instance.on.pre(async (url: string, init: any) => {
      urls.push(url)
      const result = respond ? respond(url, init) : undefined
      return [url, init, result === undefined ? {} : result]
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

describe('WebProvisioner: a list a site column names, created before SiteFields', () => {
  const later = 'SPJS Kilder'
  const lookup = (title = later) =>
    `<Field ID="{69b1c001-3ffe-414a-8774-507a9da1d89e}" Name="GtSource" StaticName="GtSource" DisplayName="Kilde" Type="Lookup" List="{listid:${title}}" ShowField="Title" />`
  const template = (title = later): any => ({
    SiteFields: [lookup(title)],
    Lists: [
      {
        Title: later,
        Description: 'Kilder',
        Template: 100,
        ContentTypesEnabled: false
      }
    ]
  })

  /**
   * Applies a template to a recording web that answers like an empty web (no lists, no columns),
   * with an id for a list or a column it adds. Lists is stubbed: it records what the context holds
   * when it starts, and sends nothing.
   *
   * @param schema - The template
   * @param options - Handlers to run, a handler sort, requests that fail, and a handler that throws
   */
  async function run(
    schema: any,
    options: {
      handlers?: string[]
      handlerSort?: any
      failAdd?: boolean
      failListsRead?: boolean
      throwing?: typeof Features | typeof Taxonomy
    } = {}
  ) {
    const requests: Array<{ method: string; path: string; body?: any }> = []
    const respond = (url: string, init: any) => {
      const method = (init && init.method) || 'GET'
      const path = decodeURIComponent(url).split('?')[0]
      const body = init && init.body ? JSON.parse(init.body) : undefined
      requests.push({ method, path, body })
      if (path.endsWith('/_api/web/lists')) {
        if (method === 'GET') {
          if (options.failListsRead) throw new Error('Throttled')
          return []
        }
        if (options.failAdd) throw new Error('Listen kan ikke opprettes')
        return { Id: 'later-guid', Title: body.Title }
      }
      if (/\/_api\/web\/(available)?fields$/.test(path)) return []
      if (path.endsWith('/createfieldasxml')) return { Id: 'field-guid' }
      if (path.endsWith('/_api/web'))
        return {
          Id: 'web-guid',
          Title: 'Demo',
          ServerRelativeUrl: '/sites/demo'
        }
      return undefined
    }
    const lists: any[] = []
    const thrower = options.throwing && options.throwing.prototype
    const originalThrower = thrower && thrower.ProvisionObjects
    if (thrower)
      thrower.ProvisionObjects = async () => {
        throw new Error('Handler failed')
      }
    const original = Lists.prototype.ProvisionObjects
    Lists.prototype.ProvisionObjects = async function (
      _web: any,
      _lists: any,
      context: any
    ) {
      lists.push({
        preCreatedLists: { ...context.preCreatedLists },
        createdLists: { ...context.createdLists },
        requests: requests.length
      })
    }
    try {
      const provisioner = new WebProvisioner(
        createRecordingWeb([], respond),
        options.handlerSort
      ).setup({})
      let error: any
      const { entries } = await captureLogs(() =>
        provisioner.applyTemplate(schema, options.handlers).catch((e) => {
          if (!thrower) throw e
          error = e
        })
      )
      return { requests, lists, entries, error }
    } finally {
      Lists.prototype.ProvisionObjects = original
      if (thrower) thrower.ProvisionObjects = originalThrower
    }
  }
  const adds = (requests: any[]) =>
    requests.filter(
      (r) => r.method === 'POST' && r.path.endsWith('/_api/web/lists')
    )
  const columns = (requests: any[]) =>
    requests.filter((r) => r.path.endsWith('/createfieldasxml'))
  const errors = (entries: any[]) =>
    entries
      .filter((entry) => entry.level === LogLevel.Error)
      .map((entry) => entry.message)

  it('creates it, with its settings, before the column, which then names its id', async () => {
    const { requests, lists, entries } = await run(template())

    const [add] = adds(requests)
    const [column] = columns(requests)
    assert.ok(add, 'the list was not created')
    assert.deepEqual(
      [add.body.Title, add.body.Description, add.body.BaseTemplate],
      [later, 'Kilder', 100]
    )
    assert.ok(requests.indexOf(add) < requests.indexOf(column))
    assert.match(column.body.parameters.SchemaXml, /List="later-guid"/)
    // Lists finds it existing, made in this setup.
    assert.deepEqual(lists, [
      {
        preCreatedLists: { [later]: true },
        createdLists: { [later]: true },
        requests: requests.length
      }
    ])
    assert.deepEqual(errors(entries), [])
    assert.equal(levelOf(entries, /created before SiteFields/), LogLevel.Info)
  })

  it('creates nothing early when Lists does not run, or runs first, and the column names the token', async () => {
    for (const options of [
      { handlers: ['SiteFields'] },
      { handlerSort: { ...DefaultHandlerSort, Lists: -2 } }
    ]) {
      const { requests, entries } = await run(template(), options)

      assert.deepEqual(adds(requests), [])
      assert.match(
        columns(requests)[0].body.parameters.SchemaXml,
        /List="\{listid:SPJS Kilder\}"/
      )
      assert.equal(errors(entries).length, 1)
      assert.match(
        errors(entries)[0],
        /GtSource \(Kilde\): \{listid:SPJS Kilder\} could not be resolved/
      )
    }
  })

  it('creates no list the template does not declare', async () => {
    const { requests, entries } = await run(template('SPJS Mangler'))

    assert.deepEqual(adds(requests), [])
    assert.match(
      errors(entries)[0],
      /\{listid:SPJS Mangler\} could not be resolved/
    )
  })

  it('logs a list it cannot create, and runs on as without the step', async () => {
    const { requests, lists, entries } = await run(template(), {
      failAdd: true
    })

    assert.match(
      errors(entries)[0],
      /Failed to create list SPJS Kilder before SiteFields.*Listen kan ikke opprettes/
    )
    assert.equal(columns(requests).length, 1)
    assert.equal(lists.length, 1)
    assert.deepEqual(lists[0].preCreatedLists, {})
  })

  it("creates nothing when the web's lists could not be read: a list that exists would look missing", async () => {
    const { requests, entries } = await run(template(), {
      failListsRead: true
    })

    assert.deepEqual(adds(requests), [])
    assert.equal(
      levelOf(
        entries,
        /lists could not be read, so no list is created before SiteFields/
      ),
      LogLevel.Warning
    )
  })

  it('creates a title the template declares twice once, from its first entry, with no error', async () => {
    const schema = template()
    schema.Lists.push({ Title: later, Description: 'Andre' })
    const { requests, entries } = await run(schema)

    assert.deepEqual(
      adds(requests).map((add) => add.body.Description),
      ['Kilder']
    )
    assert.deepEqual(errors(entries), [])
  })

  it('creates nothing when a handler before SiteFields fails', async () => {
    const { requests, error } = await run(
      { Taxonomy: { TermGroups: [] }, ...template() },
      { throwing: Taxonomy }
    )

    assert.ok(error instanceof ProvisioningError)
    assert.equal(error.handler, 'Taxonomy')
    assert.deepEqual(adds(requests), [])
  })

  it('warns about a list it created when a handler between SiteFields and Lists fails', async () => {
    const { requests, lists, entries, error } = await run(
      { Features: { WebFeatures: [] }, ...template() },
      { throwing: Features }
    )

    assert.equal(error.handler, 'Features')
    assert.equal(adds(requests).length, 1)
    assert.deepEqual(lists, [])
    assert.equal(
      levelOf(
        entries,
        /List SPJS Kilder was created before SiteFields, but the run failed before Lists reached it/
      ),
      LogLevel.Warning
    )
  })
})
