/* eslint-disable unicorn/empty-brace-spaces */
import { ConsoleListener, Logger, LogLevel } from '@pnp/logging'
import {
  DefaultHandlerMap,
  DefaultHandlerSort,
  Handler
} from './handlers/exports'
import { HandlerBase } from './handlers/handlerbase'
import { Lists } from './handlers/lists'
import { IProvisioningConfig } from './provisioningconfig'
import { ProvisioningContext } from './provisioningcontext'
import { ProvisioningError } from './provisioningerror'
import { Schema } from './schema'
import { TokenHelper } from './util/tokenhelper'
import { extractWebUrl, IWeb, Site, SPInstance, Web } from '@pnp/sp/presets/all'
import '@pnp/sp/presets/all'

/**
 * Whether a provisioner in this process has subscribed the console listener. Every
 * `applyTemplate` sets up logging, and a listener per call printed each line once per earlier
 * call (PP365's setup wizard applies a template, then each extension).
 */
let consoleListenerSubscribed = false

/**
 * Root class of Provisioning
 */
export class WebProvisioner {
  public handlerMap: Record<string, HandlerBase>
  private context: ProvisioningContext = new ProvisioningContext()
  private config: IProvisioningConfig
  /**
   * Whether `loadContext` read the web's lists in this `applyTemplate`: without them `context.lists`
   * is stale, and a list that exists would look missing.
   */
  private listsLoaded = false
  /**
   * Creates a new instance of the Provisioner class
   *
   * @param web - The Web instance to which we want to apply templates
   * @param handlerSort - A set of handlers we want to apply. The keys of the map need to match the property names in the template
   */
  constructor(
    private web: IWeb,
    public handlerSort: Record<Handler, number> = DefaultHandlerSort
  ) {}

  private async onSetup() {
    if (this.config?.logging) {
      if (!consoleListenerSubscribed) {
        Logger.subscribe(ConsoleListener())
        consoleListenerSubscribed = true
      }
      Logger.activeLogLevel = this.config.logging.activeLogLevel
    }
    this.handlerMap = DefaultHandlerMap(this.config)
    // A fresh queryable: PnPjs 4's `select()` sets the query on the instance itself, so a
    // `$select` left on the caller's web (by a handler of an earlier run, or by the caller) would
    // return only those properties, and `{sitename}`, `{webid}` and the server-relative URL
    // would be missing.
    this.context.web = await Web(this.web)()
    await this.loadContext()
  }

  /**
   * Loads the context values that tokens depend on: the web URL, the site
   * collection id (`{sitecollectionid}`), the default term store id
   * (`{sitecollectiontermstoreid}`) and the existing lists (`{listid:...}`), so
   * tokens resolve in handlers that run before `Lists` (e.g. `SiteFields`).
   * Each lookup fails soft: a warning is logged and the token stays unresolved.
   */
  private async loadContext(): Promise<void> {
    const webUrl = extractWebUrl(this.web.toUrl())
    this.context.webUrl = webUrl
    try {
      const site = await Site([this.web, webUrl]).select('Id')<{ Id: string }>()
      this.context.siteId = site.Id
    } catch (error) {
      this.logWarning(
        'Failed to load the site collection id, {sitecollectionid} falls back to the web id',
        error
      )
    }
    try {
      // PnPjs v4 dropped the `taxonomy` module, so the default term store is
      // read straight from the same endpoint v3's `TermStore` factory used.
      const termStore = await SPInstance(
        [this.web, webUrl],
        '_api/v2.1/termstore'
      ).select('id')<{ id: string }>()
      this.context.termStoreId = termStore.id
    } catch (error) {
      this.logWarning(
        'Failed to load the default term store id, {sitecollectiontermstoreid} will not be resolved',
        error
      )
    }
    this.listsLoaded = false
    try {
      // An explicit page size - without `.top()` SharePoint caps the
      // collection at 100 items and `{listid:...}` tokens silently miss lists.
      const lists = await this.web.lists.select('Id', 'Title').top(5000)<
        Array<{ Id: string; Title: string }>
      >()
      this.context.lists = lists.reduce((object, l) => {
        object[l.Title] = l.Id
        return object
      }, {} as { [key: string]: string })
      this.listsLoaded = true
    } catch (error) {
      this.logWarning('Failed to load lists for {listid:...} tokens', error)
    }
  }

  /**
   * Creates, right before `SiteFields` runs, the lists of this template that a site column's
   * `{listid:}` names (docs/specs/sitefields-re-apply.md, S1). A list's id cannot be chosen, so the list must
   * exist before the column's XML can name it: created by `Lists` after `SiteFields`, the column
   * would hold the literal token, and so would every copy of it made in the same run (content
   * types, list FieldRefs). Only lists the template's `Lists` declares, by title, with the create
   * path `Lists` uses; `Lists` then finds them existing, made in this setup. Runs only when
   * `SiteFields` and `Lists` both run, `SiteFields` first; a title the template declares twice is
   * created once, from its first entry. A list that cannot be created is an error, and the run
   * goes on as without this step (`Lists` creates it, `SiteFields` logs the token).
   *
   * @param template - The template
   * @param operations - The handlers that run, in order
   */
  private async createListsForSiteFields(
    template: Schema,
    operations: string[]
  ): Promise<void> {
    const siteFields = operations.indexOf('SiteFields')
    if (siteFields < 0 || operations.indexOf('Lists') < siteFields) return
    const tokenHelper = new TokenHelper(this.context, this.config || {})
    const titles: string[] = []
    for (const xml of template.SiteFields || [])
      for (const token of tokenHelper.unresolvedTokens(xml)) {
        const title = /^{listid:([^{}]*)}$/.exec(token)
        if (title && !titles.includes(title[1])) titles.push(title[1])
      }
    const lists = (template.Lists || []).filter(
      (lc, index, all) =>
        titles.includes(lc.Title) &&
        all.findIndex((other) => other.Title === lc.Title) === index
    )
    if (lists.length === 0) return
    if (!this.listsLoaded) {
      this.log(
        LogLevel.Warning,
        'createListsForSiteFields',
        `The web's lists could not be read, so no list is created before SiteFields (${titles.join(
          ', '
        )})`
      )
      return
    }
    for (const lc of lists) {
      try {
        await (this.handlerMap.Lists as Lists).createList(
          Web(this.web),
          lc,
          this.context
        )
        this.context.preCreatedLists[lc.Title] = true
        this.log(
          LogLevel.Info,
          'createListsForSiteFields',
          `List ${lc.Title} created before SiteFields: a site column's {listid:${lc.Title}} names it`
        )
      } catch (error) {
        this.log(
          LogLevel.Error,
          'createListsForSiteFields',
          `Failed to create list ${lc.Title} before SiteFields, so {listid:${
            lc.Title
          }} in a site column is not resolved: ${
            (error && error.message) || error
          }`,
          error
        )
      }
    }
  }

  /**
   * Warns, when a run fails, about each list created before `SiteFields` that `Lists` has not
   * reached: it exists without its fields, views and content types, and a later run (a new
   * provisioner) no longer counts it as made in this setup.
   */
  private warnPreCreatedLists(): void {
    for (const title of Object.keys(this.context.preCreatedLists || {}))
      this.log(
        LogLevel.Warning,
        'applyTemplate',
        `List ${title} was created before SiteFields, but the run failed before Lists reached it: it exists without its fields, views and content types, and a re-run will not count it as made in this setup.`
      )
  }

  /**
   * Logs a warning from `loadContext` through the pnp logger
   *
   * @param message - Message
   * @param data - Data (e.g. the error)
   */
  private logWarning(message: string, data?: any): void {
    this.log(LogLevel.Warning, 'loadContext', message, data)
  }

  /**
   * Logs through the pnp logger, as the handlers do
   *
   * @param level - Level
   * @param scope - Scope
   * @param message - Message
   * @param data - Data (e.g. the error)
   */
  private log(
    level: LogLevel,
    scope: string,
    message: string,
    data?: any
  ): void {
    Logger.log({
      message: `${
        this.config?.logging?.prefix ?? ''
      } (WebProvisioner): (${scope}): ${message}`,
      data,
      level
    })
  }

  /**
   * Applies the supplied template to the web used to create this Provisioner instance
   *
   * @param template - The template to apply
   * @param handlers - A set of handlers we want to apply
   * @param progressCallback - Callback for progress updates
   */
  public async applyTemplate(
    template: Schema,
    handlers?: string[],
    progressCallback?: (handler: Handler) => void
  ): Promise<any> {
    Logger.log({
      message: `${
        this.config?.logging?.prefix ?? ''
      } (WebProvisioner): (applyTemplate): Applying template to web`,
      data: { handlers },
      level: LogLevel.Info
    })
    await this.onSetup()
    this.context.preCreatedLists = {}

    let operations = Object.getOwnPropertyNames(template).sort(
      (name1: string, name2: string) => {
        const sort1 = this.handlerSort.hasOwnProperty(name1)
          ? this.handlerSort[name1]
          : 99
        const sort2 = this.handlerSort.hasOwnProperty(name2)
          ? this.handlerSort[name2]
          : 99
        return sort1 - sort2
      }
    )

    if (handlers) operations = operations.filter((op) => handlers.includes(op))

    operations = operations.filter((name) => this.handlerMap[name])

    let currentHandler: string
    try {
      await operations.reduce((chain: any, name: Handler) => {
        const handler = this.handlerMap[name]
        return chain.then(async () => {
          if (progressCallback) {
            progressCallback(name)
          }
          currentHandler = name
          // Inside the chain, so a handler before SiteFields (Taxonomy) that fails leaves no list.
          if (name === 'SiteFields')
            await this.createListsForSiteFields(template, operations)
          // Each handler gets a fresh queryable for the web: one that reads it with
          // `web.select(...)()` sets the query on its own copy, not on the caller's web.
          return handler.ProvisionObjects(
            Web(this.web),
            template[name],
            this.context
          )
        })
      }, Promise.resolve())
    } catch (error) {
      this.warnPreCreatedLists()
      throw new ProvisioningError(currentHandler, error)
    }
  }

  /**
   * Sets up the web provisioner
   *
   * @param config - Provisioning config
   */
  public setup(config: IProvisioningConfig): WebProvisioner {
    this.config = config
    return this
  }
}
