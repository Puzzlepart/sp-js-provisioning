/* eslint-disable unicorn/empty-brace-spaces */
import { ConsoleListener, Logger, LogLevel } from '@pnp/logging'
import {
  DefaultHandlerMap,
  DefaultHandlerSort,
  Handler
} from './handlers/exports'
import { HandlerBase } from './handlers/handlerbase'
import { IProvisioningConfig } from './provisioningconfig'
import { ProvisioningContext } from './provisioningcontext'
import { ProvisioningError } from './provisioningerror'
import { Schema } from './schema'
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
    } catch (error) {
      this.logWarning('Failed to load lists for {listid:...} tokens', error)
    }
  }

  /**
   * Logs a warning through the pnp logger
   *
   * @param message - Message
   * @param data - Data (e.g. the error)
   */
  private logWarning(message: string, data?: any): void {
    Logger.log({
      message: `${
        this.config?.logging?.prefix ?? ''
      } (WebProvisioner): (loadContext): ${message}`,
      data,
      level: LogLevel.Warning
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
        return chain.then(() => {
          if (progressCallback) {
            progressCallback(name)
          }
          currentHandler = name
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
