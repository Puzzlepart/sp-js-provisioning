// we need to import HandlerBase & TypedHash to avoid naming issues in ts transpile
import { Schema } from './schema'
import { HandlerBase } from './handlers/handlerbase'
import { Logger, LogLevel, ConsoleListener } from '@pnp/logging'
import { DefaultHandlerMap, DefaultHandlerSort } from './handlers/exports'
import { ProvisioningContext } from './provisioningcontext'
import { IProvisioningConfig } from './provisioningconfig'
import { IWeb } from '@pnp/sp/webs'
import '@pnp/sp/webs'

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
    public handlerSort: Record<string, number> = DefaultHandlerSort
  ) { }

  private async onSetup() {
    if (this.config?.logging) {
      Logger.subscribe(ConsoleListener())
      Logger.activeLogLevel = this.config.logging.activeLogLevel
    }
    this.handlerMap = DefaultHandlerMap(this.config)
    this.context.web = await this.web()
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
    progressCallback?: (message: string) => void
  ): Promise<any> {

    Logger.log({
      message: `${this.config.logging.prefix} (WebProvisioner): (applyTemplate): Applying template to web`,
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

    if (handlers) {
      operations = operations.filter((op) => handlers.includes(op))
    }

    operations = operations.filter((name) => this.handlerMap[name])

    try {
      await operations.reduce((chain: any, name: string) => {
        const handler = this.handlerMap[name]
        return chain.then(() => {
          if (progressCallback) {
            progressCallback(name)
          }
          return handler.ProvisionObjects(
            this.web,
            template[name],
            this.context
          )
        })
      }, Promise.resolve())
    } catch (error) {
      throw error
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
