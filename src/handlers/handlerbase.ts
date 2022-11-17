import { Logger, LogLevel } from '@pnp/logging'
import { Web } from '@pnp/sp'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { Handler } from './exports'

/**
 * Describes the Object Handler Base
 */
export class HandlerBase {
  public config: IProvisioningConfig = {}
  private name: Handler

  /**
   * Creates a new instance of the ObjectHandlerBase class
   *
   * @param name - Name
   * @param config - Config
   */
  constructor(name: Handler, config: IProvisioningConfig = {}) {
    this.name = name
    this.config = config
  }

  /**
   * Provisioning objects
   */
  public ProvisionObjects(
    web: Web,
    templatePart: any,
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    _context?: ProvisioningContext
  ): Promise<void> {
    Logger.log({
      data: templatePart,
      level: LogLevel.Warning,
      message: `Handler ${
        this.name
      } for web [${web.toUrl()}] does not override ProvisionObjects.`
    })
    return Promise.resolve()
  }

  /**
   * Writes to Logger when scope has started
   */
  public scope_started() {
    this.log_info('ProvisionObjects', 'Code execution scope started')
  }

  /**
   * Writes to Logger when scope has stopped
   *
   * @param error - Error
   */
  public scope_ended(error?: Error) {
    if (error)
      this.log_error(
        'ProvisionObjects',
        `Code execution scope ended with error: ${error.message}`
      )
    else this.log_info('ProvisionObjects', 'Code execution scope ended')
  }

  /**
   * Writes to Logger
   *
   * @param scope - Scope
   * @param message - Message
   * @param data - Data
   */
  public log_info(scope: string, message: string, data?: any) {
    const prefix =
      this.config.logging && this.config.logging.prefix
        ? `${this.config.logging.prefix} `
        : ''
    Logger.log({
      message: `${prefix}(${this.name}): (${scope}): ${message}`,
      data: data,
      level: LogLevel.Info
    })
  }

  /**
   * Writes a warning to the logger
   *
   * @param scope - Scope
   * @param message - Message
   * @param data - Data
   */
  public log_warn(scope: string, message: string, data?: any) {
    const prefix =
      this.config.logging && this.config.logging.prefix
        ? `${this.config.logging.prefix} `
        : ''
    Logger.log({
      message: `${prefix}(${this.name}): (${scope}): ${message}`,
      data: data,
      level: LogLevel.Warning
    })
  }

  /**
   * Writes an error to the logger
   *
   * @param scope - Scope
   * @param message - Message
   * @param data - Data
   */
  public log_error(scope: string, message: string, data?: any) {
    const prefix =
      this.config.logging && this.config.logging.prefix
        ? `${this.config.logging.prefix} `
        : ''
    Logger.log({
      message: `${prefix}(${this.name}): (${scope}): ${message}`,
      data: data,
      level: LogLevel.Error
    })
  }
}
