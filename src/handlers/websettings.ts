import { HandlerBase } from './handlerbase'
import { IWebSettings } from '../schema'
import { Web } from '@pnp/sp'
import * as omit from 'object.omit'
import { replaceUrlTokens } from '../util'
import { IProvisioningConfig } from '../provisioningconfig'

/**
 * Describes the WebSettings Object Handler
 */
export class WebSettings extends HandlerBase {
  /**
   * Creates a new instance of the WebSettings class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super('WebSettings', config)
  }

  /**
   * Provisioning WebSettings
   *
   * @param web - The web
   * @param settings - The settings
   */
  public async ProvisionObjects(
    web: Web,
    settings: IWebSettings
  ): Promise<void> {
    super.scope_started()
    for (const key of Object.keys(settings).filter(
      (key) => typeof settings[key] === 'string'
    )) {
      const value: string = replaceUrlTokens(<any>settings[key], this.config)
      super.log_info('ProvisionObjects', `Setting value of ${key} to ${value}.`)
      settings[key] = value
    }
    try {
      if (settings.WelcomePage) {
        super.log_info(
          'ProvisionObjects',
          `Setting value of WelcomePage to ${settings.WelcomePage}.`
        )
        await Promise.all([
          web.rootFolder.update({ WelcomePage: settings.WelcomePage }),
          web.update(omit(settings, 'WelcomePage'))
        ])
      }
      super.scope_ended()
    } catch (error) {
      super.scope_ended()
      throw error
    }
  }
}
