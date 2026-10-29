import { HandlerBase } from './handlerbase'
import { IWebSettings } from '../schema'
import { IWeb } from '@pnp/sp/presets/all'
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
   * Provisioning WebSettings: `WelcomePage` on the web's root folder, every other setting on the
   * web. Each is updated only when there is something to set.
   *
   * @param web - The web
   * @param settings - The settings
   */
  public async ProvisionObjects(
    web: IWeb,
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
    // Not `object.omit`: imported as a namespace, it is no function under Node's ESM interop.
    const { WelcomePage, ...properties } = settings
    try {
      const updates: Array<Promise<unknown>> = []
      if (WelcomePage) {
        super.log_info(
          'ProvisionObjects',
          `Setting value of WelcomePage to ${WelcomePage}.`
        )
        updates.push(web.rootFolder.update({ WelcomePage }))
      }
      if (Object.keys(properties).length > 0)
        updates.push(web.update(properties))
      await Promise.all(updates)
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }
}
