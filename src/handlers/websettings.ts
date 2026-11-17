import { HandlerBase } from './handlerbase'
import { IWebSettings } from '../schema'
import { IWeb } from '@pnp/sp/presets/all'
import { replaceUrlTokens } from '../util'
import { differingSettings } from '../util/settings'
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
   * web. The web is read first and only what differs is written, so an unchanged re-run writes
   * nothing (1.6.1; the welcome page was set on every run of a PP365 project setup).
   *
   * @param web - The web (a fresh queryable: it is read without `$select`, every property)
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
      // The web's `WelcomePage` is its root folder's.
      const stored: Record<string, any> = await web()
      const updates: Array<Promise<unknown>> = []
      if (WelcomePage) {
        if (stored && stored.WelcomePage === WelcomePage)
          super.log_info(
            'ProvisionObjects',
            `WelcomePage is ${WelcomePage} already.`
          )
        else {
          super.log_info(
            'ProvisionObjects',
            `Setting value of WelcomePage to ${WelcomePage}.`
          )
          updates.push(web.rootFolder.update({ WelcomePage }))
        }
      }
      const changed = differingSettings(properties, stored)
      if (Object.keys(changed).length > 0) {
        super.log_info(
          'ProvisionObjects',
          `Updating the web: ${Object.keys(changed).join(', ')}.`
        )
        updates.push(web.update(changed))
      } else if (Object.keys(properties).length > 0)
        super.log_info(
          'ProvisionObjects',
          "The web has the template's settings already."
        )
      await Promise.all(updates)
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }
}
