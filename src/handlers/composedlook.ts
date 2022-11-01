import { IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { IComposedLook } from '../schema'
import { makeUrlRelative, replaceUrlTokens } from '../util'
import { HandlerBase } from './handlerbase'

/**
 * Describes the Composed Look Object Handler
 */
export class ComposedLook extends HandlerBase {
  /**
   * Creates a new instance of the ObjectComposedLook class
   */
  constructor(config: IProvisioningConfig) {
    super(ComposedLook.name, config)
  }

  /**
   * Provisioning Composed Look
   *
   * @param web - The web
   * @param composedLook - The Composed look to provision
   */
  public async ProvisionObjects(
    web: IWeb,
    composedLook: IComposedLook
  ): Promise<void> {
    super.scope_started()
    try {
      await web.applyTheme(
        makeUrlRelative(
          replaceUrlTokens(composedLook.ColorPaletteUrl, this.config)
        ),
        makeUrlRelative(
          replaceUrlTokens(composedLook.FontSchemeUrl, this.config)
        ),
        composedLook.BackgroundImageUrl
          ? makeUrlRelative(
              replaceUrlTokens(composedLook.BackgroundImageUrl, this.config)
            )
          : null,
        false
      )
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }
}
