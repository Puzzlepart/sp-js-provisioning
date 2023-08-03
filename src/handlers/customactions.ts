import { IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { ICustomAction } from '../schema'
import { HandlerBase } from './handlerbase'
import { createBatch } from '@pnp/sp/batching'

/**
 * Describes the Custom Actions Object Handler
 */
export class CustomActions extends HandlerBase {
  /**
   * Creates a new instance of the ObjectCustomActions class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super('CustomActions', config)
  }

  /**
   * Provisioning Custom Actions
   *
   * @param web - The web
   * @param customactions - The Custom Actions to provision
   */
  public async ProvisionObjects(
    web: IWeb,
    customActions: ICustomAction[]
  ): Promise<void> {
    super.scope_started()
    try {
      const existingActions = await web.userCustomActions
        .select('Title')<{ Title: string }[]>()

      const [batch, execute] = createBatch(web)

      customActions
        .filter((action) => {
          return !existingActions.some(
            (existingAction) => existingAction.Title === action.Title
          )
        })
        .map((action) => {
          web.userCustomActions.using(batch).add(action)
        })

      await execute()
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }
}
