import { HandlerBase } from './handlerbase'
import { ICustomAction } from '../schema'
import { IProvisioningConfig } from '../provisioningconfig'
import { IWeb } from '@pnp/sp/webs'

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
    super(CustomActions.name, config)
  }

  /**
   * Provisioning Custom Actions
   *
   * @param web - The web
   * @param customActions - The Custom Actions to provision
   */
  public async ProvisionObjects(
    web: IWeb,
    customActions: ICustomAction[]
  ): Promise<void> {
    super.scope_started()
    try {
      const existingActions = await web.userCustomActions
        .select('Title')
        .get<{ Title: string }[]>()

      const batch = web.createBatch()

      customActions
        .filter((action) => {
          return !existingActions.some(
            (existingAction) => existingAction.Title === action.Title
          )
        })
        .map((action) => {
          web.userCustomActions.inBatch(batch).add(action)
        })

      await batch.execute()
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }
}
