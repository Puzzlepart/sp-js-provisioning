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
   * Creates the batch the custom actions are added through. Isolated in an
   * overridable method so the batching seam can be replaced with a fake in
   * tests - `createBatch` needs a real queryable and cannot be stubbed
   * otherwise.
   *
   * @param web - The web to batch against
   */
  protected createBatch(
    web: IWeb
  ): [(instance: any) => any, () => Promise<void>] {
    return createBatch(web) as [(instance: any) => any, () => Promise<void>]
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
      const existingActions = await web.userCustomActions.select('Title')<
        { Title: string }[]
      >()

      const [batch, execute] = this.createBatch(web)

      // In PnPjs v4 `execute()` only resolves the batch itself - the promise of
      // each batched request has to be awaited separately or a failing request
      // rejects unobserved. Both are awaited in a single `Promise.all` so that
      // a rejecting `execute()` does not leave the per-request promises
      // unobserved.
      const requests = customActions
        .filter((action) => {
          return !existingActions.some(
            (existingAction) => existingAction.Title === action.Title
          )
        })
        .map((action) => web.userCustomActions.using(batch).add(action))

      // Typed as unknown because `execute()` resolves void while the batched
      // adds resolve with the created custom actions.
      const pending: Promise<unknown>[] = [execute(), ...requests]
      await Promise.all(pending)
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }
}
