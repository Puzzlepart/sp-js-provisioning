import { IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { IFeature } from '../schema'
import { HandlerBase } from './handlerbase'

/**
 * Describes the Features Object Handler
 */
export class Features extends HandlerBase {
  /**
   * Creates a new instance of the ObjectFeatures class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super('Features', config)
  }

  /**
   * Provisioning features. A feature already in the requested state is left as it is (logged):
   * SharePoint refuses to activate an active feature (500) or deactivate an inactive one (400), so
   * re-applying a template used to stop it here. `force` re-activates an active feature.
   *
   * @param web - The web
   * @param features - The features to provision
   */
  public async ProvisionObjects(
    web: IWeb,
    features: IFeature[]
  ): Promise<void> {
    super.scope_started()
    try {
      const active = new Set(
        (
          await web.features.select('DefinitionId')<
            Array<{ DefinitionId: string }>
          >()
        ).map((feature) => Features.normalise(feature.DefinitionId))
      )
      for (const feature of features) {
        const isActive = active.has(Features.normalise(feature.id))
        if (feature.deactivate && !isActive) {
          super.log_info(
            'ProvisionObjects',
            `Feature ${feature.id} is not active, nothing to deactivate.`
          )
          continue
        }
        if (!feature.deactivate && isActive && !feature.force) {
          super.log_info(
            'ProvisionObjects',
            `Feature ${feature.id} is already active.`
          )
          continue
        }
        if (feature.deactivate) {
          await web.features.remove(feature.id, feature.force)
          active.delete(Features.normalise(feature.id))
        } else {
          await web.features.add(feature.id, feature.force)
          active.add(Features.normalise(feature.id))
        }
      }
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * A feature id for comparison: lower case, without braces.
   *
   * @param id - The id
   */
  private static normalise(id: string): string {
    return String(id).toLowerCase().replace(/[{}]/g, '')
  }
}
