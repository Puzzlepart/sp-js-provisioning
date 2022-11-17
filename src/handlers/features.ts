import { Web } from '@pnp/sp'
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
   * Provisioning features
   *
   * @param web - The web
   * @param features - The features to provision
   */
  public async ProvisionObjects(web: Web, features: IFeature[]): Promise<void> {
    super.scope_started()
    try {
      await features.reduce((chain, feature) => {
        return feature.deactivate
          ? chain.then(() => web.features.remove(feature.id, feature.force))
          : chain.then(() => web.features.add(feature.id, feature.force))
      }, Promise.resolve<any>({}))
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }
}
