import { HandlerBase } from './handlerbase'
import { IPropertyBagEntry } from '../schema'
import * as Util from '../util'
import { Logger, LogLevel } from '@pnp/logging'
import { IProvisioningConfig } from '../provisioningconfig'
import { IWeb } from '@pnp/sp/webs'

/**
 * Describes the PropertyBagEntries Object Handler
 */
export class PropertyBagEntries extends HandlerBase {
  /**
   * Creates a new instance of the PropertyBagEntries class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super(PropertyBagEntries.name, config)
  }

  /**
   * Provisioning property bag entries
   *
   * @param web - The web
   * @param entries - The property bag entries to provision
   */
  public ProvisionObjects(
    web: IWeb,
    entries: IPropertyBagEntry[]
  ): Promise<void> {
    super.scope_started()
    return new Promise<any>((resolve, reject) => {
      if (Util.isNode()) {
        Logger.write(
          'PropertyBagEntries Handler not supported in Node.',
          LogLevel.Error
        )
        reject()
      } else if (this.config.spfxContext) {
        Logger.write(
          'PropertyBagEntries Handler not supported in SPFx.',
          LogLevel.Error
        )
        reject()
      } else {
        web.get().then(({ ServerRelativeUrl }) => {
          const context = new SP.ClientContext(ServerRelativeUrl),
            spWeb = context.get_web(),
            propertyBag = spWeb.get_allProperties(),
            indexProps = []
          for (const entry of entries.filter((entry) => entry.Overwrite)) {
            propertyBag.set_item(entry.Key, entry.Value)
            if (entry.Indexed) {
              indexProps.push(Util.base64EncodeString(entry.Key))
            }
          }
          spWeb.update()
          context.load(propertyBag)
          context.executeQueryAsync(
            () => {
              if (indexProps.length > 0) {
                propertyBag.set_item(
                  'vti_indexedpropertykeys',
                  indexProps.join('|')
                )
                spWeb.update()
                context.executeQueryAsync(
                  () => {
                    super.scope_ended()
                    resolve(true)
                  },
                  () => {
                    super.scope_ended()
                    reject()
                  }
                )
              } else {
                super.scope_ended()
                resolve(true)
              }
            },
            () => {
              super.scope_ended()
              reject()
            }
          )
        })
      }
    })
  }
}
