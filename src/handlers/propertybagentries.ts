import { Logger, LogLevel } from '@pnp/logging'
import { IWeb } from '@pnp/sp/presets/all'
import { spPost } from '@pnp/sp/operations'
import { IProvisioningConfig } from '../provisioningconfig'
import { IPropertyBagEntry } from '../schema'
import * as Util from '../util'
import { HandlerBase } from './handlerbase'

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
    super('PropertyBagEntries', config)
  }

  /**
   * Provisioning property bag entries
   *
   * @param web - The web
   * @param entries - The property bag entries to provision
   */
  public async ProvisionObjects(
    web: IWeb,
    entries: IPropertyBagEntry[]
  ): Promise<void> {
    super.scope_started()
    
    if (Util.isNode()) {
      Logger.write(
        'PropertyBagEntries Handler not supported in Node.',
        LogLevel.Error
      )
      super.scope_ended()
      throw new Error('PropertyBagEntries Handler not supported in Node.')
    }

    try {
      const currentProps = await web.allProperties()
      const indexProps: string[] = []
      
      if (currentProps.vti_indexedpropertykeys) {
        indexProps.push(...currentProps.vti_indexedpropertykeys.split('|'))
      }

      const propertiesToUpdate: any = {}
      
      for (const entry of entries.filter((entry) => entry.Overwrite)) {
        propertiesToUpdate[entry.Key] = entry.Value
        if (entry.Indexed) {
          const encodedKey = Util.base64EncodeString(entry.Key)
          if (!indexProps.includes(encodedKey)) {
            indexProps.push(encodedKey)
          }
        }
      }

      if (indexProps.length > 0) {
        propertiesToUpdate.vti_indexedpropertykeys = indexProps.join('|')
      }

      await spPost(web.allProperties, { body: JSON.stringify(propertiesToUpdate) })
      
      super.scope_ended()
    } catch (error) {
      super.scope_ended()
      Logger.write(
        `Error setting property bag entries: ${error}`,
        LogLevel.Error
      )
      throw error
    }
  }
}
