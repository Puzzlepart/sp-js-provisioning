import { IWeb, Web } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { readableErrorMessage } from '../provisioningerror'
import { IPropertyBagEntry } from '../schema'
import { base64EncodeString } from '../util'
import {
  CsomError,
  CsomRequest,
  describeCsomError,
  executeQuery
} from '../util/csom'
import { HandlerBase } from './handlerbase'

/** The web property that lists the keys search indexes, base64 of UTF-16LE, each followed by `|`. */
export const INDEXED_KEYS = 'vti_indexedpropertykeys'

/**
 * A property bag key as REST encodes it (`vti_x005f_indexedpropertykeys`), decoded.
 *
 * @param key - The encoded key
 */
export function decodePropertyKey(key: string): string {
  return key.replace(/_x([\da-f]{4})_/gi, (_match, code: string) =>
    String.fromCharCode(Number.parseInt(code, 16))
  )
}

/**
 * The keys a `vti_indexedpropertykeys` value lists.
 *
 * @param value - The value
 */
function indexedKeys(value: any): string[] {
  return String(value || '')
    .split('|')
    .filter(Boolean)
}

/**
 * Describes the PropertyBagEntries Object Handler.
 *
 * Writes web property bag entries the way PnP does: an entry whose key the web lacks is written,
 * an existing one only with `Overwrite: true` (and only when its value differs), and `Indexed` adds
 * the key to the web's indexed keys, keeping the ones there. The bag is read through REST and
 * written in one CSOM request (`src/util/csom.ts`); nothing to write, no request. SharePoint
 * refuses the write on a site that blocks custom script (`DenyAddAndCustomizePages`, the default
 * on modern sites): that is logged as an error naming the cause, and the template goes on.
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
   * Provisioning property bag entries. A failure is logged, never thrown: the handlers after this
   * one do not depend on it.
   *
   * @param web - The web
   * @param entries - The property bag entries to provision
   */
  public async ProvisionObjects(
    web: IWeb,
    entries: IPropertyBagEntry[]
  ): Promise<void> {
    super.scope_started()
    const keys = entries.map((entry) => entry.Key).join(', ')
    let bag: Record<string, any>
    try {
      const raw = await Web(web).allProperties<Record<string, any>>()
      bag = {}
      for (const key of Object.keys(raw || {}))
        if (!/^odata\./.test(key)) bag[decodePropertyKey(key)] = raw[key]
    } catch (error) {
      super.log_error(
        'ProvisionObjects',
        `Failed to read the web's property bag, so no entry is written (${keys}): ${
          readableErrorMessage(error) ?? 'unknown error'
        }`
      )
      super.scope_ended()
      return
    }

    const request = new CsomRequest()
    const allProperties = request.property(request.web(), 'AllProperties')
    const written: string[] = []
    for (const entry of entries) {
      const value = String(entry.Value)
      const have = bag[entry.Key]
      if (have === undefined || have === null) {
        request.call(allProperties, 'SetFieldValue', entry.Key, value)
        written.push(entry.Key)
      } else if (String(have) === value)
        super.log_info(
          'ProvisionObjects',
          `Property bag entry ${entry.Key} is unchanged.`
        )
      else if (entry.Overwrite) {
        request.call(allProperties, 'SetFieldValue', entry.Key, value)
        written.push(entry.Key)
      } else
        super.log_info(
          'ProvisionObjects',
          `Property bag entry ${entry.Key} exists with another value and is left as it is (Overwrite is not set).`
        )
    }

    const current = indexedKeys(bag[INDEXED_KEYS])
    const missing = entries
      .filter((entry) => entry.Indexed)
      .map((entry) => base64EncodeString(entry.Key))
      .filter((key, index, all) => all.indexOf(key) === index)
      .filter((key) => !current.includes(key))
    if (missing.length > 0) {
      request.call(
        allProperties,
        'SetFieldValue',
        INDEXED_KEYS,
        [...current, ...missing].map((key) => `${key}|`).join('')
      )
      written.push(INDEXED_KEYS)
    }

    if (request.actionCount === 0) {
      super.log_info('ProvisionObjects', 'The property bag is unchanged.')
      super.scope_ended()
      return
    }
    request.call(request.web(), 'Update')
    try {
      await executeQuery(web, request)
      super.log_info(
        'ProvisionObjects',
        `Wrote property bag entries: ${written.join(', ')}`
      )
    } catch (error) {
      super.log_error(
        'ProvisionObjects',
        `Property bag entries ${written.join(', ')} were not written: ${
          error instanceof CsomError && error.accessDenied
            ? 'SharePoint refused the write, as it does on a site that blocks custom script (DenyAddAndCustomizePages; a SharePoint admin can allow it)'
            : 'SharePoint refused the write'
        }: ${describeCsomError(error)}`
      )
    }
    super.scope_ended()
  }
}
