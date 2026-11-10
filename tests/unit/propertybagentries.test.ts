import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { LogLevel } from '@pnp/logging'
import {
  decodePropertyKey,
  PropertyBagEntries
} from '../../src/handlers/propertybagentries'
import { IPropertyBagEntry } from '../../src/schema'
import { base64EncodeString } from '../../src/util'
import { CsomFailure, CSOM_PATH, FakeCsom, ICsomAction } from '../helpers/csom'
import { captureLogs, levelOf } from '../helpers/logs'
import { createWireWeb } from '../helpers/transport'

const KEY_A = base64EncodeString('GtA')
const KEY_B = base64EncodeString('GtTestPackage')

/**
 * A web whose property bag REST returns as `bag` (keys encoded as REST encodes them), and whose
 * CSOM requests are answered by a fake that records them.
 *
 * @param bag - The property bag, as REST returns it
 * @param options - A failing read, and actions SharePoint refuses
 */
function createSite(
  bag: Record<string, any>,
  options: {
    readFails?: boolean
    fail?: (action: ICsomAction) => CsomFailure | undefined
  } = {}
) {
  const csom = new FakeCsom((action) => {
    const refused = options.fail && options.fail(action)
    if (refused) throw refused
    return undefined
  })
  const { web, requests } = createWireWeb((request) => {
    if (request.path === '_api/web/allproperties') {
      if (options.readFails)
        throw Object.assign(new Error('Something went wrong'), { status: 500 })
      return { 'odata.metadata': 'ignored', ...bag }
    }
    if (request.path === CSOM_PATH) return csom.respond(request.body)
    return undefined
  })
  return { web, requests, csom }
}

/**
 * Runs the handler and returns what it logged.
 *
 * @param site - The fake web
 * @param entries - The template's entries
 */
async function provision(
  site: ReturnType<typeof createSite>,
  entries: IPropertyBagEntry[]
) {
  const { entries: logged } = await captureLogs(() =>
    new PropertyBagEntries({ spfxContext: {} }).ProvisionObjects(
      site.web,
      entries
    )
  )
  return {
    logged,
    errors: logged
      .filter((entry) => entry.level === LogLevel.Error)
      .map((entry) => entry.message.replace(/^.*?\): \(\w+\): /, ''))
  }
}

const BAG = 'Web.AllProperties'

describe('PropertyBagEntries', () => {
  it('writes a missing entry, an existing one only with Overwrite and another value, in one request', async () => {
    const site = createSite({
      GtSame: 'v',
      GtKept: 'set by an admin',
      GtOverwritten: 'old'
    })

    const { errors } = await provision(site, [
      { Key: 'GtNew', Value: 'ny' },
      { Key: 'GtSame', Value: 'v', Overwrite: true },
      { Key: 'GtKept', Value: 'from the template' },
      { Key: 'GtOverwritten', Value: 'new', Overwrite: true }
    ])

    assert.deepEqual(errors, [])
    assert.equal(site.csom.requests.length, 1)
    assert.deepEqual(site.csom.ran, [
      `call ${BAG}.SetFieldValue("GtNew", "ny")`,
      `call ${BAG}.SetFieldValue("GtOverwritten", "new")`,
      'call Web.Update()'
    ])
  })

  it('adds an indexed key to the ones the web has, as SharePoint writes them', async () => {
    const site = createSite({
      vti_x005f_indexedpropertykeys: `${KEY_A}|`,
      GtTestPackage: 'x'
    })

    await provision(site, [
      { Key: 'GtA', Value: 'a', Indexed: true },
      { Key: 'GtTestPackage', Value: 'x', Indexed: true }
    ])

    assert.deepEqual(site.csom.ran, [
      `call ${BAG}.SetFieldValue("GtA", "a")`,
      `call ${BAG}.SetFieldValue("vti_indexedpropertykeys", "${KEY_A}|${KEY_B}|")`,
      'call Web.Update()'
    ])
  })

  it('sends nothing when the web has every entry and indexed key', async () => {
    const site = createSite({
      vti_x005f_indexedpropertykeys: `${KEY_B}|${KEY_A}|`,
      GtA: 'a',
      GtTestPackage: 'other'
    })

    const { logged } = await provision(site, [
      { Key: 'GtA', Value: 'a', Indexed: true, Overwrite: true },
      { Key: 'GtTestPackage', Value: 'x', Indexed: true }
    ])

    assert.deepEqual(site.csom.requests, [])
    assert.equal(
      levelOf(logged, /The property bag is unchanged\./),
      LogLevel.Info
    )
    assert.equal(
      levelOf(
        logged,
        /Property bag entry GtTestPackage exists with another value and is left as it is \(Overwrite is not set\)\./
      ),
      LogLevel.Info
    )
  })

  it('logs a refused write as an error that names custom script, and does not throw', async () => {
    const site = createSite(
      {},
      {
        fail: (action) =>
          action.name === 'SetFieldValue'
            ? new CsomFailure(
                'Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))',
                'System.UnauthorizedAccessException',
                -2147024891
              )
            : undefined
      }
    )

    const { errors } = await provision(site, [
      { Key: 'GtA', Value: 'a', Indexed: true }
    ])

    assert.deepEqual(errors, [
      'Property bag entries GtA, vti_indexedpropertykeys were not written: SharePoint refused the write, as it does on a site that blocks custom script (DenyAddAndCustomizePages; a SharePoint admin can allow it): Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED)) (System.UnauthorizedAccessException)'
    ])
  })

  it('logs another refusal and a failed read as errors, without a request for the read', async () => {
    const refused = createSite(
      {},
      { fail: () => new CsomFailure('Verdien er for lang.') }
    )
    const { errors } = await provision(refused, [{ Key: 'GtA', Value: 'a' }])
    assert.deepEqual(errors, [
      'Property bag entries GtA were not written: SharePoint refused the write: Verdien er for lang. (Microsoft.SharePoint.SPException)'
    ])

    const unread = createSite({}, { readFails: true })
    const failed = await provision(unread, [
      { Key: 'GtA', Value: 'a' },
      { Key: 'GtB', Value: 'b' }
    ])
    assert.deepEqual(failed.errors, [
      "Failed to read the web's property bag, so no entry is written (GtA, GtB): Something went wrong"
    ])
    assert.deepEqual(unread.csom.requests, [])
  })

  it('decodes the keys REST encodes', () => {
    assert.equal(
      decodePropertyKey('vti_x005f_indexedpropertykeys'),
      'vti_indexedpropertykeys'
    )
    assert.equal(decodePropertyKey('Gt_x002e_Key'), 'Gt.Key')
  })
})
