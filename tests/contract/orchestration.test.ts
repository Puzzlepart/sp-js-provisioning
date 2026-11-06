import assert from 'node:assert/strict'
import { after, before, describe, it } from 'node:test'
import { ClientSidePages } from '../../src/handlers/clientsidepages'
import { ComposedLook } from '../../src/handlers/composedlook'
import { ContentTypes } from '../../src/handlers/contenttypes'
import { CustomActions } from '../../src/handlers/customactions'
import { DefaultHandlerSort } from '../../src/handlers/exports'
import { Features } from '../../src/handlers/features'
import { Files } from '../../src/handlers/files'
import { Hooks } from '../../src/handlers/hooks'
import { Lists } from '../../src/handlers/lists'
import { Navigation } from '../../src/handlers/navigation'
import { PropertyBagEntries } from '../../src/handlers/propertybagentries'
import { Security } from '../../src/handlers/security'
import { SiteFields } from '../../src/handlers/sitefields'
import { Taxonomy } from '../../src/handlers/taxonomy'
import { WebSettings } from '../../src/handlers/websettings'
import { ProvisioningError } from '../../src/provisioningerror'
import { WebProvisioner } from '../../src/webprovisioner'
import { createWireWeb } from '../helpers/transport'

/**
 * Orchestration: the order WebProvisioner runs the handlers in, the `handlers` filter, a custom
 * `handlerSort` and a failing handler. Every handler's ProvisionObjects is replaced by a recorder, so
 * only WebProvisioner's own requests (the context it loads) go to the transport.
 */

const handlers: Record<string, any> = {
  ClientSidePages,
  ComposedLook,
  ContentTypes,
  CustomActions,
  Features,
  Files,
  Hooks,
  Lists,
  Navigation,
  PropertyBagEntries,
  Security,
  SiteFields,
  Taxonomy,
  WebSettings
}

/** A template with a section for every handler. */
const template = (): any =>
  Object.keys(handlers).reduce(
    (schema: any, name) => {
      schema[name] =
        name === 'Navigation' ||
        name === 'Taxonomy' ||
        name === 'Security' ||
        name === 'WebSettings' ||
        name === 'ComposedLook'
          ? {}
          : []
      return schema
    },
    { Version: '1.0', Parameters: {} }
  )

describe('contract: WebProvisioner orchestration', () => {
  const ran: string[] = []
  let failing: string | undefined
  const originals: Record<string, any> = {}

  before(() => {
    for (const [name, handler] of Object.entries(handlers)) {
      originals[name] = handler.prototype.ProvisionObjects
      handler.prototype.ProvisionObjects = async function () {
        ran.push(name)
        if (name === failing) throw new Error(`${name} failed`)
      }
    }
  })
  after(() => {
    for (const [name, handler] of Object.entries(handlers))
      handler.prototype.ProvisionObjects = originals[name]
  })

  /**
   * Applies a template and returns the handlers that ran, in order, and what was thrown.
   *
   * @param schema - The template
   * @param options - Handlers to run, a handler sort, a handler that fails
   */
  async function apply(
    schema: any,
    options: { handlers?: string[]; sort?: any; fail?: string } = {}
  ) {
    ran.length = 0
    failing = options.fail
    const { web } = createWireWeb((request) =>
      request.path.endsWith('/lists') ? [] : undefined
    )
    const provisioner = new WebProvisioner(web, options.sort).setup({})
    let error: any
    try {
      await provisioner.applyTemplate(schema, options.handlers)
    } catch (e) {
      error = e
    }
    return { ran: [...ran], error }
  }

  it('runs every section in DefaultHandlerSort order and skips the keys without a handler', async () => {
    const { ran: order, error } = await apply(template())
    assert.equal(error, undefined)
    assert.deepEqual(
      order,
      Object.keys(handlers).sort(
        (a, b) => DefaultHandlerSort[a] - DefaultHandlerSort[b]
      )
    )
    assert.ok(!order.includes('Version') && !order.includes('Parameters'))
  })

  it('runs only the handlers asked for, still in sort order', async () => {
    const { ran: order } = await apply(template(), {
      handlers: ['Navigation', 'SiteFields', 'Lists']
    })
    assert.deepEqual(order, ['SiteFields', 'Lists', 'Navigation'])
  })

  it('runs a handler a custom handlerSort leaves out last', async () => {
    const sort: any = { ...DefaultHandlerSort }
    delete sort.Security
    const { ran: order } = await apply(template(), { sort })
    assert.equal(order[order.length - 1], 'Security')
  })

  it('stops at a handler that throws and wraps it in a ProvisioningError naming the handler', async () => {
    const { ran: order, error } = await apply(template(), { fail: 'Lists' })
    assert.ok(
      error instanceof ProvisioningError || error.name === 'ProvisioningError'
    )
    assert.equal(error.handler, 'Lists')
    assert.match(error.message, /Lists failed/)
    assert.equal(order[order.length - 1], 'Lists')
    assert.ok(!order.includes('Navigation'))
  })
})
