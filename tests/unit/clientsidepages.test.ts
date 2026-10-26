import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { ClientSidePages } from '../../src/handlers/clientsidepages'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { TokenHelper } from '../../src/util/tokenhelper'

const QUICK_LINKS = 'c70391ea-0b10-4ee9-b2b4-006d3fcad0cd'

/** A web part definition as `getClientsideWebParts()` returns it. */
const definition = {
  Id: `{${QUICK_LINKS}}`,
  Name: 'QuickLinks',
  Manifest: JSON.stringify({
    preconfiguredEntries: [
      {
        title: { default: 'Quick links' },
        description: { default: '' },
        properties: {}
      }
    ]
  })
}

/**
 * The handler, set up as ProvisionObjects sets it up, for a web with the given title.
 *
 * @param title - The web's title (`{sitename}`)
 * @param parameters - Values for `{parameter:...}`
 */
function handlerFor(title: string, parameters: Record<string, string> = {}) {
  const context = new ProvisioningContext()
  context.web = { Id: 'web-guid', Title: title }
  const handler: any = new ClientSidePages({ parameters })
  handler.tokenHelper = new TokenHelper(context, { parameters })
  return handler
}

describe('ClientSidePages: tokens in web part properties', () => {
  it('keeps a web part whose token value holds a quote or a backslash', () => {
    // The tokens used to be replaced in the properties' JSON text: a web titled
    // 'Rehab "Gamle skole"' made JSON.parse throw, and the page lost the part.
    const handler = handlerFor('Rehab "Gamle skole"', { path: 'C:\\Prosjekt' })
    const { part } = handler._getClientSideWebPart([definition], {
      Id: QUICK_LINKS,
      Properties: {
        title: '{sitename}',
        items: [{ title: 'Mappe', url: '{parameter:path}' }]
      },
      ServerProcessedContent: {
        htmlStrings: {},
        searchablePlainTexts: { title: '{sitename}' },
        imageSources: {},
        links: {}
      }
    })

    assert.deepEqual(part.data.webPartData.properties, {
      title: 'Rehab "Gamle skole"',
      items: [{ title: 'Mappe', url: 'C:\\Prosjekt' }]
    })
    assert.deepEqual(
      part.data.webPartData.serverProcessedContent.searchablePlainTexts,
      { title: 'Rehab "Gamle skole"' }
    )
  })
})
