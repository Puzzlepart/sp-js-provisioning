import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { LogLevel } from '@pnp/logging'
import { ClientSidePages } from '../../src/handlers/clientsidepages'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { IClientSidePage } from '../../src/schema'
import { TokenHelper } from '../../src/util/tokenhelper'
import { captureLogs, levelOf } from '../helpers/logs'
import { createWireWeb, IWireRequest, writes } from '../helpers/transport'

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

const WEB = 'https://contoso.sharepoint.com/sites/demo'
/** The page's file, as the handler reads it. */
const FILE =
  "_api/web/getFileByServerRelativePath(decodedUrl='/sites/demo/SitePages/Status.aspx')"
/** `SitePages`, as `EnsureClientRenderedSitePagesLibrary` names it. */
const LIBRARY = "_api/Web/Lists(guid'pages-list')"
const PAGE = '_api/sitepages/pages(4)'

/** What a draft left by a stopped setup holds: a text control and the page settings. */
const DRAFT_CANVAS = [
  {
    controlType: 4,
    id: 'draft-text',
    position: {
      zoneIndex: 1,
      sectionIndex: 1,
      controlIndex: 1,
      sectionFactor: 12,
      layoutIndex: 1
    },
    innerHTML: '<p>Draft</p>',
    editorType: 'CKEditor'
  },
  {
    controlType: 0,
    pageSettingsSlice: { isDefaultDescription: true, isDefaultThumbnail: true }
  }
]

/**
 * A web whose `SitePages` has `Status.aspx` (item 4), answering the requests PnPjs sends to load,
 * save and publish a page.
 *
 * @param majorVersion - The file's `MajorVersion`: 0 for a page that has never been published
 */
function createPagesSite(majorVersion: number) {
  const item = {
    Id: 4,
    CommentsDisabled: false,
    'odata.metadata': `${WEB}/_api/$metadata#SP.ListData.SitePagesItems/@Element`,
    'odata.editLink': "Web/Lists(guid'pages-list')/Items(4)"
  }
  return createWireWeb((request: IWireRequest) => {
    switch (request.path) {
      case '_api/web/GetClientSideWebParts':
        return []
      case '_api/web':
        return { ServerRelativeUrl: '/sites/demo' }
      case `${FILE}/listItemAllFields`:
      case `${LIBRARY}/items(4)`:
        return item
      case FILE:
        return { MajorVersion: majorVersion }
      case '_api/lists/EnsureClientRenderedSitePagesLibrary':
        return { Id: 'pages-list', 'odata.id': `${WEB}/${LIBRARY}` }
      case PAGE:
        return {
          Id: 4,
          // A stopped setup leaves the page titled with its file name.
          Title: 'Status',
          PageLayoutType: 'Article',
          CanvasContent1: JSON.stringify(DRAFT_CANVAS),
          LayoutWebpartsContent: '[]',
          IsPageCheckedOutToCurrentUser: true
        }
      default:
        return undefined
    }
  }, WEB)
}

/** The template's page: `Overwrite: false`, a text control, comments off. */
const STATUS: IClientSidePage = {
  Name: 'Status.aspx',
  Title: 'Prosjektstatus',
  PageLayoutType: 'Article',
  CommentsDisabled: true,
  Overwrite: false,
  Sections: [
    {
      Columns: [
        {
          Factor: 12,
          Controls: [{ Id: 'Text', Text: '<p>Status</p>', Properties: {} }]
        }
      ]
    }
  ]
}

/**
 * Runs the handler on the web with the template's page and returns what it logged.
 *
 * @param web - The fake web
 */
async function provisionStatus(web: any) {
  const { entries } = await captureLogs(() =>
    new ClientSidePages({}).ProvisionObjects(
      web,
      [JSON.parse(JSON.stringify(STATUS))],
      new ProvisioningContext()
    )
  )
  return entries
}

describe('ClientSidePages: a page the web has, with Overwrite false', () => {
  it('writes the template to a page that has never been published, in place, and publishes it', async () => {
    // A setup that stopped part way (throttled, PP365 stress test row 13) leaves its page a
    // draft: Navigation is refused a node to it, so a re-run must publish it.
    const { web, requests } = createPagesSite(0)

    const entries = await provisionStatus(web)

    assert.deepEqual(
      writes(requests).map((request) => `${request.method} ${request.path}`),
      [
        // PnPjs's read of the pages library (load, and the comments' item).
        'POST _api/lists/EnsureClientRenderedSitePagesLibrary',
        `POST ${PAGE}/savepage`,
        `POST ${PAGE}/publish`,
        'POST _api/lists/EnsureClientRenderedSitePagesLibrary',
        'POST _api/lists/EnsureClientRenderedSitePagesLibrary',
        `MERGE ${LIBRARY}/Items(4)/SetCommentsDisabled(true)`
      ],
      'neither a new page nor a deleted one'
    )
    const saved = requests.find((request) => request.path.endsWith('/savepage'))
      .body
    assert.equal(saved.Title, 'Prosjektstatus')
    // The template's content replaces the draft's.
    assert.deepEqual(
      JSON.parse(saved.CanvasContent1)
        .filter((control: any) => control.controlType === 4)
        .map((control: any) => control.innerHTML),
      ['<p>Status</p>']
    )
    assert.equal(
      levelOf(
        entries,
        /Client side page Status\.aspx exists but has never been published/
      ),
      LogLevel.Info
    )
    assert.deepEqual(
      entries.filter((entry) => entry.level >= LogLevel.Warning),
      []
    )
  })

  it('leaves a published page as it is', async () => {
    const { web, requests } = createPagesSite(1)

    const entries = await provisionStatus(web)

    assert.deepEqual(
      writes(requests).map((request) => `${request.method} ${request.path}`),
      ['POST _api/lists/EnsureClientRenderedSitePagesLibrary']
    )
    assert.equal(
      levelOf(
        entries,
        /Client side page Status\.aspx already exists and overwrite option is disabled\. Skipping/
      ),
      LogLevel.Info
    )
  })
})
