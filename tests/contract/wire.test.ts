import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { Lists } from '../../src/handlers/lists'
import { Navigation } from '../../src/handlers/navigation'
import { SiteFields } from '../../src/handlers/sitefields'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { TokenHelper } from '../../src/util/tokenhelper'
import { createWireWeb, IWireRequest, writes } from '../helpers/transport'

/**
 * The contract layer: what the handlers send, through the real PnPjs queryables, with a transport
 * that answers in the process (tests/helpers/transport.ts). A change to a URL, a method or a body
 * fails here, where the unit tests' fakes would not notice.
 */

const line = (request: IWireRequest) => `${request.method} ${request.path}`

describe('contract: SiteFields', () => {
  const ID = 'a1a1a1a1-0000-4000-8000-000000000001'
  const stored = `<Field Type="Text" DisplayName="Kode" Name="GtCode" StaticName="GtCode" ID="{${ID}}" SourceID="{web}" Version="3" />`

  /**
   * A web that has the column GtCode as `stored`.
   */
  const web = () =>
    createWireWeb((request) => {
      if (request.method === 'GET' && request.path === '_api/web/fields')
        return [
          {
            Id: ID,
            InternalName: 'GtCode',
            TypeAsString: 'Text',
            SchemaXml: stored
          }
        ]
      if (
        request.method === 'GET' &&
        request.path === '_api/web/availablefields'
      )
        return [{ Id: ID, InternalName: 'GtCode' }]
      if (request.path.endsWith('/createfieldasxml')) return { Id: 'new' }
      return undefined
    })

  it('reads the columns once, with the properties it compares, and no $top', async () => {
    const { web: w, requests } = web()
    await new SiteFields({}).ProvisionObjects(w, [], new ProvisioningContext())
    assert.deepEqual(
      requests.map((r) => [line(r), r.query.$select, r.query.$top]),
      [
        [
          'GET _api/web/fields',
          'Id,InternalName,TypeAsString,SchemaXml',
          undefined
        ],
        ['GET _api/web/availablefields', 'Id,InternalName', undefined]
      ]
    )
  })

  it('creates a missing column with createfieldasxml and the XML as it stands', async () => {
    const { web: w, requests } = web()
    const xml =
      '<Field Type="Note" DisplayName="Ny" Name="GtNew" StaticName="GtNew" ID="{b2b2b2b2-0000-4000-8000-000000000002}" />'
    await new SiteFields({}).ProvisionObjects(
      w,
      [xml],
      new ProvisioningContext()
    )
    assert.deepEqual(
      writes(requests).map((r) => [line(r), r.body]),
      [
        [
          'POST _api/web/fields/createfieldasxml',
          { parameters: { SchemaXml: xml } }
        ]
      ]
    )
  })

  it('merges a changed column onto its stored schema and sends it with MERGE, Version dropped', async () => {
    const { web: w, requests } = web()
    await new SiteFields({}).ProvisionObjects(
      w,
      [
        `<Field Type="Text" DisplayName="Kode ny" Name="GtCode" StaticName="GtCode" ID="{${ID}}" />`
      ],
      new ProvisioningContext()
    )
    const [merge] = writes(requests)
    assert.equal(line(merge), `MERGE _api/web/fields('${ID}')`)
    assert.deepEqual(Object.keys(merge.body), ['SchemaXml'])
    assert.match(merge.body.SchemaXml, /DisplayName="Kode ny"/)
    assert.match(merge.body.SchemaXml, /SourceID="\{web\}"/)
    assert.doesNotMatch(merge.body.SchemaXml, /Version=/)
  })

  it('sends nothing for an unchanged column', async () => {
    const { web: w, requests } = web()
    await new SiteFields({}).ProvisionObjects(
      w,
      [stored],
      new ProvisioningContext()
    )
    assert.deepEqual(writes(requests), [])
  })

  it("sets a taxonomy column's empty term store with a property MERGE, not in its SchemaXml", async () => {
    const TAX = 'c3c3c3c3-0000-4000-8000-000000000003'
    const store = '33333333-3333-4333-8333-333333333333'
    const value = (text: string) =>
      `<Property><Name>SspId</Name><Value xmlns:q1="http://www.w3.org/2001/XMLSchema" p4:type="q1:string" xmlns:p4="http://www.w3.org/2001/XMLSchema-instance">${text}</Value></Property>`
    const field = (sspId: string, extra = '') =>
      `<Field Type="TaxonomyFieldTypeMulti" DisplayName="Fag" Name="GtFag" StaticName="GtFag" ID="{${TAX}}" Mult="TRUE" ShowField="Term1044"${extra}><Customization><ArrayOfProperty>${value(
        sspId
      )}</ArrayOfProperty></Customization></Field>`
    const { web: w, requests } = createWireWeb((request) => {
      if (request.method === 'GET' && request.path === '_api/web/fields')
        return [
          {
            Id: TAX,
            InternalName: 'GtFag',
            TypeAsString: 'TaxonomyFieldTypeMulti',
            SchemaXml: field(
              '00000000-0000-0000-0000-000000000000',
              ' List="{11111111-1111-4111-8111-111111111111}" WebId="web" Version="1"'
            )
          }
        ]
      if (request.method === 'GET') return []
      return undefined
    })
    const context = new ProvisioningContext()
    context.termStoreId = store
    await new SiteFields({}).ProvisionObjects(
      w,
      [field('{sitecollectiontermstoreid}')],
      context
    )
    assert.deepEqual(
      writes(requests).map((r) => [line(r), r.body]),
      [[`MERGE _api/web/fields('${TAX}')`, { SspId: store }]]
    )
  })
})

describe('contract: Lists DataRows', () => {
  /**
   * Runs a list's DataRows against a web whose list has the columns Title and GtSortOrder and the
   * given items for every key lookup.
   *
   * @param rows - The DataRows
   * @param existing - The items a key lookup finds
   */
  async function run(rows: any, existing: any[] = []) {
    const { web, requests } = createWireWeb((request) => {
      if (request.method === 'GET' && request.path.endsWith('/fields'))
        return [
          { Id: '1', InternalName: 'Title', TypeAsString: 'Text' },
          { Id: '2', InternalName: 'GtSortOrder', TypeAsString: 'Number' }
        ]
      if (request.method === 'GET' && request.path.endsWith('/items'))
        return existing
      if (request.method === 'POST' && request.path.endsWith('/items'))
        return { Id: 7 }
      return undefined
    })
    const handler: any = new Lists({})
    handler.context = new ProvisioningContext()
    handler.tokenHelper = new TokenHelper(handler.context, {})
    await handler.processListDataRows(web, { Title: 'Fase', DataRows: rows })
    return requests
  }

  it('matches a row by its key with an OData-escaped filter and $top=2, then adds it', async () => {
    const requests = await run({
      KeyColumn: 'Title',
      Rows: [{ Title: "O'Brien", GtSortOrder: 10 }]
    })
    const lookup = requests.find(
      (r) => r.method === 'GET' && r.path.endsWith('/items')
    )
    assert.deepEqual(
      [
        line(lookup),
        lookup.query.$filter,
        lookup.query.$select,
        lookup.query.$top
      ],
      [
        "GET _api/web/lists/getByTitle('Fase')/items",
        "Title eq 'O''Brien'",
        'Id',
        '2'
      ]
    )
    assert.deepEqual(
      writes(requests).map((r) => [line(r), r.body]),
      [
        [
          "POST _api/web/lists/getByTitle('Fase')/items",
          { Title: "O'Brien", GtSortOrder: 10 }
        ]
      ]
    )
  })

  it('quotes a Number key as SharePoint accepts it', async () => {
    const requests = await run({
      KeyColumn: 'GtSortOrder',
      Rows: [{ Title: 'Sjekkpunkt', GtSortOrder: 10 }]
    })
    const lookup = requests.find(
      (r) => r.method === 'GET' && r.path.endsWith('/items')
    )
    assert.equal(lookup.query.$filter, "GtSortOrder eq '10'")
  })

  it('writes nothing for a matched row with Skip, and MERGEs it with Overwrite', async () => {
    const rows = (UpdateBehavior: string) => ({
      KeyColumn: 'Title',
      UpdateBehavior,
      Rows: [{ Title: 'Finnes', GtSortOrder: 20 }]
    })
    assert.deepEqual(writes(await run(rows('Skip'), [{ Id: 4 }])), [])
    assert.deepEqual(
      writes(await run(rows('Overwrite'), [{ Id: 4 }])).map((r) => [
        line(r),
        r.body
      ]),
      [
        [
          "MERGE _api/web/lists/getByTitle('Fase')/items(4)",
          { Title: 'Finnes', GtSortOrder: 20 }
        ]
      ]
    )
  })
})

describe('contract: Navigation', () => {
  /**
   * A Quick Launch the requests act on: adds append, DELETE removes, MoveAfter reorders and a
   * MERGE sets the URL, so the handler's read-back sees what it did.
   *
   * @param nodes - The stored top-level nodes
   */
  function quickLaunch(
    nodes: Array<{ Id: number; Title: string; Url: string }>
  ) {
    let next = 100
    const tree = nodes.map((node) => ({ ...node, Children: [] }))
    return createWireWeb((request) => {
      const byId = /quicklaunch\((\d+)\)$/.exec(request.path)
      if (request.method === 'GET' && request.path === '_api/web')
        return {
          Url: 'https://contoso.sharepoint.com/sites/demo',
          ServerRelativeUrl: '/sites/demo'
        }
      // A copy: the handler must not see later changes through what it read.
      if (request.method === 'GET' && request.path.endsWith('/quicklaunch'))
        return JSON.parse(JSON.stringify(tree))
      if (request.method === 'POST' && request.path.endsWith('/quicklaunch')) {
        const node = {
          Id: next++,
          Title: request.body.Title,
          Url: request.body.Url,
          Children: []
        }
        tree.push(node)
        return node
      }
      if (request.method === 'DELETE' && byId)
        tree.splice(
          tree.findIndex((node) => node.Id === Number(byId[1])),
          1
        )
      if (request.method === 'MERGE' && byId)
        Object.assign(
          tree.find((node) => node.Id === Number(byId[1])),
          request.body
        )
      if (request.path.endsWith('/quicklaunch/MoveAfter')) {
        const { nodeId, previousNodeId } = request.body
        const [node] = tree.splice(
          tree.findIndex((n) => n.Id === nodeId),
          1
        )
        tree.splice(tree.findIndex((n) => n.Id === previousNodeId) + 1, 0, node)
      }
      return undefined
    })
  }

  it('reconciles with adds, URL MERGEs, deletes and moves, in that order', async () => {
    const { web, requests } = quickLaunch([
      { Id: 1, Title: 'Hjem', Url: '/sites/demo' },
      { Id: 2, Title: 'Gammel', Url: '/sites/demo/SitePages/Gammel.aspx' },
      { Id: 3, Title: 'Lenke', Url: 'https://www.contoso.com/a' }
    ])
    await new Navigation({}).ProvisionObjects(web, {
      QuickLaunch: [
        { Title: 'Ny', Url: 'SitePages/Ny.aspx' },
        { Title: 'Hjem', Url: '' },
        {
          Title: 'Lenke',
          Url: 'https://www.contoso.com/b',
          IgnoreExisting: true
        }
      ]
    } as any)
    assert.deepEqual(
      writes(requests).map((r) => [line(r), r.body]),
      [
        [
          'POST _api/web/navigation/quicklaunch',
          { IsVisible: true, Title: 'Ny', Url: 'SitePages/Ny.aspx' }
        ],
        [
          'MERGE _api/web/navigation/quicklaunch(3)',
          { Url: 'https://www.contoso.com/b' }
        ],
        ['DELETE _api/web/navigation/quicklaunch(2)', undefined],
        // Each node after the first is moved after its predecessor once something changed.
        [
          'POST _api/web/navigation/quicklaunch/MoveAfter',
          { nodeId: 1, previousNodeId: 100 }
        ],
        [
          'POST _api/web/navigation/quicklaunch/MoveAfter',
          { nodeId: 3, previousNodeId: 1 }
        ]
      ]
    )
  })

  it('reads the collection with its children in one request, and writes nothing when it is unchanged', async () => {
    const { web, requests } = quickLaunch([
      { Id: 1, Title: 'Hjem', Url: '/sites/demo' }
    ])
    await new Navigation({}).ProvisionObjects(web, {
      QuickLaunch: [{ Title: 'Hjem', Url: '' }]
    } as any)
    assert.deepEqual(writes(requests), [])
    const [read] = requests
    assert.deepEqual(
      [line(read), read.query.$expand],
      ['GET _api/web/navigation/quicklaunch', 'Children']
    )
  })
})
