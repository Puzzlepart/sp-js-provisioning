import { LogLevel } from '@pnp/logging'
import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { Navigation } from '../../src/handlers/navigation'
import { INavigation } from '../../src/schema'
import { captureLogs, levelOf } from '../helpers/logs'

// `replaceUrlTokens` resolves `{hosturl}` from `window` on every call, so the
// handler needs one even when a node url contains no tokens.
const globals: any = globalThis
globals.window = {
  location: { protocol: 'https:', host: 'contoso.sharepoint.com', port: '' }
}

const WEB = '/sites/demo'
const ORIGIN = 'https://contoso.sharepoint.com'

/** A node as the fake stores it, children in display order. */
interface IFakeNode {
  Id: number
  Title: string
  Url: string
  AudienceIds?: string[]
  children: IFakeNode[]
}

/** A node to seed the fake with. */
interface ISeed {
  Title: string
  Url: string
  AudienceIds?: string[]
  children?: ISeed[]
}

/** A node as the assertions compare it: no ids. */
interface IShape {
  Title: string
  Url: string
  children: IShape[]
}

interface IFakeOptions {
  /** URLs SharePoint refuses, on add and on a `Url` update (a missing file or folder) */
  refuse?: (url: string) => boolean
  /** Node ids whose delete fails */
  failDelete?: (id: number) => boolean
  /** Node ids whose delete answers success and leaves the node (as for a missing node, X1) */
  ignoreDelete?: (id: number) => boolean
  /** Node ids whose move fails */
  failMove?: (id: number) => boolean
  /** Node ids whose move answers success and leaves the node where it is (SharePoint checks nothing, N4) */
  ignoreMove?: (id: number) => boolean
  /** Collections (by name) whose read fails */
  failRead?: (name: string) => boolean
  /** Expand only this many levels, whatever the read asks for */
  expandDepth?: number
}

interface IFakeWeb {
  /** The fake `IWeb`, passed to the handler */
  web: any
  /** Every request, in order: `GET ...`, `add ...`, `update ...`, `delete ...`, `moveAfter ...` */
  calls: string[]
  /** The requests that write */
  writes: () => string[]
  seed: (collection: 'quicklaunch' | 'topnavigationbar', nodes: ISeed[]) => void
  /** A collection without ids */
  tree: (collection: 'quicklaunch' | 'topnavigationbar') => IShape[]
  /** A collection's ids, depth first */
  ids: (collection: 'quicklaunch' | 'topnavigationbar') => number[]
  /** A stored node by title, any collection, any level */
  byTitle: (title: string) => IFakeNode
}

/**
 * The error PnPjs 4 rejects with for a failed request: the response body after `::> `.
 *
 * @param code - SharePoint's error code
 * @param message - SharePoint's message
 */
function httpError(code: string, message: string): Error {
  const error: any = new Error(
    `Error making HttpClient request in queryable [500] Internal Server Error ::> ${JSON.stringify(
      { 'odata.error': { code, message: { lang: 'nb-NO', value: message } } }
    )}`
  )
  error.status = 500
  return error
}

/** A node URL as SharePoint stores it: web-relative resolved against the web (probe N1). */
const stored = (url: string) =>
  /^(\/|[a-z]+:)/i.test(url)
    ? url
    : url === ''
    ? WEB
    : url === '#'
    ? `${WEB}#`
    : `${WEB}/${url}`

/**
 * Hand-rolled stand-in for a web's navigation in PnPjs v4: `web.navigation.quicklaunch` and
 * `topNavigationBar`, invokable with `select`/`expand` (`Children` returned to the depth the read
 * expanded, as SharePoint does), `add` resolving with an `INavNodeInfo`, `moveAfter`, and
 * `getById` reaching any node of the web (probe N6) with `children`, `update` and `delete`. One
 * store for both collections and one call log, so the order of the requests across them shows.
 *
 * @param options - Failures to inject
 */
function createFakeWeb(options: IFakeOptions = {}): IFakeWeb {
  const collections: Record<string, IFakeNode[]> = {
    quicklaunch: [],
    topnavigationbar: []
  }
  const calls: string[] = []
  let nextId = 1

  /** The node with this id and the list it is in. */
  function locate(
    id: number,
    lists: IFakeNode[][] = Object.values(collections)
  ): { node: IFakeNode; siblings: IFakeNode[] } | undefined {
    for (const siblings of lists)
      for (const node of siblings) {
        if (node.Id === id) return { node, siblings }
        const found = locate(id, [node.children])
        if (found) return found
      }
    return undefined
  }

  function shape(node: IFakeNode, depth: number): any {
    const info: any = {
      Id: node.Id,
      Title: node.Title,
      Url: node.Url,
      AudienceIds: node.AudienceIds || null
    }
    if (depth > 0)
      info.Children = node.children.map((child) => shape(child, depth - 1))
    return info
  }

  function nodeOf(root: string, id: number): any {
    const path = `${root}(${id})`
    const read: any = async () => {
      calls.push(`GET ${path}`)
      const found = locate(id)
      return found ? shape(found.node, 0) : {}
    }
    read.select = () => read
    read.children = nodesOf(root, `${path}/children`, () => {
      const found = locate(id)
      if (!found) throw httpError('-1', `Finner ikke ressurs ${path}`)
      return found.node.children
    })
    read.update = async (properties: any) => {
      calls.push(`update ${id} ${JSON.stringify(properties)}`)
      if (options.refuse && options.refuse(properties.Url))
        throw httpError(
          '-2130247147, Microsoft.SharePoint.SPException',
          `Kan ikke åpne ${properties.Url}: Filen eller mappen finnes ikke.`
        )
      const found = locate(id)
      if (found) found.node.Url = stored(properties.Url)
      return { data: {}, node: read }
    }
    read.delete = async () => {
      calls.push(`delete ${id}`)
      if (options.failDelete && options.failDelete(id))
        throw httpError('-1, System.Exception', 'Too many requests')
      if (options.ignoreDelete && options.ignoreDelete(id)) return
      const found = locate(id)
      if (found) found.siblings.splice(found.siblings.indexOf(found.node), 1)
    }
    return read
  }

  function nodesOf(root: string, name: string, list: () => IFakeNode[]): any {
    const query = (expand: string[]): any => {
      const read: any = async () => {
        calls.push(
          `GET ${name}${
            expand.length > 0 ? ` $expand=${expand.join(',')}` : ''
          }`
        )
        if (options.failRead && options.failRead(name))
          throw httpError('-1, System.Exception', 'Too many requests')
        const depth = Math.min(
          Math.max(0, ...expand.map((p) => p.split('/').length)),
          options.expandDepth === undefined ? Infinity : options.expandDepth
        )
        return list().map((node) => shape(node, depth))
      }
      read.select = () => read
      read.expand = (...paths: string[]) => query([...expand, ...paths])
      return read
    }
    const nodes = query([])
    nodes.add = async (title: string, url: string) => {
      calls.push(`add ${name} '${title}' ${url}`)
      if (options.refuse && options.refuse(url))
        throw httpError(
          '-2130247147, Microsoft.SharePoint.SPException',
          `Kan ikke åpne ${stored(url)}: Filen eller mappen finnes ikke.`
        )
      const node = {
        Id: nextId++,
        Title: title,
        Url: stored(url),
        children: []
      }
      list().push(node)
      return { Id: node.Id, Title: title, Url: node.Url }
    }
    nodes.getById = (id: number) => nodeOf(root, id)
    nodes.moveAfter = async (nodeId: number, previousNodeId: number) => {
      calls.push(`moveAfter ${name} ${nodeId} ${previousNodeId}`)
      if (options.ignoreMove && options.ignoreMove(nodeId)) return
      const moving = locate(nodeId)
      if (!moving || (options.failMove && options.failMove(nodeId)))
        throw httpError(
          '-1, Microsoft.SharePoint.Client.UnknownError',
          'Ukjent feil'
        )
      moving.siblings.splice(moving.siblings.indexOf(moving.node), 1)
      const siblings = list()
      const previous = siblings.find((node) => node.Id === previousNodeId)
      siblings.splice(
        previous ? siblings.indexOf(previous) + 1 : siblings.length,
        0,
        moving.node
      )
    }
    return nodes
  }

  const plant = (seeds: ISeed[]): IFakeNode[] =>
    seeds.map((seed) => ({
      Id: nextId++,
      Title: seed.Title,
      Url: seed.Url,
      AudienceIds: seed.AudienceIds,
      children: plant(seed.children || [])
    }))
  const strip = (nodes: IFakeNode[]): IShape[] =>
    nodes.map((node) => ({
      Title: node.Title,
      Url: node.Url,
      children: strip(node.children)
    }))
  const idsOf = (nodes: IFakeNode[]): number[] =>
    nodes.reduce(
      (all: number[], node) => [...all, node.Id, ...idsOf(node.children)],
      []
    )
  const find = (nodes: IFakeNode[], title: string): IFakeNode | undefined =>
    nodes.reduce(
      (found: IFakeNode | undefined, node) =>
        found || (node.Title === title ? node : find(node.children, title)),
      undefined
    )

  return {
    web: {
      navigation: {
        get quicklaunch() {
          return nodesOf(
            'quicklaunch',
            'quicklaunch',
            () => collections.quicklaunch
          )
        },
        get topNavigationBar() {
          return nodesOf(
            'topnavigationbar',
            'topnavigationbar',
            () => collections.topnavigationbar
          )
        }
      },
      select: () => async () => {
        calls.push('GET web')
        return { Url: `${ORIGIN}${WEB}`, ServerRelativeUrl: WEB }
      }
    },
    calls,
    writes: () => calls.filter((call) => !call.startsWith('GET')),
    seed: (collection, nodes) => {
      collections[collection] = plant(nodes)
    },
    tree: (collection) => strip(collections[collection]),
    ids: (collection) => idsOf(collections[collection]),
    byTitle: (title) =>
      find([...collections.quicklaunch, ...collections.topnavigationbar], title)
  }
}

/**
 * Applies a navigation template through the handler.
 *
 * @param fake - The fake web
 * @param navigation - The template's `Navigation`
 */
function provision(fake: IFakeWeb, navigation: INavigation): Promise<void> {
  return new Navigation({}).ProvisionObjects(fake.web, navigation)
}

/** A template's nodes as the fake stores them after a clean run. */
function expected(nodes: INavigation['QuickLaunch']): IShape[] {
  return (nodes || []).map((node) => ({
    Title: node.Title,
    Url: stored(node.Url),
    children: expected(node.Children)
  }))
}

describe('Navigation: v4 node creation result', () => {
  it('adds child nodes under the node returned by add', async () => {
    const fake = createFakeWeb()

    await provision(fake, {
      QuickLaunch: [
        {
          Title: 'Prosjekter',
          Url: '/sites/demo/Lists/Prosjekter',
          Children: [{ Title: 'Aktive', Url: '/sites/demo/aktive' }]
        }
      ]
    })

    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'Prosjekter' /sites/demo/Lists/Prosjekter",
      "add quicklaunch(1)/children 'Aktive' /sites/demo/aktive"
    ])
    assert.deepEqual(fake.tree('quicklaunch'), [
      {
        Title: 'Prosjekter',
        Url: '/sites/demo/Lists/Prosjekter',
        children: [{ Title: 'Aktive', Url: '/sites/demo/aktive', children: [] }]
      }
    ])
  })

  it('replaces the existing nodes rather than appending to them', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [{ Title: 'Gammel', Url: '/sites/demo/gammel' }])

    await provision(fake, {
      QuickLaunch: [{ Title: 'Ny', Url: '/sites/demo/ny' }]
    })

    // The add comes first, the delete last (nothing irreversible before the last add).
    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'Ny' /sites/demo/ny",
      'delete 1'
    ])
    assert.deepEqual(fake.tree('quicklaunch'), [
      { Title: 'Ny', Url: '/sites/demo/ny', children: [] }
    ])
  })
})

describe('Navigation: reconcile in place (replace mode)', () => {
  const template = (): INavigation => ({
    QuickLaunch: [
      { Title: 'Hjem', Url: '' },
      {
        Title: 'Prosjekter',
        Url: 'Lists/Prosjekter',
        Children: [
          { Title: 'Aktive', Url: 'SitePages/Aktive.aspx' },
          {
            Title: 'Arkiv',
            Url: 'https://www.contoso.com/arkiv',
            IgnoreExisting: true,
            Children: [{ Title: 'Eldre', Url: 'https://www.contoso.com/eldre' }]
          }
        ]
      },
      { Title: 'Notatblokk', Url: '#', IgnoreExisting: true }
    ],
    TopNavigationBar: [{ Title: 'Topp', Url: 'https://www.fabrikam.com/' }]
  })

  it('an unchanged second run reads each collection once and writes nothing', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      { Title: 'Samtaler', Url: '/sites/demo/_layouts/15/groupstatus.aspx' }
    ])
    await provision(fake, template())
    const ids = [...fake.ids('quicklaunch'), ...fake.ids('topnavigationbar')]
    const tree = fake.tree('quicklaunch')
    assert.deepEqual(tree, expected(template().QuickLaunch))
    // Audiences an owner set on a kept node.
    fake.byTitle('Prosjekter').AudienceIds = [
      '8f0f7c0b-0000-4000-8000-000000000001'
    ]
    fake.calls.length = 0

    await provision(fake, template())

    assert.deepEqual(fake.writes(), [])
    assert.deepEqual(fake.calls, [
      'GET quicklaunch $expand=Children,Children/Children,Children/Children/Children',
      // The IgnoreExisting nodes compare their URLs with the template's, resolved against the web.
      'GET web',
      'GET topnavigationbar $expand=Children'
    ])
    assert.deepEqual(
      [...fake.ids('quicklaunch'), ...fake.ids('topnavigationbar')],
      ids
    )
    assert.deepEqual(fake.tree('quicklaunch'), tree)
    assert.deepEqual(fake.byTitle('Prosjekter').AudienceIds, [
      '8f0f7c0b-0000-4000-8000-000000000001'
    ])
  })

  it('does not change the template object', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      { Title: 'Hjem', Url: '/sites/demo/SitePages/Home.aspx' },
      {
        Title: 'Prosjekter',
        Url: '/sites/demo/owner',
        children: [{ Title: 'Aktive', Url: '/sites/demo/owner/aktive' }]
      }
    ])
    const navigation = template()
    const copy = JSON.parse(JSON.stringify(navigation))

    await provision(fake, navigation)

    // The kept nodes' stored URLs were not written into it (it used to set `node.Url`).
    assert.deepEqual(navigation, copy)
  })

  it('moves the kept nodes into the template order and writes nothing else', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      { Title: 'A', Url: '/sites/demo/a' },
      { Title: 'B', Url: '/sites/demo/b' },
      { Title: 'C', Url: '/sites/demo/c' }
    ])

    await provision(fake, {
      QuickLaunch: [
        { Title: 'C', Url: 'c' },
        { Title: 'A', Url: 'a' },
        { Title: 'B', Url: 'b' }
      ]
    })

    // There is no move to the first position: each node goes after the one before it in the
    // template, and C ends up first once A and B are after it.
    assert.deepEqual(fake.writes(), [
      'moveAfter quicklaunch 1 3',
      'moveAfter quicklaunch 2 1'
    ])
    assert.deepEqual(
      fake.tree('quicklaunch').map((node) => node.Title),
      ['C', 'A', 'B']
    )
    assert.deepEqual(fake.ids('quicklaunch'), [3, 1, 2])
    // The order is read back after the moves.
    assert.equal(fake.calls[fake.calls.length - 1], 'GET quicklaunch')
  })

  it('moves a kept node after a new first node, and deletes before it moves', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      { Title: 'Owner', Url: '/sites/demo/owner' },
      { Title: 'B', Url: '/sites/demo/b' }
    ])

    await provision(fake, {
      QuickLaunch: [
        { Title: 'A', Url: '/sites/demo/a' },
        { Title: 'B', Url: '/sites/demo/b' }
      ]
    })

    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'A' /sites/demo/a",
      'delete 1',
      'moveAfter quicklaunch 2 3'
    ])
    assert.deepEqual(
      fake.tree('quicklaunch').map((node) => node.Title),
      ['A', 'B']
    )
  })

  it('reconciles the children of a kept node: adds, deletes and moves there, ids kept', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      {
        Title: 'Prosjekter',
        Url: '/sites/demo/Lists/Prosjekter',
        children: [
          { Title: 'Owner', Url: '/sites/demo/owner' },
          { Title: 'Arkiv', Url: '/sites/demo/arkiv' },
          { Title: 'Aktive', Url: '/sites/demo/aktive' }
        ]
      }
    ])

    await provision(fake, {
      QuickLaunch: [
        {
          Title: 'Prosjekter',
          Url: 'Lists/Prosjekter',
          Children: [
            { Title: 'Aktive', Url: 'aktive' },
            { Title: 'Ny', Url: 'ny' },
            { Title: 'Arkiv', Url: 'arkiv' }
          ]
        }
      ]
    })

    assert.deepEqual(fake.writes(), [
      "add quicklaunch(1)/children 'Ny' ny",
      'delete 2',
      // Ny is already after Aktive; Arkiv goes after Ny.
      'moveAfter quicklaunch(1)/children 3 5'
    ])
    assert.deepEqual(fake.ids('quicklaunch'), [1, 4, 5, 3])
  })

  it('removes the stored children of a kept node whose template node has no Children', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      {
        Title: 'Prosjekter',
        Url: '/sites/demo/p',
        children: [
          { Title: 'c1', Url: '/sites/demo/c1' },
          { Title: 'c2', Url: '/sites/demo/c2' }
        ]
      }
    ])

    await provision(fake, {
      QuickLaunch: [{ Title: 'Prosjekter', Url: '/sites/demo/p' }]
    })

    assert.deepEqual(fake.writes(), ['delete 2', 'delete 3'])
  })

  it('a kept node keeps its stored URL at every level without IgnoreExisting', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      {
        Title: 'Hjem',
        Url: '/sites/demo/owner.aspx',
        children: [{ Title: 'Barn', Url: 'https://www.contoso.com/owner' }]
      }
    ])

    await provision(fake, {
      QuickLaunch: [
        {
          Title: 'Hjem',
          Url: '/sites/demo/template.aspx',
          Children: [{ Title: 'Barn', Url: 'https://www.contoso.com/template' }]
        }
      ]
    })

    assert.deepEqual(fake.writes(), [])
    assert.equal(fake.byTitle('Barn').Url, 'https://www.contoso.com/owner')
  })

  it('IgnoreExisting sets the template URL on a kept node only where it differs', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      { Title: 'Hjem', Url: '/sites/demo/SitePages/Home.aspx' },
      {
        Title: 'Lenker',
        Url: 'https://www.contoso.com/old',
        children: [{ Title: 'Barn', Url: '/sites/demo/Delte dokumenter' }]
      }
    ])

    await provision(fake, {
      QuickLaunch: [
        // The same URL once normalised: case, web-relative, %20, trailing slash.
        { Title: 'Hjem', Url: 'sitepages/home.aspx', IgnoreExisting: true },
        {
          Title: 'Lenker',
          Url: 'https://www.contoso.com/new',
          IgnoreExisting: true,
          Children: [
            {
              Title: 'Barn',
              Url:
                'https://contoso.sharepoint.com/sites/demo/Delte%20dokumenter/',
              IgnoreExisting: true
            }
          ]
        }
      ]
    })

    // Only `Url` is sent, so the node keeps its audiences.
    assert.deepEqual(fake.writes(), [
      'update 2 {"Url":"https://www.contoso.com/new"}'
    ])
  })

  it('a title stored twice: the first is kept with the template URL, the second deleted', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      { Title: 'Dup', Url: '/sites/demo/one' },
      { Title: 'Dup', Url: '/sites/demo/two' }
    ])

    await provision(fake, {
      QuickLaunch: [{ Title: 'Dup', Url: '/sites/demo/template' }]
    })

    assert.deepEqual(fake.writes(), [
      'update 1 {"Url":"/sites/demo/template"}',
      'delete 2'
    ])
    assert.deepEqual(fake.ids('quicklaunch'), [1])
  })

  it('a title the template declares twice, stored once, is not rewritten on the next run', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [{ Title: 'Dup', Url: '/sites/demo/owner' }])
    const template = (): INavigation => ({
      QuickLaunch: [
        { Title: 'Dup', Url: '/sites/demo/t0' },
        { Title: 'Dup', Url: '/sites/demo/t1' }
      ]
    })

    await provision(fake, template())
    fake.calls.length = 0
    await provision(fake, template())

    assert.deepEqual(fake.writes(), [])
    assert.deepEqual(
      fake.tree('quicklaunch').map((node: any) => node.Url),
      ['/sites/demo/owner', '/sites/demo/t1']
    )
  })

  it('matches titles exactly, case included', async () => {
    const fake = createFakeWeb()
    fake.seed('quicklaunch', [
      { Title: 'Site contents', Url: '/sites/demo/_layouts/15/viewlsts.aspx' }
    ])

    await provision(fake, {
      QuickLaunch: [
        { Title: 'Site Contents', Url: '_layouts/15/viewlsts.aspx' }
      ]
    })

    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'Site Contents' _layouts/15/viewlsts.aspx",
      'delete 1'
    ])
  })

  it('reads the children of a kept node on their own where SharePoint left them out', async () => {
    const fake = createFakeWeb({ expandDepth: 1 })
    fake.seed('quicklaunch', [
      {
        Title: 'P',
        Url: '/sites/demo/p',
        children: [
          {
            Title: 'C',
            Url: '/sites/demo/c',
            children: [{ Title: 'G', Url: '/sites/demo/g' }]
          }
        ]
      }
    ])

    await provision(fake, {
      QuickLaunch: [
        {
          Title: 'P',
          Url: '/sites/demo/p',
          Children: [
            {
              Title: 'C',
              Url: '/sites/demo/c',
              Children: [{ Title: 'G', Url: '/sites/demo/g' }]
            }
          ]
        }
      ]
    })

    // C came without `Children`: its children were read, and G kept rather than added again.
    assert.ok(
      fake.calls.includes('GET quicklaunch(2)/children $expand=Children')
    )
    assert.deepEqual(fake.writes(), [])
  })

  it('leaves a collection the template does not name alone', async () => {
    const fake = createFakeWeb()
    fake.seed('topnavigationbar', [{ Title: 'Topp', Url: '/sites/demo' }])

    await provision(fake, {
      QuickLaunch: [{ Title: 'A', Url: '/sites/demo/a' }]
    })

    assert.deepEqual(
      fake.calls.filter((call) => /topnavigationbar/.test(call)),
      []
    )
  })
})

describe('Navigation: a refused node changes nothing', () => {
  const refuseMissing = (url: string) => /Missing/.test(url)

  it('a refused add removes the nodes added before it, deletes nothing and names the node', async () => {
    const fake = createFakeWeb({ refuse: refuseMissing })
    fake.seed('quicklaunch', [
      { Title: 'Hjem', Url: '/sites/demo' },
      { Title: 'Mine', Url: '/sites/demo/mine' }
    ])
    const before = fake.tree('quicklaunch')

    await assert.rejects(
      provision(fake, {
        QuickLaunch: [
          { Title: 'Added', Url: 'SitePages' },
          { Title: 'Missing', Url: 'SitePages/Missing.aspx' },
          { Title: 'After', Url: 'SitePages/After.aspx' }
        ]
      }),
      {
        message:
          "Navigation node 'Missing' (SitePages/Missing.aspx) was refused by SharePoint: " +
          'Kan ikke åpne /sites/demo/SitePages/Missing.aspx: Filen eller mappen finnes ikke.'
      }
    )

    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'Added' SitePages",
      "add quicklaunch 'Missing' SitePages/Missing.aspx",
      'delete 3'
    ])
    assert.deepEqual(fake.tree('quicklaunch'), before)
    assert.deepEqual(fake.ids('quicklaunch'), [1, 2])
  })

  it('a refused URL update removes the nodes the run added and sets the URLs it changed back', async () => {
    const fake = createFakeWeb({ refuse: refuseMissing })
    fake.seed('quicklaunch', [
      { Title: 'Hjem', Url: '/sites/demo' },
      { Title: 'Dokumenter', Url: '/sites/demo/Delte dokumenter' }
    ])
    const before = fake.tree('quicklaunch')

    await assert.rejects(
      provision(fake, {
        QuickLaunch: [
          { Title: 'Ny', Url: '/sites/demo/ny' },
          {
            Title: 'Hjem',
            Url: '/sites/demo/SitePages/Home.aspx',
            IgnoreExisting: true
          },
          {
            Title: 'Dokumenter',
            Url: '/sites/demo/Missing',
            IgnoreExisting: true
          }
        ]
      }),
      {
        message: /^Navigation node 'Dokumenter' \(\/sites\/demo\/Missing\) was refused by SharePoint: Kan ikke åpne/
      }
    )

    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'Ny' /sites/demo/ny",
      'update 1 {"Url":"/sites/demo/SitePages/Home.aspx"}',
      'update 2 {"Url":"/sites/demo/Missing"}',
      // The rollback, the last write first.
      'update 1 {"Url":"/sites/demo"}',
      'delete 3'
    ])
    assert.deepEqual(fake.tree('quicklaunch'), before)
  })

  it('a refused child under a new parent removes the parent, and its children with it', async () => {
    const fake = createFakeWeb({ refuse: refuseMissing })

    await assert.rejects(
      provision(fake, {
        QuickLaunch: [
          {
            Title: 'Parent',
            Url: '/sites/demo/parent',
            Children: [
              { Title: 'Ok', Url: '/sites/demo/ok' },
              { Title: 'Missing', Url: '/sites/demo/Missing.aspx' }
            ]
          }
        ]
      }),
      /Navigation node 'Missing' \(\/sites\/demo\/Missing\.aspx\)/
    )

    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'Parent' /sites/demo/parent",
      "add quicklaunch(1)/children 'Ok' /sites/demo/ok",
      "add quicklaunch(1)/children 'Missing' /sites/demo/Missing.aspx",
      'delete 1'
    ])
    assert.deepEqual(fake.tree('quicklaunch'), [])
  })

  it('a refusal in the Quick Launch sends nothing to the top navigation bar', async () => {
    const fake = createFakeWeb({ refuse: refuseMissing })
    fake.seed('topnavigationbar', [{ Title: 'Topp', Url: '/sites/demo' }])

    await assert.rejects(
      provision(fake, {
        QuickLaunch: [{ Title: 'Missing', Url: '/sites/demo/Missing' }],
        TopNavigationBar: [{ Title: 'Ny', Url: '/sites/demo/ny' }]
      }),
      /Navigation node 'Missing'/
    )

    assert.deepEqual(
      fake.calls.filter((call) => /topnavigationbar/.test(call)),
      []
    )
  })

  it('a refusal in the top navigation bar undoes the Quick Launch, which deletes nothing', async () => {
    const fake = createFakeWeb({ refuse: refuseMissing })
    fake.seed('quicklaunch', [
      { Title: 'Hjem', Url: '/sites/demo' },
      { Title: 'Owner', Url: '/sites/demo/owner' }
    ])
    const before = fake.tree('quicklaunch')

    await assert.rejects(
      provision(fake, {
        QuickLaunch: [
          { Title: 'Ny', Url: '/sites/demo/ny' },
          { Title: 'Hjem', Url: '/sites/demo/home', IgnoreExisting: true }
        ],
        TopNavigationBar: [{ Title: 'Missing', Url: '/sites/demo/Missing' }]
      }),
      /Navigation node 'Missing'/
    )

    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'Ny' /sites/demo/ny",
      'update 1 {"Url":"/sites/demo/home"}',
      "add topnavigationbar 'Missing' /sites/demo/Missing",
      'update 1 {"Url":"/sites/demo"}',
      'delete 3'
    ])
    assert.deepEqual(fake.tree('quicklaunch'), before)
  })

  it('a collection that cannot be read fails the run, naming it, and undoes the other one', async () => {
    const fake = createFakeWeb({
      failRead: (name) => name === 'topnavigationbar'
    })
    fake.seed('quicklaunch', [{ Title: 'Owner', Url: '/sites/demo/owner' }])

    await assert.rejects(
      provision(fake, {
        QuickLaunch: [{ Title: 'Ny', Url: '/sites/demo/ny' }],
        TopNavigationBar: [{ Title: 'Topp', Url: '/sites/demo' }]
      }),
      {
        message: 'Failed to read the top navigation bar: Too many requests'
      }
    )

    assert.deepEqual(fake.writes(), [
      "add quicklaunch 'Ny' /sites/demo/ny",
      'delete 2'
    ])
    assert.deepEqual(fake.tree('quicklaunch'), [
      { Title: 'Owner', Url: '/sites/demo/owner', children: [] }
    ])
  })

  it('a rollback step that fails is logged as an error naming the node left behind', async () => {
    const fake = createFakeWeb({
      refuse: refuseMissing,
      failDelete: (id) => id === 1
    })

    const { entries } = await captureLogs(() =>
      assert.rejects(
        provision(fake, {
          QuickLaunch: [
            { Title: 'Added', Url: '/sites/demo/added' },
            { Title: 'Missing', Url: '/sites/demo/Missing' }
          ]
        }),
        /Navigation node 'Missing'/
      )
    )

    assert.equal(
      levelOf(
        entries,
        /Navigation node 'Added' \(1\), added by this run, is left in the Quick Launch: Too many requests/
      ),
      LogLevel.Error
    )
  })

  it('a rollback delete SharePoint answers but does not apply is found by reading back', async () => {
    const fake = createFakeWeb({
      refuse: refuseMissing,
      ignoreDelete: (id) => id === 1
    })

    const { entries } = await captureLogs(() =>
      assert.rejects(
        provision(fake, {
          QuickLaunch: [
            { Title: 'Added', Url: '/sites/demo/added' },
            { Title: 'Missing', Url: '/sites/demo/Missing' }
          ]
        })
      )
    )

    assert.equal(
      levelOf(
        entries,
        /Navigation node 'Added' \(1\), added by this run, is left in the Quick Launch: SharePoint reported success/
      ),
      LogLevel.Error
    )
  })
})

describe('Navigation: a failed delete or move fails the run', () => {
  it('reads the order back after moving and fails when a move did not take', async () => {
    const fake = createFakeWeb({ ignoreMove: () => true })
    fake.seed('quicklaunch', [
      { Title: 'A', Url: '/sites/demo/a' },
      { Title: 'B', Url: '/sites/demo/b' }
    ])

    const { entries } = await captureLogs(() =>
      assert.rejects(
        provision(fake, {
          QuickLaunch: [
            { Title: 'B', Url: '/sites/demo/b' },
            { Title: 'A', Url: '/sites/demo/a' }
          ]
        }),
        {
          message:
            "The nodes in the Quick Launch are not in the template's order after moving them: read 1, 2, expected 2, 1."
        }
      )
    )

    assert.equal(
      levelOf(entries, /not in the template's order/),
      LogLevel.Error
    )
  })

  it('names the node it could not delete', async () => {
    const fake = createFakeWeb({ failDelete: () => true })
    fake.seed('quicklaunch', [{ Title: 'Gammel', Url: '/sites/demo/gammel' }])

    const { entries } = await captureLogs(() =>
      assert.rejects(
        provision(fake, {
          QuickLaunch: [{ Title: 'Ny', Url: '/sites/demo/ny' }]
        }),
        {
          message: /^Failed to delete navigation node 'Gammel' \(1\) from the Quick Launch: Too many requests$/
        }
      )
    )

    assert.equal(
      levelOf(entries, /Failed to delete navigation node 'Gammel'/),
      LogLevel.Error
    )
  })

  it('names the node it could not move', async () => {
    const fake = createFakeWeb({ failMove: () => true })
    fake.seed('quicklaunch', [
      { Title: 'A', Url: '/sites/demo/a' },
      { Title: 'B', Url: '/sites/demo/b' }
    ])

    const { entries } = await captureLogs(() =>
      assert.rejects(
        provision(fake, {
          QuickLaunch: [
            { Title: 'B', Url: '/sites/demo/b' },
            { Title: 'A', Url: '/sites/demo/a' }
          ]
        }),
        {
          message: /^Failed to move navigation node 'A' after 'B' in the Quick Launch: Ukjent feil$/
        }
      )
    )

    assert.equal(
      levelOf(entries, /Failed to move navigation node 'A'/),
      LogLevel.Error
    )
  })
})
