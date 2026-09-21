import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { Navigation } from '../../src/handlers/navigation'

// `replaceUrlTokens` resolves `{hosturl}` from `window` on every call, so the
// handler needs one even when a node url contains no tokens.
const globals: any = globalThis
globals.window = {
  location: { protocol: 'https:', host: 'contoso.sharepoint.com', port: '' }
}

interface IFakeNodeCollection {
  /** The fake `INavigationNodes`, passed to the handler */
  nodes: any
  /** Node infos added through this collection, in order */
  added: any[]
  /** Node ids deleted through this collection, in order */
  deleted: number[]
  /** Child collections, keyed by the id of their parent node */
  children: Map<number, IFakeNodeCollection>
}

/**
 * Hand-rolled stand-in for a PnPjs v4 `INavigationNodes` collection: invokable
 * to read the current nodes, `add` resolving with an `INavNodeInfo` (v4 no
 * longer returns a `{ node, data }` wrapper) and `getById` exposing `children`.
 */
function createNodeCollection(): IFakeNodeCollection {
  const collection: IFakeNodeCollection = {
    nodes: undefined,
    added: [],
    deleted: [],
    children: new Map()
  }
  let current: any[] = []
  let nextId = 1

  const nodes: any = async () => [...current]
  nodes.add = async (title: string, url: string) => {
    const info = { Id: nextId++, Title: title, Url: url }
    collection.added.push(info)
    current.push(info)
    return info
  }
  nodes.getById = (id: number) => {
    if (!collection.children.has(id)) {
      collection.children.set(id, createNodeCollection())
    }
    return {
      children: collection.children.get(id).nodes,
      delete: async () => {
        collection.deleted.push(id)
        current = current.filter((node) => node.Id !== id)
      }
    }
  }

  collection.nodes = nodes
  return collection
}

describe('Navigation: v4 node creation result', () => {
  it('adds child nodes under the node returned by add', async () => {
    const handler: any = new Navigation({})
    const root = createNodeCollection()

    await handler.processNavTree(root.nodes, [
      {
        Title: 'Prosjekter',
        Url: '/sites/demo/Lists/Prosjekter',
        Children: [{ Title: 'Aktive', Url: '/sites/demo/aktive' }]
      }
    ])

    assert.deepEqual(
      root.added.map((node) => node.Title),
      ['Prosjekter']
    )
    const parentId = root.added[0].Id
    const childCollection = root.children.get(parentId)
    assert.ok(childCollection, 'children were not read from getById(parentId)')
    assert.deepEqual(
      childCollection.added.map((node) => node.Title),
      ['Aktive']
    )
  })

  it('replaces the existing nodes rather than appending to them', async () => {
    const handler: any = new Navigation({})
    const root = createNodeCollection()
    await root.nodes.add('Gammel', '/sites/demo/gammel')
    root.added.length = 0

    await handler.processNavTree(root.nodes, [
      { Title: 'Ny', Url: '/sites/demo/ny' }
    ])

    assert.deepEqual(root.deleted, [1])
    assert.deepEqual(
      root.added.map((node) => node.Title),
      ['Ny']
    )
  })
})
