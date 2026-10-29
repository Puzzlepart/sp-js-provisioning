import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import {
  INavigationNodeState,
  INavigationState,
  navigationIds,
  readNavigation,
  readNavigationTree,
  withoutIds
} from '../../debug/lib/readback/navigation'
import { INavigationNode } from '../../src/schema'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

/**
 * The template's nodes as the read-back should show them: `{site}` (the handler's only token)
 * replaced, a web-relative URL (`SitePages`, as the hosting packages write them) resolved
 * against the web by SharePoint, children in order, no ids.
 *
 * @param nodes - The template's nodes
 * @param urls - What `{site}` stands for, and the web's server-relative URL
 */
function declared(
  nodes: INavigationNode[] | undefined,
  urls: { site: string; web: string }
): INavigationNodeState[] {
  return (nodes || []).map((node) => {
    const url = node.Url.replace(/{site}/g, urls.site)
    return {
      Title: node.Title,
      Url: /^(\/|https?:)/.test(url) ? url : `${urls.web}/${url}`,
      children: declared(node.Children, urls)
    }
  })
}

/**
 * Every title in a read-back, depth first.
 *
 * @param state - The read-back
 */
function titles(state: INavigationState): string[] {
  const of = (nodes: INavigationNodeState[]): string[] =>
    nodes.reduce(
      (all: string[], node) => [...all, node.Title, ...of(node.children)],
      []
    )
  return [...of(state.quickLaunch), ...of(state.topNavigationBar)]
}

/**
 * The Navigation handler on a scratch web: Quick Launch and top navigation bar with children;
 * `{site}`, web-relative and absolute URLs; `IgnoreExisting`. The handler deletes every node of a
 * collection it provisions and adds the template's, so each run re-adds the nodes (new ids). A
 * top-level node whose title exists exactly once before the run keeps that node's URL unless it
 * sets `IgnoreExisting` (`navigation.b.json`).
 *
 * SharePoint opens a server-relative URL that is not under the current web's `_layouts` as a file
 * or folder and refuses the node when there is none (500, "Filen eller mappen finnes ikke"). On a
 * subweb `{site}` is the site collection, so `{site}/_layouts/...` is refused there: the fixtures
 * point `{site}` URLs at folders every site collection has (`_catalogs/...`).
 */
describe('Navigation', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('navigation')
  let defaults: INavigationState

  before(async () => {
    // A fresh STS#3 web's own Quick Launch (Home, Documents, Pages, ... in the web's language).
    defaults = await readNavigation(scratch().conn)
  })

  /**
   * What `{site}` stands for (`pageContext.site`: the site collection, not the scratch web below
   * it), and the scratch web.
   */
  function urls() {
    const { conn } = scratch()
    return {
      site: conn.siteInfo.ServerRelativeUrl,
      web: conn.webInfo.ServerRelativeUrl
    }
  }

  it('replaces the existing nodes with the template, children and order kept, and re-adds the same nodes on a second apply', async () => {
    const { conn } = scratch()
    const { site, web } = urls()
    assert.notEqual(site, web)
    assert.ok(
      defaults.quickLaunch.length > 0,
      'a fresh scratch web has Quick Launch nodes of its own'
    )
    // A fresh scratch web has no top navigation bar nodes: give it one to replace.
    await conn.web.navigation.topNavigationBar.add('SPJS Itest Stale', web)
    const existing = titles(await readNavigation(conn))
    assert.ok(existing.includes('SPJS Itest Stale'))

    const ids: number[][] = []
    const { first, afterFirst } = await applyTwice(
      conn,
      () => loadFixture('navigation'),
      async (c) => {
        const tree = await readNavigationTree(c)
        ids.push(navigationIds(tree))
        return withoutIds(tree)
      }
    )
    assert.deepEqual(first.order, ['Navigation'])

    const { Navigation } = loadFixture('navigation')
    assert.deepEqual(afterFirst, {
      quickLaunch: declared(Navigation.QuickLaunch, urls()),
      topNavigationBar: declared(Navigation.TopNavigationBar, urls())
    })
    // `{site}` became the site collection's path; `SitePages` the web's library.
    assert.equal(afterFirst.quickLaunch[0].Url, site)
    assert.equal(afterFirst.quickLaunch[1].Url, `${web}/SitePages`)
    // The scratch web's own nodes and the stale one are gone, not kept beside the template's.
    const provisioned = titles(afterFirst)
    for (const title of existing)
      assert.ok(!provisioned.includes(title), `${title} is still there`)

    // The second run deleted and re-added every node: same titles, URLs and order (asserted by
    // applyTwice), new ids.
    assert.equal(ids[1].length, ids[0].length)
    assert.deepEqual(
      ids[1].filter((id) => ids[0].includes(id)),
      [],
      'the second run kept node ids'
    )
  })

  it('keeps the URL of an existing top-level title unless IgnoreExisting (B)', async () => {
    const { conn } = scratch()
    const { site, web } = urls()
    assertClean(await apply(conn, loadFixture('navigation')), 'the A run')
    const b = await apply(conn, loadFixture('navigation.b'))
    assertClean(b, 'the B run')
    assert.deepEqual(b.order, ['Navigation'])

    assert.deepEqual(await readNavigation(conn), {
      quickLaunch: [
        // `IgnoreExisting`: the template's URL.
        {
          Title: 'SPJS Itest Fabrikam',
          Url: 'https://www.fabrikam.com/spjs-itest/b',
          children: []
        },
        // An existing title: A's URL, not B's.
        { Title: 'SPJS Itest Home', Url: site, children: [] },
        {
          Title: 'SPJS Itest Pages',
          Url: `${web}/SitePages`,
          // A child never keeps its URL: its parent is re-added without children, so no
          // existing child has its title (B's URL, although A had the same title). A's other
          // children are gone.
          children: [
            {
              Title: 'SPJS Itest Master pages',
              Url: 'https://www.contoso.com/spjs-itest/b/child',
              children: []
            }
          ]
        },
        // A new title: the template's URL.
        { Title: 'SPJS Itest New', Url: `${site}/_catalogs/wp`, children: [] }
      ],
      // A's URL; A's children and A's other node are gone, as B has none.
      topNavigationBar: [
        { Title: 'SPJS Itest Top', Url: `${site}/_catalogs/wp`, children: [] }
      ]
    })
  })

  it(
    'a node SharePoint refuses fails the run and leaves the Quick Launch as it was',
    {
      todo:
        'the handler deletes every node before adding any: a refused node (a missing file) ' +
        'stops the run with the Quick Launch holding only the nodes added before it'
    },
    async () => {
      const { conn } = scratch()
      assertClean(await apply(conn, loadFixture('navigation')), 'the A run')
      const before = await readNavigation(conn)
      const result = await apply(conn, {
        Navigation: {
          QuickLaunch: [
            { Title: 'SPJS Itest Added', Url: 'SitePages' },
            {
              Title: 'SPJS Itest Missing',
              Url: 'SitePages/SPJSItestMissing.aspx'
            }
          ]
        }
      })
      // Compare strings: the error carries the HTTP response (headers, cookies).
      assert.equal(result.error && result.error.handler, 'Navigation')
      assert.match(
        String(result.error && result.error.message),
        /SPJSItestMissing\.aspx/
      )
      const after = await readNavigation(conn)
      // Top navigation is not in the template and is untouched (holds today).
      assert.deepEqual(after.topNavigationBar, before.topNavigationBar)
      // Today: only 'SPJS Itest Added' is left.
      assert.deepEqual(after.quickLaunch, before.quickLaunch)
    }
  )
})
