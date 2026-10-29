import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
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
 * A node of a read-back with ids, by title (depth first, both collections).
 *
 * @param state - The read-back
 * @param title - The node's title
 */
function nodeOf(
  state: INavigationState<number>,
  title: string
): INavigationNodeState<number> {
  const find = (
    nodes: Array<INavigationNodeState<number>>
  ): INavigationNodeState<number> | undefined =>
    nodes.reduce(
      (found: INavigationNodeState<number> | undefined, node) =>
        found || (node.Title === title ? node : find(node.children)),
      undefined
    )
  const node = find([...state.quickLaunch, ...state.topNavigationBar])
  assert.ok(node, `${title} is not in the navigation`)
  return node
}

/**
 * A connection whose web records every request the engine sends through it (method and path
 * below the site), so a test can see what a run asked for, not only what it left behind.
 *
 * @param conn - The scratch web
 * @param requests - Where the requests go
 */
function recording(conn: ISiteConnection, requests: string[]): ISiteConnection {
  return Object.create(conn, {
    web: {
      get: () => {
        const web = conn.sp.web
        web.on.pre(function (url: string, init: any, result: any) {
          // PnPjs sends MERGE and DELETE as a POST with X-HTTP-Method.
          const tunnelled =
            init &&
            init.headers &&
            new Headers(init.headers).get('X-HTTP-Method')
          requests.push(
            `${
              tunnelled || (init && init.method) || 'GET'
            } ${decodeURIComponent(url.slice(conn.url.length))}`
          )
          return Promise.resolve<[string, any, any]>([url, init, result])
        })
        return web
      }
    }
  })
}

/** The requests that write: anything but a GET. */
const writesOf = (requests: string[]) =>
  requests.filter((request) => !request.startsWith('GET '))

/**
 * The Navigation handler on a scratch web: Quick Launch and top navigation bar with children
 * (three levels); `{site}`, web-relative and absolute URLs; `IgnoreExisting`. The template owns
 * each collection it names (docs/specs/navigation-re-apply.md): a node is matched to a stored one
 * at the same level by exact title and kept with its id, URL and audiences; a missing node is
 * added, a node the template does not declare is deleted, and the nodes are moved into the
 * template's order. A kept node keeps its stored URL, at every level, unless it sets
 * `IgnoreExisting` (`navigation.b.json`). An unchanged re-apply writes nothing.
 *
 * SharePoint opens a server-relative URL that is not under the current web's `_layouts` as a file
 * or folder and refuses the node when there is none (500, "Filen eller mappen finnes ikke"). On a
 * subweb `{site}` is the site collection, so `{site}/_layouts/...` is refused there: the fixtures
 * point `{site}` URLs at folders every site collection has (`_catalogs/...`). A refused node fails
 * the run and leaves both collections as they were.
 *
 * The tests share one scratch web and run in order, each from the menu the one before left.
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

  it('replaces the existing nodes with the template, children and order kept, and a second apply keeps every node and writes nothing', async () => {
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
    const requests: string[] = []
    // Where each read-back started and ended in `requests`: the second run lies between the
    // first read-back's end and the second's start.
    const marks: number[] = []
    const { first, afterFirst } = await applyTwice(
      recording(conn, requests),
      () => loadFixture('navigation'),
      async (c) => {
        marks.push(requests.length)
        const tree = await readNavigationTree(c)
        marks.push(requests.length)
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

    // The second run kept every node (same titles, URLs and order, asserted by applyTwice, and
    // the same ids) and sent nothing but reads: one per collection, three levels deep.
    assert.deepEqual(ids[1], ids[0], 'the second run changed node ids')
    const second = requests.slice(marks[1], marks[2])
    assert.deepEqual(writesOf(second), [], 'the second run wrote')
    assert.deepEqual(
      second
        .filter((request) => /\/navigation\//i.test(request))
        .map((request) =>
          /navigation\/(\w+)\?.*\$expand=([^&]*)/i.exec(request)
        )
        .map((match) => match && `${match[1]} ${match[2]}`),
      [
        'quicklaunch Children,Children/Children,Children/Children/Children',
        'topnavigationbar Children,Children/Children'
      ]
    )
  })

  it('keeps the nodes an owner changed (ids, audiences, a child URL), removes the owner node and restores the order with moves', async () => {
    const { conn } = scratch()
    const tree = await readNavigationTree(conn)
    const quickLaunch = () => conn.web.navigation.quicklaunch
    const home = nodeOf(tree, 'SPJS Itest Home').Id
    const fabrikam = nodeOf(tree, 'SPJS Itest Fabrikam').Id
    const pages = nodeOf(tree, 'SPJS Itest Pages').Id
    const contoso = nodeOf(tree, 'SPJS Itest Contoso').Id
    // What a site owner does between two runs.
    const owner = await quickLaunch().add(
      'SPJS Itest Owner',
      'https://www.contoso.com/spjs-itest/owner'
    )
    await quickLaunch().moveAfter(fabrikam, home)
    await quickLaunch()
      .getById(pages)
      .children.moveAfter(
        nodeOf(tree, 'SPJS Itest Contents').Id,
        nodeOf(tree, 'SPJS Itest Master pages').Id
      )
    await quickLaunch()
      .getById(contoso)
      .update({ Url: 'https://www.contoso.com/spjs-itest/owner-child' })
    // SharePoint stores audience ids as given, without checking them (probe N3c).
    const audiences = ['1c0ffee0-5915-4c3a-9e2e-000000000001']
    await quickLaunch().getById(home).update({ AudienceIds: audiences })
    assert.deepEqual(
      (await readNavigation(conn)).quickLaunch.map((node) => node.Title),
      [
        'SPJS Itest Home',
        'SPJS Itest Fabrikam',
        'SPJS Itest Pages',
        'SPJS Itest Owner'
      ]
    )

    const requests: string[] = []
    assertClean(
      await apply(recording(conn, requests), loadFixture('navigation')),
      'the run'
    )

    // One delete (the owner's node) and two moves (Pages after Home, Contoso after Master
    // pages): nothing added or updated.
    const writes = writesOf(requests)
    assert.equal(writes.length, 3, writes.join('\n'))
    assert.match(
      writes[0],
      new RegExp(`^DELETE .*quicklaunch\\(${owner.Id}\\)`)
    )
    assert.match(writes[1], /^POST .*quicklaunch\/MoveAfter$/i)
    assert.match(
      writes[2],
      new RegExp(`^POST .*quicklaunch\\(${pages}\\)/children/MoveAfter$`, 'i')
    )

    const after = await readNavigationTree(conn)
    const { Navigation } = loadFixture('navigation')
    const expected = declared(Navigation.QuickLaunch, urls())
    // The child keeps the owner's URL: the template's has no IgnoreExisting.
    expected[1].children[1].Url =
      'https://www.contoso.com/spjs-itest/owner-child'
    assert.deepEqual(withoutIds(after).quickLaunch, expected)
    assert.deepEqual(
      navigationIds(after),
      navigationIds(tree),
      'the run changed node ids'
    )
    assert.deepEqual(
      (
        await quickLaunch().getById(home).select('AudienceIds')<{
          AudienceIds: string[] | null
        }>()
      ).AudienceIds,
      audiences
    )
  })

  it('keeps the URL of an existing title at every level unless IgnoreExisting, and the ids of the nodes it keeps (B)', async () => {
    const { conn } = scratch()
    const { site, web } = urls()
    assertClean(await apply(conn, loadFixture('navigation')), 'the A run')
    const a = await readNavigationTree(conn)
    const b = await apply(conn, loadFixture('navigation.b'))
    assertClean(b, 'the B run')
    assert.deepEqual(b.order, ['Navigation'])

    const after = await readNavigationTree(conn)
    assert.deepEqual(withoutIds(after), {
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
          children: [
            // A child keeps its stored URL too (A's, not B's).
            {
              Title: 'SPJS Itest Master pages',
              Url: `${site}/_catalogs/masterpage`,
              children: []
            },
            // `IgnoreExisting` on a child: the template's URL. B declares no children for it,
            // so A's grandchild is gone, and so is A's 'Contents' child.
            {
              Title: 'SPJS Itest Contoso',
              Url: 'https://www.contoso.com/spjs-itest/b/contoso',
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
    // The nodes B kept are A's nodes, updated in place.
    for (const title of [
      'SPJS Itest Fabrikam',
      'SPJS Itest Home',
      'SPJS Itest Pages',
      'SPJS Itest Master pages',
      'SPJS Itest Contoso',
      'SPJS Itest Top'
    ])
      assert.equal(nodeOf(after, title).Id, nodeOf(a, title).Id, title)
  })

  it('a node SharePoint refuses fails the run and leaves the Quick Launch as it was', async () => {
    const { conn } = scratch()
    assertClean(await apply(conn, loadFixture('navigation')), 'the A run')
    const before = await readNavigationTree(conn)
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
      /^Navigation node 'SPJS Itest Missing' \(SitePages\/SPJSItestMissing\.aspx\) was refused by SharePoint: .*SPJSItestMissing\.aspx/
    )
    // Nothing deleted, the node added before the refused one removed again: the same nodes, ids
    // included. The top navigation bar is not in the template and is untouched.
    assert.deepEqual(await readNavigationTree(conn), before)
  })

  it('a node refused in the top navigation bar undoes the adds and URL changes in the Quick Launch', async () => {
    const { conn } = scratch()
    const before = await readNavigationTree(conn)
    const result = await apply(conn, {
      Navigation: {
        QuickLaunch: [
          {
            Title: 'SPJS Itest Fabrikam',
            Url: 'https://www.fabrikam.com/spjs-itest/changed',
            IgnoreExisting: true
          },
          { Title: 'SPJS Itest Home', Url: '{site}' },
          { Title: 'SPJS Itest Added', Url: 'SitePages' }
        ],
        TopNavigationBar: [
          { Title: 'SPJS Itest Top', Url: '{site}/_catalogs/wp' },
          {
            Title: 'SPJS Itest Missing top',
            Url: 'SitePages/SPJSItestMissingTop.aspx'
          }
        ]
      }
    })
    assert.equal(result.error && result.error.handler, 'Navigation')
    assert.match(
      String(result.error && result.error.message),
      /^Navigation node 'SPJS Itest Missing top' \(SitePages\/SPJSItestMissingTop\.aspx\) was refused by SharePoint: /
    )
    // The rollback left nothing behind (it logs an error per node it could not undo).
    assert.deepEqual(
      result.errors
        .map((entry) => entry.message)
        .filter((message) => /\(rollback\)/.test(message)),
      []
    )
    assert.deepEqual(await readNavigationTree(conn), before)
  })

  it('a node whose target is deleted goes with it, and re-adding it is refused without changing the menu', async () => {
    const { conn } = scratch()
    const { site } = urls()
    const library = (
      await conn.web.defaultDocumentLibrary.rootFolder.select(
        'ServerRelativeUrl'
      )<{ ServerRelativeUrl: string }>()
    ).ServerRelativeUrl
    const target = `${library}/SPJSItestTarget.txt`
    await conn.web
      .getFolderByServerRelativePath(library)
      .files.addUsingPath('SPJSItestTarget.txt', 'SPJS Itest', {
        Overwrite: true
      })
    const template = () => ({
      Navigation: {
        QuickLaunch: [
          { Title: 'SPJS Itest Home', Url: '{site}' },
          { Title: 'SPJS Itest Target', Url: target }
        ]
      }
    })
    assertClean(await apply(conn, template()), 'the first run')
    assert.deepEqual((await readNavigation(conn)).quickLaunch, [
      { Title: 'SPJS Itest Home', Url: site, children: [] },
      { Title: 'SPJS Itest Target', Url: target, children: [] }
    ])

    await conn.web.getFileByServerRelativePath(target).recycle()
    // SharePoint removes the nodes of a target that is recycled (probe N3b).
    const recycled = await readNavigationTree(conn)
    assert.deepEqual(
      recycled.quickLaunch.map((node) => node.Title),
      ['SPJS Itest Home']
    )

    // The unchanged template now names a missing file: its node is added again and refused.
    const result = await apply(conn, template())
    assert.equal(result.error && result.error.handler, 'Navigation')
    assert.match(
      String(result.error && result.error.message),
      /^Navigation node 'SPJS Itest Target' \(.*\/SPJSItestTarget\.txt\) was refused by SharePoint: /
    )
    assert.deepEqual(await readNavigationTree(conn), recycled)
  })
})
