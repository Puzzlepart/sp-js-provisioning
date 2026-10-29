import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { PermissionKind, SPInstance } from '@pnp/sp/presets/all'
import { ISiteConnection } from '../../debug/lib/connect'
import { skipReason } from '../../debug/lib/env'
import {
  createScratchWeb,
  deleteScratchWeb,
  groupsOfScratchWeb,
  listScratchWebs
} from '../../debug/lib/scratch'
import { connectToSite } from './helpers'

/**
 * Can the suites run against the reserved site? Run first by `npm run test:integration`, and alone
 * by `npm run check:tenant`. Each test names what is missing; nothing here leaves anything behind.
 */
describe('reachability', { skip: skipReason() }, () => {
  let conn: ISiteConnection

  before(async () => {
    conn = await connectToSite()
  })

  it('signs in and reads the reserved site', (t) => {
    // Server-relative URLs only: reports go to CI artifacts (no tenant URL).
    t.diagnostic(
      `signed in with ${conn.mode} to ${conn.webInfo.ServerRelativeUrl}`
    )
    t.diagnostic(
      `web "${conn.webInfo.Title}", language ${conn.webInfo.Language}, site collection ${conn.siteInfo.ServerRelativeUrl}`
    )
    assert.ok(conn.siteInfo.Id)
    assert.ok(conn.webInfo.Id)
  })

  it('signs in as a site collection administrator', async (t) => {
    // Cleanup deletes site groups, including one that owns itself.
    if (conn.mode === 'certificate') {
      t.skip(
        'app-only: the app has no site admin flag; it needs Sites.FullControl.All'
      )
      return
    }
    const user = await conn.web.currentUser.select('Title', 'IsSiteAdmin')<{
      Title: string
      IsSiteAdmin: boolean
    }>()
    t.diagnostic(`site collection administrator: ${user.IsSiteAdmin}`)
    assert.ok(
      user.IsSiteAdmin,
      `the signed-in account must be a site collection administrator of ${conn.siteInfo.ServerRelativeUrl}`
    )
  })

  it('reads the default term store ({sitecollectiontermstoreid})', async (t) => {
    try {
      const store = await SPInstance(
        [conn.web, conn.url],
        '_api/v2.1/termstore'
      ).select('id')<{ id: string }>()
      t.diagnostic(`term store ${store.id}`)
      assert.ok(store.id)
    } catch (error) {
      if (conn.mode === 'certificate' && error && error.status === 403) {
        t.skip('SharePoint refuses the app (403): it needs TermStore.Read.All')
        return
      }
      throw error
    }
  })

  it('reports whether custom script is allowed', async (t) => {
    // NoScript (DenyAddAndCustomizePages) removes the permission even for admins; the
    // ScriptLink and PropertyBagEntries suites skip when it is missing.
    const allowed = await conn.web.currentUserHasPermissions(
      PermissionKind.AddAndCustomizePages
    )
    t.diagnostic(
      allowed
        ? 'custom script is allowed'
        : 'custom script is blocked (DenyAddAndCustomizePages): the ScriptLink and PropertyBagEntries suites will skip'
    )
  })

  it('creates and deletes a scratch web with its own associated groups', async (t) => {
    const scratch = await createScratchWeb(conn, 'reachability')
    try {
      const unique = await scratch.conn.web.select('HasUniqueRoleAssignments')<{
        HasUniqueRoleAssignments: boolean
      }>()
      assert.equal(unique.HasUniqueRoleAssignments, true)
      const owners = await scratch.conn.web.associatedOwnerGroup.select(
        'Title'
      )<{
        Title: string
      }>()
      assert.match(owners.Title, new RegExp(scratch.stamp))
      t.diagnostic(
        `scratch web ${scratch.conn.webInfo.ServerRelativeUrl}, owners "${owners.Title}"`
      )
    } finally {
      const { groups } = await deleteScratchWeb(conn, scratch)
      t.diagnostic(`deleted it and ${groups.length} group(s)`)
    }
    assert.deepEqual(await groupsOfScratchWeb(conn, scratch.title), [])
  })

  it('reports leftover scratch webs', async (t) => {
    const webs = await listScratchWebs(conn)
    const now = Date.now()
    for (const web of webs) {
      const minutes = Math.round((now - web.created.getTime()) / 60_000)
      t.diagnostic(
        `${web.serverRelativeUrl}: ${minutes} min old${
          minutes < 120 ? ' (another run may be active)' : ''
        }`
      )
    }
    if (!webs.length) t.diagnostic('no leftover scratch webs')
  })
})
