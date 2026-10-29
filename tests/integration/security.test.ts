import { Web } from '@pnp/sp/presets/all'
import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import { connect } from '../../debug/lib/connect'
import { env } from '../../debug/lib/env'
import { readSecurity } from '../../debug/lib/readback/security'
import { describeError } from '../../debug/lib/redact'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

/**
 * The Security handler on a scratch web: site groups (owners, settings, members), and role
 * inheritance and role assignments on the web. The groups are named after the web through
 * `{sitename}`, so their titles start with its title and they are deleted with it.
 */
describe('Security', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('security')
  let user: string | undefined
  let roles: {
    limitedAccess: string
    read: string
    contribute: string
    fullControl: string
  }

  before(async () => {
    const { conn } = scratch()
    // The login name SharePoint stores (claims), which the members read back as: the signed-in
    // user, or app-only INTEGRATION_USER as SharePoint resolves it.
    user =
      conn.mode === 'cookies'
        ? (
            await conn.web.currentUser.select('LoginName')<{
              LoginName: string
            }>()
          ).LoginName
        : env.user && (await conn.web.ensureUser(env.user)).LoginName
    const roleName = async (kind: 1 | 2 | 3 | 5) =>
      (
        await conn.web.roleDefinitions.getByType(kind).select('Name')<{
          Name: string
        }>()
      ).Name
    roles = {
      limitedAccess: await roleName(1),
      read: await roleName(2),
      contribute: await roleName(3),
      fullControl: await roleName(5)
    }
  })

  it('creates groups with owners, settings and members, grants them on the web, and changes nothing on a second apply', async (t) => {
    if (!user) {
      t.skip('app-only without INTEGRATION_USER: no user for the member')
      return
    }
    const { conn } = scratch()
    const site = conn.webInfo.Title
    const titles = {
      readers: `${site} Readers`,
      owners: `${site} Owners`,
      members: `${site} Members`
    }
    // The web's security, and the role assignments of a list that grants a group created by the
    // same template (Security runs before Lists).
    const read = async (c: typeof conn) => ({
      ...(await readSecurity(c, Object.values(titles))),
      list: (
        await c.web.lists
          .getByTitle('SPJS Itest Security List')
          .roleAssignments.expand('Member', 'RoleDefinitionBindings')<
          Array<{
            Member: { Title: string }
            RoleDefinitionBindings: Array<{ Name: string }>
          }>
        >()
      ).reduce((all, assignment) => {
        all[assignment.Member.Title] = assignment.RoleDefinitionBindings.map(
          (binding) => binding.Name
        ).sort()
        return all
      }, {} as Record<string, string[]>)
    })
    const { afterFirst } = await applyTwice(
      conn,
      () => loadFixture('security', { user }),
      read
    )
    const ownerGroup = (
      await conn.web.associatedOwnerGroup.select('Title')<{ Title: string }>()
    ).Title
    const visitorGroup = (
      await conn.web.associatedVisitorGroup.select('Title')<{ Title: string }>()
    ).Title
    const [readers, owners, members] = afterFirst.groups

    // Owners: the associated owner group through its token, and a group that owns itself (set
    // through CSOM; REST's SetUserAsOwner ignored a group id).
    assert.deepEqual(
      afterFirst.groups.map((group) => [group.Title, group.Owner]),
      [
        [titles.readers, ownerGroup],
        [titles.owners, titles.owners],
        [titles.members, ownerGroup]
      ]
    )
    assert.equal(readers.Description, `Readers of ${site}`)
    assert.equal(readers.AllowMembersEditMembership, false)
    assert.equal(readers.OnlyAllowMembersViewMembership, true)
    assert.equal(owners.AllowRequestToJoinLeave, true)
    assert.equal(owners.AutoAcceptRequestToJoinLeave, false)
    assert.equal(
      owners.RequestToJoinLeaveEmailSetting,
      'spjs-itest@contoso.com'
    )
    // ClearExistingMembers leaves exactly the template's members.
    assert.deepEqual(members.members, [user.toLowerCase()])

    // The scratch web already has unique permissions: inheritance is left as is.
    assert.equal(afterFirst.hasUniqueRoleAssignments, true)
    // `Read`/`Contribute` resolve by role type on a site in any language.
    // Limited Access on the web comes with the list grant below: SharePoint adds it so the group
    // can reach the list.
    assert.deepEqual(
      afterFirst.roleAssignments[titles.readers],
      [roles.limitedAccess, roles.read].sort()
    )
    assert.deepEqual(afterFirst.roleAssignments[titles.members], [
      roles.contribute
    ])
    assert.ok(afterFirst.roleAssignments[visitorGroup].includes(roles.read))

    // List Security grants the group created by the same template.
    assert.deepEqual(afterFirst.list[titles.readers], [roles.read])
    assert.deepEqual(afterFirst.list[ownerGroup], [roles.fullControl])
  })

  it('removes a member added since, with ClearExistingMembers', async (t) => {
    if (!user) {
      t.skip('app-only without INTEGRATION_USER: no user for the member')
      return
    }
    const { conn } = scratch()
    const group = conn.web.siteGroups.getByName(
      `${conn.webInfo.Title} Members`.replace(/'/g, "''")
    )
    const logins = async () =>
      (await group.users.select('LoginName')<Array<{ LoginName: string }>>())
        .map((member) => member.LoginName.toLowerCase())
        .sort()
    assertClean(
      await apply(conn, loadFixture('security', { user })),
      'the first run'
    )
    // The group is there (a missing one fails here, not as a skip below).
    assert.deepEqual(await logins(), [user.toLowerCase()])
    // Everyone: a principal every tenant has, so no real person is added to the group.
    const everyone = 'c:0(.s|true'
    try {
      try {
        await group.users.add(everyone)
      } catch (error) {
        // Only SharePoint refusing the principal skips; a timeout or a 5xx fails.
        if (!error || (error.status !== 400 && error.status !== 403))
          throw error
        t.skip(
          `SharePoint refuses Everyone as a member: ${describeError(error)}`
        )
        return
      }
      assert.deepEqual(await logins(), [everyone, user.toLowerCase()].sort())

      assertClean(
        await apply(conn, loadFixture('security', { user })),
        'the second run'
      )
      assert.deepEqual(await logins(), [user.toLowerCase()])
    } finally {
      // The group has Contribute on the web: never leave Everyone in it, also not on a kept web
      // (after the run above it is gone, and this answers an error, which is fine).
      await group.users.removeByLoginName(everyone).catch(() => undefined)
    }
  })

  it('breaks inheritance on a web that inherits, copying the assignments', async (t) => {
    const { conn } = scratch()
    const url = 'inherits'
    const nestedUrl = `${conn.url}/${url}`
    try {
      await conn.web.webs.add(
        `${conn.webInfo.Title} inherits`,
        url,
        '',
        'STS#3',
        conn.webInfo.Language,
        true
      )
      const nested = await connect(nestedUrl, conn.mode)
      const inherits = () =>
        nested.web
          .select('HasUniqueRoleAssignments')<{
            HasUniqueRoleAssignments: boolean
          }>()
          .then((web) => !web.HasUniqueRoleAssignments)
      assert.equal(await inherits(), true)
      // An inheriting web has no associated groups of its own: grant the scratch web's.
      const visitors = (
        await conn.web.associatedVisitorGroup.select('Title')<{
          Title: string
        }>()
      ).Title
      const template = () => ({
        Security: {
          BreakRoleInheritance: true,
          CopyRoleAssignments: true,
          RoleAssignments: [{ Principal: visitors, RoleDefinition: 'Read' }]
        }
      })

      assertClean(await apply(nested, template()), 'the first run')
      assert.equal(await inherits(), false)
      const assignments = await readSecurity(nested, [])
      // Copied from the scratch web: its associated owners still have full control.
      const ownerGroup = (
        await conn.web.associatedOwnerGroup.select('Title')<{ Title: string }>()
      ).Title
      assert.ok(assignments.roleAssignments[ownerGroup])

      assertClean(await apply(nested, template()), 'the second run')
      assert.deepEqual(await readSecurity(nested, []), assignments)
    } finally {
      // A web with a subweb cannot be deleted: this one goes first, by its URL, also when the test
      // stopped before it connected (a web that was never made is no error). A failure here is
      // reported, not thrown: it would hide the test's own, and after() removes the web anyway.
      await Web([conn.web, nestedUrl])
        .delete()
        .catch((error) => {
          if (error && error.status === 404) return
          t.diagnostic(
            `could not delete ${url}: ${describeError(
              error
            )} (after() removes it with the scratch web)`
          )
        })
    }
  })
})
