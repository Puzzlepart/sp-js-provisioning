import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  DefaultHandlerMap,
  DefaultHandlerSort
} from '../../src/handlers/exports'
import { Security } from '../../src/handlers/security'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { LogLevel } from '@pnp/logging'
import { captureLogs, levelOf } from '../helpers/logs'

/**
 * Makes a fake queryable: invokable (resolving with `result()`) and selectable,
 * the way handlers call `x.select(...)()`.
 *
 * @param result - Produces the value the queryable resolves with
 */
function queryable(result: () => any): any {
  const instance: any = async () => result()
  instance.select = () => instance
  return instance
}

interface IFakeGroup {
  Id: number
  Title: string
  users: Array<{ Id: number; LoginName: string }>
}

/**
 * Creates a fake web that keeps site groups in memory and records, in order,
 * every write the `Security` handler makes.
 *
 * @param options - Existing groups, whether the web has unique permissions,
 * logins `ensureUser` should reject, group titles `siteGroups.add` should
 * reject, and whether SharePoint ignores an owner change (as REST's
 * `SetUserAsOwner` does for a group id)
 */
function createWeb(
  options: {
    groups?: IFakeGroup[]
    hasUniqueRoleAssignments?: boolean
    failEnsure?: string[]
    failAdd?: string[]
    ignoreOwner?: boolean
  } = {}
) {
  const groups: IFakeGroup[] = options.groups || []
  const events: string[] = []
  const getByNameCalls: string[] = []
  let nextId = 100
  const findGroup = (id: number) => groups.find((g) => g.Id === id)
  // The associated owner, visitor and member groups, and who owns each group
  // (the signed-in user, 7, until the handler changes it).
  const associatedGroupIds = [3, 4, 5]
  const owners = new Map<number, number>()
  const setOwner = (groupId: number, ownerId: number) => {
    if (!options.ignoreOwner) owners.set(groupId, ownerId)
  }
  const roleDefinitions: { [name: string]: number } = { Lese: 1073741826 }
  const web: any = {
    siteGroups: {
      getByName: (name: string) => {
        getByNameCalls.push(name)
        return queryable(() => {
          const group = groups.find((g) => g.Title === name.replace(/''/g, "'"))
          if (!group)
            throw Object.assign(new Error('Not found'), { status: 404 })
          return { Id: group.Id }
        })
      },
      add: async (properties: any) => {
        if ((options.failAdd || []).includes(properties.Title))
          throw new Error(`Cannot create ${properties.Title}`)
        const group = { Id: nextId++, Title: properties.Title, users: [] }
        groups.push(group)
        events.push(`add ${JSON.stringify(properties)}`)
        return { Id: group.Id, Title: group.Title }
      },
      getById: (id: number) => {
        const users: any = queryable(() => findGroup(id).users)
        users.add = async (loginName: string) => {
          const group = findGroup(id)
          // SharePoint matches login names case-insensitively.
          const isMember = group.users.some(
            (u) => u.LoginName.toLowerCase() === loginName.toLowerCase()
          )
          if (!isMember)
            group.users.push({ Id: nextId++, LoginName: loginName })
          events.push(`member ${id} ${loginName}`)
        }
        users.removeById = async (userId: number) => {
          const group = findGroup(id)
          group.users = group.users.filter((u) => u.Id !== userId)
          events.push(`remove ${id} ${userId}`)
        }
        return {
          users,
          update: async (properties: any) => {
            events.push(`update ${id} ${JSON.stringify(properties)}`)
          },
          setUserAsOwner: async (ownerId: number) => {
            events.push(`owner ${id} ${ownerId}`)
            setOwner(id, ownerId)
          },
          // `getById(id).select('Id')()`: does the group exist?
          select: () =>
            queryable(() => {
              if (!findGroup(id) && !associatedGroupIds.includes(id))
                throw Object.assign(new Error('Group not found'), {
                  status: 404
                })
              return { Id: id }
            }),
          // `getById(id).expand('Owner').select('Owner/Id')()`
          expand: () =>
            queryable(() => ({ Owner: { Id: owners.get(id) ?? 7 } }))
        }
      }
    },
    ensureUser: async (loginName: string) => {
      if ((options.failEnsure || []).includes(loginName))
        throw new Error(`Unknown user ${loginName}`)
      return { Id: 50, LoginName: `i:0#.f|membership|${loginName}` }
    },
    associatedOwnerGroup: queryable(() => ({ Id: 3 })),
    associatedMemberGroup: queryable(() => ({ Id: 5 })),
    associatedVisitorGroup: queryable(() => ({ Id: 4 })),
    roleDefinitions: {
      getByName: (name: string) =>
        queryable(() => {
          if (!(name in roleDefinitions)) throw new Error(`No role ${name}`)
          return { Id: roleDefinitions[name] }
        }),
      getByType: (kind: number) => queryable(() => ({ Id: 9000 + kind }))
    },
    // `web.select(...)` is invoked like a queryable: `web.select(...)()`.
    select: () =>
      queryable(() => ({
        HasUniqueRoleAssignments: !!options.hasUniqueRoleAssignments
      })),
    breakRoleInheritance: async (copy: boolean, clear: boolean) => {
      events.push(`break ${copy} ${clear}`)
    },
    roleAssignments: {
      add: async (principalId: number, roleDefinitionId: number) => {
        events.push(`assign ${principalId} ${roleDefinitionId}`)
      }
    }
  }
  /**
   * The CSOM request that sets a group's owner, replayed against the fake:
   * the first Int32 parameter is the group, the second its new owner.
   *
   * @param _web - The web
   * @param request - The request XML
   */
  const processQuery = async (_web: any, request: string) => {
    const ids: number[] = []
    const pattern = /<Parameter Type="Int32">(\d+)<\/Parameter>/g
    let match: RegExpExecArray | null
    while ((match = pattern.exec(request))) ids.push(Number(match[1]))
    events.push(`csom-owner ${ids[0]} ${ids[1]}`)
    setOwner(ids[0], ids[1])
    return [{ ErrorInfo: null }]
  }
  // Test-only: the CSOM stand-in `securityFor` hands to the handler.
  web.csom = processQuery
  return { web, groups, events, getByNameCalls, owners }
}

/**
 * A `Security` handler whose CSOM requests go to the fake web.
 *
 * @param web - The fake web from `createWeb`
 */
function securityFor(web: any): Security {
  const handler = new Security({})
  handler.processQuery = web.csom
  return handler
}

/**
 * Creates a provisioning context for a web titled `Prosjekt A`.
 */
function createContext(): ProvisioningContext {
  const context = new ProvisioningContext()
  context.web = { Id: 'web-guid', Title: 'Prosjekt A' }
  return context
}

describe('Security: site groups', () => {
  it('creates a missing group with its properties, members and owner', async () => {
    const { web, events } = createWeb()

    await securityFor(web).ProvisionObjects(
      web,
      {
        SiteGroups: [
          {
            Title: '{sitename} Styringsgruppe',
            Description: 'Styringsgruppen for {sitename}',
            AllowMembersEditMembership: false,
            Owner: '{associatedownergroupid}',
            Members: ['ola@contoso.no']
          }
        ]
      },
      createContext()
    )

    // The owner is set last, so the provisioning user still owns the group
    // while its members are added. A group owner goes through CSOM.
    assert.deepEqual(events, [
      `add ${JSON.stringify({
        Title: 'Prosjekt A Styringsgruppe',
        Description: 'Styringsgruppen for Prosjekt A',
        AllowMembersEditMembership: false
      })}`,
      'member 100 i:0#.f|membership|ola@contoso.no',
      'csom-owner 100 3'
    ])
  })

  it('updates an existing group instead of creating a duplicate', async () => {
    const { web, events, getByNameCalls } = createWeb({
      groups: [{ Id: 7, Title: "Prosjektleder's", users: [] }]
    })

    await securityFor(web).ProvisionObjects(
      web,
      {
        SiteGroups: [
          { Title: "Prosjektleder's", Description: 'Oppdatert' },
          // Nothing but the title: the existing group is left untouched.
          { Title: "Prosjektleder's" }
        ]
      },
      createContext()
    )

    assert.deepEqual(events, [
      `update 7 ${JSON.stringify({ Description: 'Oppdatert' })}`
    ])
    // Apostrophes are escaped for the OData string literal in the url.
    assert.deepEqual(getByNameCalls, ["Prosjektleder''s", "Prosjektleder''s"])
  })

  it('can make a new group its own owner', async () => {
    const { web, events } = createWeb()

    await securityFor(web).ProvisionObjects(
      web,
      { SiteGroups: [{ Title: 'Styringsgruppe', Owner: 'Styringsgruppe' }] },
      createContext()
    )

    assert.equal(events[events.length - 1], 'csom-owner 100 100')
  })

  it('makes a user the owner through SetUserAsOwner', async () => {
    const { web, events, owners } = createWeb()

    const { entries } = await captureLogs(() =>
      securityFor(web).ProvisionObjects(
        web,
        { SiteGroups: [{ Title: 'Styringsgruppe', Owner: 'kari@contoso.no' }] },
        createContext()
      )
    )

    // Not a site group, so ensured as a user (id 50).
    assert.equal(events[events.length - 1], 'owner 100 50')
    assert.equal(owners.get(100), 50)
    assert.equal(
      levelOf(entries, /Owner of site group Styringsgruppe set to/),
      LogLevel.Info
    )
  })

  it('takes a bare owner id that is a site group as a group', async () => {
    const { web, events } = createWeb()

    await securityFor(web).ProvisionObjects(
      web,
      { SiteGroups: [{ Title: 'Styringsgruppe', Owner: '5' }] },
      createContext()
    )

    assert.equal(events[events.length - 1], 'csom-owner 100 5')
  })

  it('logs an error when SharePoint does not apply the owner', async () => {
    // What the first tenant run found: REST's SetUserAsOwner answered
    // success for a group id and left the owner unchanged.
    const { web } = createWeb({ ignoreOwner: true })

    const { entries } = await captureLogs(() =>
      securityFor(web).ProvisionObjects(
        web,
        {
          SiteGroups: [
            { Title: 'Styringsgruppe', Owner: '{associatedownergroupid}' }
          ]
        },
        createContext()
      )
    )

    assert.equal(
      levelOf(
        entries,
        /Owner of site group Styringsgruppe is 7 after setting it to .* \(3\): SharePoint did not apply the change/
      ),
      LogLevel.Error
    )
    assert.equal(
      levelOf(entries, /set to \{associatedownergroupid\}/),
      undefined
    )
  })

  it('removes members not in the template with ClearExistingMembers', async () => {
    const { web, groups } = createWeb({
      groups: [
        {
          Id: 7,
          Title: 'Styringsgruppe',
          users: [
            { Id: 1, LoginName: 'i:0#.f|membership|kari@contoso.no' },
            { Id: 2, LoginName: 'I:0#.F|MEMBERSHIP|OLA@CONTOSO.NO' }
          ]
        }
      ]
    })

    await securityFor(web).ProvisionObjects(
      web,
      {
        SiteGroups: [
          {
            Title: 'Styringsgruppe',
            Members: ['ola@contoso.no'],
            ClearExistingMembers: true
          }
        ]
      },
      createContext()
    )

    // Login names are compared case-insensitively.
    assert.deepEqual(
      groups[0].users.map((u) => u.Id),
      [2]
    )
  })

  it('keeps existing members when a member cannot be added', async () => {
    const { web, groups } = createWeb({
      groups: [
        {
          Id: 7,
          Title: 'Styringsgruppe',
          users: [{ Id: 1, LoginName: 'i:0#.f|membership|kari@contoso.no' }]
        }
      ],
      failEnsure: ['ukjent@contoso.no']
    })

    await securityFor(web).ProvisionObjects(
      web,
      {
        SiteGroups: [
          {
            Title: 'Styringsgruppe',
            Members: ['ukjent@contoso.no', 'ola@contoso.no'],
            ClearExistingMembers: true
          }
        ]
      },
      createContext()
    )

    assert.deepEqual(
      groups[0].users.map((u) => u.LoginName),
      ['i:0#.f|membership|kari@contoso.no', 'i:0#.f|membership|ola@contoso.no']
    )
  })

  it('logs a failing group and still provisions the next one', async () => {
    const { web, groups } = createWeb({ failAdd: ['Feiler'] })

    await securityFor(web).ProvisionObjects(
      web,
      { SiteGroups: [{ Title: 'Feiler' }, { Title: 'Prosjektledere' }] },
      createContext()
    )

    assert.deepEqual(
      groups.map((g) => g.Title),
      ['Prosjektledere']
    )
  })
})

describe('Security: web permissions', () => {
  it('breaks inheritance and adds role assignments for groups from the template', async () => {
    const { web, events } = createWeb()

    await securityFor(web).ProvisionObjects(
      web,
      {
        SiteGroups: [{ Title: '{sitename} Styringsgruppe' }],
        BreakRoleInheritance: true,
        CopyRoleAssignments: true,
        RoleAssignments: [
          { Principal: '{sitename} Styringsgruppe', RoleDefinition: 'Lese' },
          // English well-known names fall back to the role type (Reader=2).
          { Principal: '{associatedvisitorgroupid}', RoleDefinition: 'Read' }
        ]
      },
      createContext()
    )

    assert.deepEqual(events.slice(1), [
      'break true false',
      'assign 100 1073741826',
      'assign 4 9002'
    ])
  })

  it('leaves inheritance alone when the web already has unique permissions', async () => {
    // Breaking again with ClearSubscopes would wipe list permissions on every
    // re-run of the template.
    const { web, events } = createWeb({ hasUniqueRoleAssignments: true })

    await securityFor(web).ProvisionObjects(
      web,
      {
        BreakRoleInheritance: true,
        ClearSubscopes: true,
        RoleAssignments: [
          { Principal: '{associatedmembergroupid}', RoleDefinition: 'Lese' }
        ]
      },
      createContext()
    )

    assert.deepEqual(events, ['assign 5 1073741826'])
  })

  it('logs a failing role assignment and still adds the next one', async () => {
    const { web, events } = createWeb()

    await securityFor(web).ProvisionObjects(
      web,
      {
        RoleAssignments: [
          {
            Principal: '{associatedmembergroupid}',
            RoleDefinition: 'Finnes ikke'
          },
          { Principal: '{associatedmembergroupid}', RoleDefinition: 'Lese' }
        ]
      },
      createContext()
    )

    assert.deepEqual(events, ['assign 5 1073741826'])
  })
})

describe('Security: associated groups', () => {
  it('logs a clear error for an associated group the web does not have', async () => {
    // A web that inherits permissions has no associated groups of its own:
    // the id came back undefined and was sent as principalid=undefined.
    const { web, events } = createWeb()
    web.associatedVisitorGroup = queryable(() => ({}))

    const { entries } = await captureLogs(() =>
      securityFor(web).ProvisionObjects(
        web,
        {
          RoleAssignments: [
            { Principal: '{associatedvisitorgroupid}', RoleDefinition: 'Lese' },
            { Principal: '{associatedmembergroupid}', RoleDefinition: 'Lese' }
          ]
        },
        createContext()
      )
    )

    assert.equal(
      levelOf(entries, /The web has no visitor group of its own/),
      LogLevel.Error
    )
    // The next assignment is still added.
    assert.deepEqual(events, ['assign 5 1073741826'])
  })
})

describe('Security: handler registration', () => {
  it('runs before Lists, so list Security can use the groups it creates', () => {
    assert.ok(DefaultHandlerSort.Security < DefaultHandlerSort.Lists)
    assert.ok(DefaultHandlerSort.Security > DefaultHandlerSort.Features)
    assert.ok(DefaultHandlerMap({}).Security instanceof Security)
  })
})
