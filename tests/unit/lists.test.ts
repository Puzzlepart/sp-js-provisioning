import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { Lists } from '../../src/handlers/lists'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { TokenHelper } from '../../src/util/tokenhelper'

/**
 * Creates a `Lists` handler with the two members that the private methods
 * under test read directly, so they can be exercised without going through
 * `ProvisionObjects` (which initializes JSOM).
 *
 * @param context - Provisioning context the handler should write to
 */
function createHandler(context: ProvisioningContext): any {
  const handler: any = new Lists({})
  handler.context = context
  handler.tokenHelper = new TokenHelper(context, {})
  return handler
}

/**
 * Creates a fake list content type collection: invokable (to read the current
 * bindings) with an `addAvailableContentType` that records what was bound.
 *
 * @param bound - Array the bound content type ids are pushed to
 */
function createContentTypes(bound: string[]): any {
  const contentTypes: any = async () => []
  contentTypes.addAvailableContentType = async (contentTypeId: string) => {
    bound.push(contentTypeId)
  }
  return contentTypes
}

describe('Lists: v4 list creation result', () => {
  it('reads Id and Title straight off the lists.add result', async () => {
    // PnPjs v4 resolves `lists.add` with the IListInfo itself - there is no
    // longer a `{ list, data }` wrapper to unpack.
    const context = new ProvisioningContext()
    const handler = createHandler(context)
    const bound: string[] = []
    const createdList = { contentTypes: createContentTypes(bound) }
    const getByIdCalls: string[] = []
    const web: any = {
      lists: {
        add: async (title: string) => ({ Id: 'list-guid', Title: title }),
        getById: (id: string) => {
          getByIdCalls.push(id)
          return createdList
        },
        ensure: async () => {
          throw new Error('ensure must not be called for a new list')
        }
      }
    }

    await handler.processList(web, {
      Title: 'Prosjektstatus',
      ContentTypeBindings: [{ ContentTypeID: '0x0100AA' }]
    })

    assert.deepEqual(getByIdCalls, ['list-guid'])
    assert.equal(context.lists['Prosjektstatus'], 'list-guid')
    assert.deepEqual(bound, ['0x0100AA'])
  })

  it('still uses the list returned by ensure for an existing list', async () => {
    // `lists.ensure` is one of the few v4 calls that kept its result wrapper.
    const context = new ProvisioningContext()
    context.lists['Prosjektstatus'] = 'list-guid'
    const handler = createHandler(context)
    const bound: string[] = []
    const ensuredList = { contentTypes: createContentTypes(bound) }
    const web: any = {
      lists: {
        add: async () => {
          throw new Error('add must not be called for an existing list')
        },
        ensure: async (title: string) => ({
          list: ensuredList,
          created: false,
          data: { Id: 'list-guid', Title: title }
        })
      }
    }

    await handler.processList(web, {
      Title: 'Prosjektstatus',
      ContentTypeBindings: [{ ContentTypeID: '0x0100BB' }]
    })

    assert.deepEqual(bound, ['0x0100BB'])
  })
})

describe('Lists: list lookup paging', () => {
  it('asks for 5000 lists both before and after provisioning', async () => {
    // Without an explicit $top SharePoint returns the first 100 lists only, and
    // {listid:...} tokens for anything beyond that silently stay unresolved.
    const handler: any = new Lists({})
    const context = new ProvisioningContext()
    context.web = { ServerRelativeUrl: '/sites/demo' } as any
    const tops: number[] = []
    const lists: any = async () => []
    lists.select = () => lists
    lists.top = (count: number) => {
      tops.push(count)
      return lists
    }

    await handler.ProvisionObjects({ lists } as any, [], context)

    assert.deepEqual(tops, [5000, 5000])
  })
})

describe('Lists: v4 field creation result', () => {
  // The schema XML id is braced and upper case, the way SharePoint field XML
  // ships it - `fields.getById` adds the braces itself in v4, so the handler
  // has to normalise before using it.
  const schemaXml =
    '<Field ID="{11111111-AAAA-1111-1111-111111111111}" Name="GtTestField" DisplayName="Testfelt" Type="Text" />'
  const normalizedId = '11111111-aaaa-1111-1111-111111111111'

  /**
   * Creates a fake web whose single list records every field delete/update.
   *
   * @param createFieldResult - What `createFieldAsXml` should resolve with
   * @param updates - Array the `{ id, props }` of each update is pushed to
   * @param deleted - Array the id of each deleted field is pushed to
   */
  function createWeb(
    createFieldResult: any,
    updates: any[],
    deleted: string[]
  ): any {
    const fields: any = {
      getById: (id: string) => ({
        delete: async () => {
          deleted.push(id)
        },
        update: async (props: any) => {
          updates.push({ id, props })
        }
      }),
      createFieldAsXml: async () => createFieldResult
    }
    return { lists: { getByTitle: () => ({ fields }) } }
  }

  it('updates the field id returned by createFieldAsXml', async () => {
    // v4 resolves `createFieldAsXml` with Partial<IFieldInfo>, so the field has
    // to be re-resolved by id instead of the removed `result.field`.
    const updates: any[] = []
    const deleted: string[] = []
    const handler = createHandler(new ProvisioningContext())
    const web = createWeb({ Id: 'created-field-id' }, updates, deleted)

    await handler.processField(web, { Title: 'Prosjektstatus' }, schemaXml)

    // A braced id here would produce a double-braced `getById('{{guid}}')` url.
    assert.deepEqual(deleted, [normalizedId])
    assert.deepEqual(updates, [
      { id: 'created-field-id', props: { Title: 'Testfelt' } }
    ])
  })

  it('falls back to the schema xml id when the response carries none', async () => {
    const updates: any[] = []
    const deleted: string[] = []
    const handler = createHandler(new ProvisioningContext())
    const web = createWeb({}, updates, deleted)

    await handler.processField(web, { Title: 'Prosjektstatus' }, schemaXml)

    assert.deepEqual(updates, [
      { id: normalizedId, props: { Title: 'Testfelt' } }
    ])
    assert.deepEqual(deleted, [normalizedId])
  })
})

describe('Lists: v4 field ref creation result', () => {
  it('normalises the braced field ref id it falls back to', async () => {
    // Same double-brace trap as `processField`: the id comes from the template
    // and may be braced, while v4's `getById` braces the value itself.
    const context = new ProvisioningContext()
    context.lists['Prosjektstatus'] = 'list-guid'
    const handler = createHandler(context)
    const updates: any[] = []
    const list: any = {
      fields: {
        getById: (id: string) => ({
          update: async (props: any) => {
            updates.push({ id, props })
          }
        }),
        createFieldAsXml: async () => ({})
      }
    }
    const fieldReference: any = {
      ID: '{22222222-BBBB-2222-2222-222222222222}',
      Name: 'GtTestField',
      DisplayName: 'Testfelt',
      Required: false,
      Hidden: false
    }
    const webFields: any[] = [
      {
        Id: '{22222222-BBBB-2222-2222-222222222222}',
        InternalName: 'GtTestField',
        SchemaXml:
          '<Field ID="{22222222-BBBB-2222-2222-222222222222}" Name="GtTestField" Type="Text" />'
      }
    ]

    await handler.processFieldRef(
      list,
      { Title: 'Prosjektstatus' },
      fieldReference,
      [],
      webFields
    )

    assert.deepEqual(
      updates.map((update) => update.id),
      ['22222222-bbbb-2222-2222-222222222222']
    )
  })
})

describe('Lists: v4 ensureUser result', () => {
  /**
   * Creates a fake invokable that always rejects, standing in for a site group
   * lookup that finds nothing.
   */
  function createMissingGroup(): any {
    const queryable: any = async () => {
      throw new Error('group not found')
    }
    queryable.select = () => queryable
    return queryable
  }

  it('reads Id off the ensureUser result when resolving a principal', async () => {
    // v4 resolves `ensureUser` with ISiteUserInfo - the v3 `{ data }` wrapper
    // is gone.
    const handler = createHandler(new ProvisioningContext())
    let ensuredLogin: string
    const web: any = {
      siteGroups: { getByName: () => createMissingGroup() },
      ensureUser: async (loginName: string) => {
        ensuredLogin = loginName
        return { Id: 42, LoginName: loginName }
      }
    }

    const principalId = await handler._resolvePrincipalId(
      web,
      'i:0#.f|membership|kari@contoso.no'
    )

    assert.equal(principalId, 42)
    assert.equal(ensuredLogin, 'i:0#.f|membership|kari@contoso.no')
  })

  it('reads Id off the ensureUser result when resolving a user column', async () => {
    const handler = createHandler(new ProvisioningContext())
    const web: any = { ensureUser: async () => ({ Id: 7 }) }

    assert.equal(await handler._resolveUserId(web, 'kari@contoso.no'), 7)
  })

  it('does not unwrap a v3-shaped ensureUser result', async () => {
    // Guards against reintroducing `result.data.Id`: a v3-shaped response has
    // no usable id under v4 and must not silently resolve to one.
    const handler = createHandler(new ProvisioningContext())
    const web: any = { ensureUser: async () => ({ data: { Id: 7 } }) }

    assert.equal(
      await handler._resolveUserId(web, 'kari@contoso.no'),
      undefined
    )
  })
})
