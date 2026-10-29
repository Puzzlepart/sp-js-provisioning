import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { Lists } from '../../src/handlers/lists'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { resolvePrincipalId } from '../../src/util/security'
import { TokenHelper } from '../../src/util/tokenhelper'
import { LogLevel } from '@pnp/logging'
import { ClientSidePages } from '../../src/handlers/clientsidepages'
import { captureLogs, levelOf } from '../helpers/logs'

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

/**
 * A fake PnPjs queryable: calling it resolves `value` (or rejects with it, when it is an Error),
 * and `select`, `filter`, `top` and `expand` return itself.
 *
 * @param value - What a call resolves with
 * @param calls - Where each call's filter is recorded, when given
 */
function queryable(value: any, calls?: string[]): any {
  const instance: any = async () => {
    if (calls) calls.push(instance.lastFilter)
    if (value instanceof Error) throw value
    const result =
      typeof value === 'function' ? value(instance.lastFilter) : value
    return Array.isArray(result) && instance.limit
      ? result.slice(0, instance.limit)
      : result
  }
  instance.select = () => instance
  instance.top = (count: number) => {
    instance.limit = count
    return instance
  }
  instance.expand = () => instance
  instance.filter = (filter: string) => {
    instance.lastFilter = filter
    return instance
  }
  return instance
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

  it('marks a list it creates as made in this setup', async () => {
    const context = new ProvisioningContext()
    const handler = createHandler(context)
    const web: any = {
      lists: {
        add: async (title: string) => ({ Id: 'list-guid', Title: title }),
        getById: () => ({ contentTypes: createContentTypes([]) })
      }
    }

    await handler.processList(web, { Title: 'Risikoer' })

    assert.deepEqual(context.createdLists, { Risikoer: true })
  })

  it('updates an existing list with only what the template defines', async () => {
    // `lists.ensure` sent PnPjs' defaults for what the template leaves out (an empty Description,
    // content types off): every hub import cleared the descriptions of the hub lists.
    const context = new ProvisioningContext()
    context.lists['Prosjektstatus'] = 'list-guid'
    const handler = createHandler(context)
    const bound: string[] = []
    const updates: any[] = []
    const ids: string[] = []
    const web: any = {
      lists: {
        add: async () => {
          throw new Error('add must not be called for an existing list')
        },
        ensure: async () => {
          throw new Error('ensure must not be called: it sends defaults')
        },
        getById: (id: string) => {
          ids.push(id)
          return {
            contentTypes: createContentTypes(bound),
            update: async (properties: any) => {
              updates.push(properties)
            }
          }
        }
      }
    }

    await handler.processList(web, {
      Title: 'Prosjektstatus',
      Template: 100,
      AdditionalSettings: { EnableVersioning: true },
      ContentTypeBindings: [{ ContentTypeID: '0x0100BB' }]
    })

    assert.deepEqual(ids, ['list-guid'])
    assert.deepEqual(updates, [{ EnableVersioning: true }])
    assert.deepEqual(bound, ['0x0100BB'])
    assert.deepEqual(context.createdLists, {})
  })
})

describe('Lists: a list created before SiteFields', () => {
  it('finds it existing, sends its settings no second time, and still counts it as made in this setup', async () => {
    // WebProvisioner created it with the template's settings because a site column's {listid:}
    // names it (docs/specs/sitefields-re-apply.md).
    const context = new ProvisioningContext()
    context.lists['Kilder'] = 'list-guid'
    context.createdLists['Kilder'] = true
    context.preCreatedLists['Kilder'] = true
    const handler = createHandler(context)
    const updates: any[] = []
    const web: any = {
      lists: {
        add: async () => {
          throw new Error(
            'add must not be called for a list created before SiteFields'
          )
        },
        getById: () => ({
          contentTypes: createContentTypes([]),
          update: async (properties: any) => {
            updates.push(properties)
          }
        })
      }
    }

    await handler.processList(web, {
      Title: 'Kilder',
      Description: 'Kildeliste',
      AdditionalSettings: { EnableVersioning: true }
    })

    assert.deepEqual(updates, [])
    assert.deepEqual(context.createdLists, { Kilder: true })
    assert.deepEqual(context.preCreatedLists, {})
  })

  it('creates a list with its settings through createList, recording it in the context it is given', async () => {
    const context = new ProvisioningContext()
    const added: any[] = []
    const web: any = {
      lists: {
        add: async (...parameters: any[]) => {
          added.push(parameters)
          return { Id: 'list-guid', Title: parameters[0] }
        },
        getById: (id: string) => ({ id })
      }
    }

    const list = await new Lists({}).createList(
      web,
      {
        Title: 'Kilder',
        Description: 'Kildeliste',
        Template: 100,
        ContentTypesEnabled: true,
        AdditionalSettings: { EnableVersioning: true }
      },
      context
    )

    assert.deepEqual(added, [
      ['Kilder', 'Kildeliste', 100, true, { EnableVersioning: true }]
    ])
    assert.deepEqual(list, { id: 'list-guid' })
    assert.deepEqual(context.lists, { Kilder: 'list-guid' })
    assert.deepEqual(context.createdLists, { Kilder: true })
  })
})

describe('Lists: lists made in this setup', () => {
  it('keeps those an earlier applyTemplate made', async () => {
    const handler: any = new Lists({})
    const context = new ProvisioningContext()
    context.web = { ServerRelativeUrl: '/sites/demo' } as any
    context.createdLists['Fasesjekkliste'] = true
    const lists: any = async () => []
    lists.select = () => lists
    lists.top = () => lists

    await handler.ProvisionObjects({ lists } as any, [], context)

    assert.deepEqual(context.createdLists, { Fasesjekkliste: true })
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

describe('Lists: list fields', () => {
  const id = '11111111-aaaa-1111-1111-111111111111'
  const templateXml = (extra = '') =>
    `<Field ID="{11111111-AAAA-1111-1111-111111111111}" Name="GtCode" DisplayName="Kode" Type="Text"${extra} />`
  const storedXml = (extra = '') =>
    `<Field ID="{${id}}" Name="GtCode" StaticName="GtCode" DisplayName="Kode" Type="Text" ColName="nvarchar7" RowOrdinal="0" Version="2"${extra} />`

  /**
   * A fake list that records what is sent to its fields.
   *
   * @param created - What `createFieldAsXml` resolves with
   */
  function createList(created: any = { Id: 'created-field-id' }) {
    const sent: any[] = []
    const list: any = {
      fields: {
        getById: (fieldId: string) => ({
          delete: async () => {
            sent.push({ delete: fieldId })
          },
          update: async (properties: any, type?: string) => {
            sent.push({ update: fieldId, properties, type })
          }
        }),
        createFieldAsXml: async (xml: string) => {
          sent.push({ create: xml })
          if (created instanceof Error) throw created
          return created
        }
      }
    }
    return { list, sent }
  }
  const lc: any = { Title: 'Risikoer' }

  it('creates a missing field, its display name set after, and deletes nothing', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList()

    await handler.processField(list, lc, templateXml(), [])

    assert.equal(sent.length, 2)
    // Created with the internal name as display name, so it becomes the internal name.
    assert.match(sent[0].create, /DisplayName="GtCode"/)
    assert.deepEqual(sent[1], {
      update: 'created-field-id',
      properties: { Title: 'Kode' },
      type: 'SP.Field'
    })
  })

  it('falls back to the schema xml id, normalised, when the response carries none', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList({})

    await handler.processField(list, lc, templateXml(), [])

    // A braced id here would produce a double-braced `getById('{{guid}}')` url.
    assert.equal(sent[1].update, id)
  })

  it('leaves an unchanged field alone on a second apply', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList()

    await handler.processField(list, lc, templateXml(), [
      {
        Id: id,
        InternalName: 'GtCode',
        TypeAsString: 'Text',
        SchemaXml: storedXml()
      }
    ])

    assert.deepEqual(sent, [])
  })

  it('updates a changed field in place: one MERGE of the merged schema xml, nothing deleted', async () => {
    // Deleting and recreating it lost the values in the column on every apply.
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList()

    await handler.processField(list, lc, templateXml(' MaxLength="100"'), [
      {
        Id: id,
        InternalName: 'GtCode',
        TypeAsString: 'Text',
        SchemaXml: storedXml()
      }
    ])

    assert.equal(sent.length, 1)
    assert.equal(sent[0].update, id)
    assert.equal(sent[0].type, 'SP.Field')
    const schemaXml: string = sent[0].properties.SchemaXml
    assert.match(schemaXml, /MaxLength="100"/)
    assert.match(schemaXml, /ColName="nvarchar7"/)
    assert.doesNotMatch(schemaXml, /Version=/)
  })

  it('applies a widening type change and warns about any other, changing nothing', async () => {
    const handler = createHandler(new ProvisioningContext())
    const widening = createList()
    const narrowing = createList()
    const stored = (type: string) => [
      {
        Id: id,
        InternalName: 'GtCode',
        TypeAsString: type,
        SchemaXml: storedXml().replace('Type="Text"', `Type="${type}"`)
      }
    ]

    await handler.processField(
      widening.list,
      lc,
      templateXml().replace('Type="Text"', 'Type="Note"'),
      stored('Text')
    )
    const { entries } = await captureLogs(() =>
      handler.processField(narrowing.list, lc, templateXml(), stored('Note'))
    )

    assert.match(widening.sent[0].properties.SchemaXml, /Type="Note"/)
    assert.deepEqual(narrowing.sent, [])
    assert.equal(
      levelOf(
        entries,
        /is Note, the template says Text: SharePoint would lose values/
      ),
      LogLevel.Warning
    )
  })

  it('keeps the target of a lookup the template points elsewhere, with a warning', async () => {
    // SharePoint accepts a new List and keeps the stored item ids, which then name other items.
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList()
    const lookup = (listId: string, extra = '') =>
      `<Field ID="{${id}}" Name="GtLink" DisplayName="Lenke" Type="Lookup" List="{${listId}}" ShowField="Title"${extra} />`

    const { entries } = await captureLogs(() =>
      handler.processField(
        list,
        lc,
        lookup('22222222-2222-4222-8222-222222222222', ' Description="Ny"'),
        [
          {
            Id: id,
            InternalName: 'GtLink',
            TypeAsString: 'Lookup',
            SchemaXml: lookup('11111111-1111-4111-8111-111111111111')
          }
        ]
      )
    )

    assert.equal(
      levelOf(entries, /points at another list or term set/),
      LogLevel.Warning
    )
    const schemaXml: string = sent[0].properties.SchemaXml
    assert.match(schemaXml, /List="\{11111111-1111-4111-8111-111111111111\}"/)
    assert.match(schemaXml, /Description="Ny"/)
  })

  it('points a field at its new target on a list made in this setup, which holds no values yet', async () => {
    // A content type brings a site column with another term set or list than the template's list
    // field (the vei hub's Fasesjekkliste); on a new list the template wins.
    const context = new ProvisioningContext()
    context.createdLists['Risikoer'] = true
    const handler = createHandler(context)
    const { list, sent } = createList()
    const lookup = (listId: string) =>
      `<Field ID="{${id}}" Name="GtLink" DisplayName="Lenke" Type="Lookup" List="{${listId}}" ShowField="Title" />`

    const { entries } = await captureLogs(() =>
      handler.processField(
        list,
        lc,
        lookup('22222222-2222-4222-8222-222222222222'),
        [
          {
            Id: id,
            InternalName: 'GtLink',
            TypeAsString: 'Lookup',
            SchemaXml: lookup('11111111-1111-4111-8111-111111111111')
          }
        ]
      )
    )

    assert.equal(
      levelOf(entries, /points at another list or term set/),
      LogLevel.Info
    )
    assert.match(
      sent[0].properties.SchemaXml,
      /List="\{22222222-2222-4222-8222-222222222222\}"/
    )
  })

  it('sees a template written Type="Lookup" Mult="TRUE" as the LookupMulti SharePoint stores, and writes nothing', async () => {
    // Merged as written, Type="Lookup" would be set on a column with several values in its items.
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList()
    const lookup = (type: string) =>
      `<Field ID="{${id}}" Name="GtLinks" DisplayName="Lenker" Type="${type}" Mult="TRUE" List="{11111111-1111-4111-8111-111111111111}" ShowField="Title" />`

    await handler.processField(list, lc, lookup('Lookup'), [
      {
        Id: id,
        InternalName: 'GtLinks',
        TypeAsString: 'LookupMulti',
        SchemaXml: lookup('LookupMulti')
      }
    ])

    assert.deepEqual(sent, [])
  })

  it('creates a primary lookup before the dependent the template lists first', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList()
    const web: any = {
      lists: {
        getByTitle: () => ({
          ...list,
          fields: Object.assign(queryable([]), list.fields)
        })
      }
    }
    const primary = `<Field ID="{a1a1a1a1-0000-4000-8000-000000000001}" Name="GtGain" DisplayName="Gevinst" Type="Lookup" List="{x}" ShowField="Title" />`
    const dependent = `<Field ID="{a1a1a1a1-0000-4000-8000-000000000002}" Name="GtGain_ID" DisplayName="Gevinst:ID" Type="Lookup" List="{x}" ShowField="ID" FieldRef="{a1a1a1a1-0000-4000-8000-000000000001}" ReadOnly="TRUE" />`

    await handler.processListFields(web, {
      Title: 'Risikoer',
      Fields: [dependent, primary]
    })

    assert.deepEqual(
      sent
        .filter((call) => call.create)
        .map((call) => / Name="([^"]+)"/.exec(call.create)[1]),
      ['GtGain', 'GtGain_ID']
    )
  })

  it('does not create a field whose internal name the list has under another id', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList()

    const { entries } = await captureLogs(() =>
      handler.processField(list, lc, templateXml(), [
        {
          Id: 'other-id',
          InternalName: 'GtCode',
          TypeAsString: 'Text',
          SchemaXml: '<Field />'
        }
      ])
    )

    assert.deepEqual(sent, [])
    assert.equal(
      levelOf(entries, /with another id \(other-id\)/),
      LogLevel.Warning
    )
  })

  it('writes nothing for a list whose fields cannot be read, and says so', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { list, sent } = createList()
    const web: any = {
      lists: {
        getByTitle: () => ({
          ...list,
          fields: Object.assign(queryable(new Error('throttled')), list.fields)
        })
      }
    }

    const { entries } = await captureLogs(() =>
      handler.processListFields(web, {
        Title: 'Risikoer',
        Fields: [templateXml()]
      })
    )

    assert.deepEqual(sent, [])
    assert.equal(
      levelOf(entries, /Failed to read the fields of list Risikoer/),
      LogLevel.Error
    )
  })

  it('makes a primary lookup before the dependent lookup that refers to it', async () => {
    const primary = `<Field ID="{a1a1a1a1-0000-4000-8000-000000000001}" Name="GtGain" DisplayName="Gevinst" Type="Lookup" List="{x}" ShowField="Title" />`
    const dependent = `<Field ID="{a1a1a1a1-0000-4000-8000-000000000002}" Name="GtGain_ID" DisplayName="Gevinst:ID" Type="Lookup" List="{x}" ShowField="ID" FieldRef="{A1A1A1A1-0000-4000-8000-000000000001}" ReadOnly="TRUE" />`
    const other = templateXml()

    assert.deepEqual(
      (Lists as any)._primariesFirst([dependent, other, primary]),
      [primary, dependent, other]
    )
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

describe('Lists: a field ref to a column that exists nowhere', () => {
  it('is an error naming the field ref and the list, and nothing is sent', async () => {
    // A site column that failed or was skipped: the field ref used to do nothing, silently.
    const handler = createHandler(new ProvisioningContext())
    const sent: any[] = []
    const list: any = {
      fields: {
        getById: (id: string) => ({
          update: async () => {
            sent.push(id)
          }
        }),
        createFieldAsXml: async (xml: string) => {
          sent.push(xml)
        }
      }
    }

    const { entries } = await captureLogs(() =>
      handler.processFieldRef(
        list,
        { Title: 'Risikoer' },
        { ID: '{33333333-cccc-3333-3333-333333333333}', Name: 'GtMissing' },
        [],
        []
      )
    )

    assert.deepEqual(sent, [])
    assert.equal(
      levelOf(
        entries,
        /Field ref '\{33333333-cccc-3333-3333-333333333333\}' \(GtMissing\) for list Risikoer: no column with that id on the list or the web/
      ),
      LogLevel.Error
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
    let ensuredLogin: string
    const web: any = {
      siteGroups: { getByName: () => createMissingGroup() },
      ensureUser: async (loginName: string) => {
        ensuredLogin = loginName
        return { Id: 42, LoginName: loginName }
      }
    }

    const principalId = await resolvePrincipalId(
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

describe('Lists: security principals', () => {
  it('replaces tokens in a principal before resolving it', async () => {
    // Groups created by the `Security` handler are often named per site, so
    // list role assignments must resolve the same tokens.
    const context = new ProvisioningContext()
    context.web = { Title: 'Prosjekt A' }
    const handler: any = new Lists({ parameters: { Rolle: 'Lese' } })
    handler.context = context
    handler.tokenHelper = new TokenHelper(context, handler.config)
    const groupNames: string[] = []
    const assignments: number[][] = []
    const select = (id: number) => {
      const queryable: any = async () => ({ Id: id })
      queryable.select = () => queryable
      return queryable
    }
    const web: any = {
      lists: {
        getByTitle: () => ({
          breakRoleInheritance: async () => undefined,
          roleAssignments: {
            add: async (principalId: number, roleDefinitionId: number) => {
              assignments.push([principalId, roleDefinitionId])
            }
          }
        })
      },
      siteGroups: {
        getByName: (name: string) => {
          groupNames.push(name)
          return select(12)
        }
      },
      roleDefinitions: { getByName: () => select(1073741826) }
    }

    await handler.processListSecurity(web, {
      Title: 'Risikoer',
      Security: {
        BreakRoleInheritance: true,
        RoleAssignments: [
          {
            Principal: '{sitename} Styringsgruppe',
            RoleDefinition: '{parameter:Rolle}'
          }
        ]
      }
    })

    assert.deepEqual(groupNames, ['Prosjekt A Styringsgruppe'])
    assert.deepEqual(assignments, [[12, 1073741826]])
  })
})

describe('Lists: failure log levels', () => {
  // A failure the handler logs and carries on after must be logged as an
  // error, or `--strict` and the tests' zero-error assertion cannot see it.
  const failing = async () => {
    throw new Error('SharePoint said no')
  }
  /**
   * A fake queryable that rejects, with `select` returning itself.
   */
  const rejecting = (): any => {
    const instance: any = failing
    instance.select = () => instance
    return instance
  }
  const lc: any = { Title: 'Risikoer' }

  it('logs a content type binding that fails as an error', async () => {
    const handler = createHandler(new ProvisioningContext())
    const list: any = { contentTypes: failing }

    const { entries } = await captureLogs(() =>
      handler.processContentTypeBinding(lc, list, '0x0100AA')
    )

    assert.equal(
      levelOf(entries, /Failed to add Content Type 0x0100AA/),
      LogLevel.Error
    )
  })

  it('logs a content type removal SharePoint refuses as in use (500) as a warning, any other as an error', async () => {
    const handler = createHandler(new ProvisioningContext())
    const removal = async (error: Error) => {
      const contentTypes: any = async () => [
        { Id: { StringValue: '0x0100BB' } }
      ]
      contentTypes.getById = () => ({
        delete: async () => {
          throw error
        }
      })
      const { entries } = await captureLogs(() =>
        handler.processContentTypeBindings(lc, { contentTypes }, [], true)
      )
      return levelOf(entries, /Failed to remove content type 0x0100BB/)
    }

    assert.equal(
      await removal(Object.assign(new Error('in use'), { status: 500 })),
      LogLevel.Warning
    )
    assert.equal(
      await removal(Object.assign(new Error('Forbidden'), { status: 403 })),
      LogLevel.Error
    )
  })

  it('logs a list field that cannot be added as an error', async () => {
    const handler = createHandler(new ProvisioningContext())
    const list: any = { fields: { createFieldAsXml: failing } }

    const { entries } = await captureLogs(() =>
      handler.processField(
        list,
        lc,
        '<Field ID="{11111111-AAAA-1111-1111-111111111111}" Name="GtLookup" DisplayName="Oppslag" Type="Lookup" List="{listid:Missing}" />',
        []
      )
    )

    assert.equal(
      levelOf(
        entries,
        /Failed to add field 'Oppslag' to list Risikoer: SharePoint said no/
      ),
      LogLevel.Error
    )
  })

  it("logs a view's fields that fail as an error", async () => {
    const handler = createHandler(new ProvisioningContext())
    // The view's fields differ from the template's, so the handler gets to removeAll.
    const view: any = {
      fields: Object.assign(async () => ({ Items: ['LinkTitle'] }), {
        removeAll: failing
      })
    }

    const { entries } = await captureLogs(() =>
      handler.processViewFields(view, { Title: 'Alle', ViewFields: ['Title'] })
    )

    assert.equal(
      levelOf(
        entries,
        /Failed to process view fields for view Alle: SharePoint said no/
      ),
      LogLevel.Error
    )
  })

  it('logs a root folder or a folder that fails as an error', async () => {
    const handler = createHandler(new ProvisioningContext())
    const web: any = {
      lists: { getByTitle: () => ({ rootFolder: rejecting() }) }
    }
    const parent: any = {
      folders: rejecting(),
      addSubFolderUsingPath: failing
    }

    const root = await captureLogs(() =>
      handler.processListFolders(web, {
        Title: 'Dokumenter',
        Folders: [{ Name: 'A' }]
      })
    )
    const folder = await captureLogs(() =>
      handler._processFolderLevel(
        parent,
        '/sites/x/Dokumenter',
        [{ Name: 'A' }],
        'Dokumenter'
      )
    )

    assert.equal(
      levelOf(root.entries, /Could not resolve root folder/),
      LogLevel.Error
    )
    assert.equal(
      levelOf(folder.entries, /Failed to provision folder/),
      LogLevel.Error
    )
  })

  it('logs a data row that fails as an error', async () => {
    const handler = createHandler(new ProvisioningContext())
    const fields: any = async () => []
    fields.select = () => fields
    const web: any = {
      lists: { getByTitle: () => ({ fields, items: { add: failing } }) }
    }

    const { entries } = await captureLogs(() =>
      handler.processListDataRows(web, {
        Title: 'Risikoer',
        DataRows: { Rows: [{ Title: 'En risiko' }] }
      })
    )

    assert.equal(
      levelOf(entries, /Failed to provision row 0 in list Risikoer/),
      LogLevel.Error
    )
  })
})

describe('Lists: field ref additional properties', () => {
  it('sets them on the schema xml read after the title update, keeping the title', async () => {
    // It used to send DisplayName = the internal name, drop attributes that were undefined, and
    // send the schema xml read before the title update.
    const context = new ProvisioningContext()
    const handler = createHandler(context)
    const id = '22222222-bbbb-2222-2222-222222222222'
    const sent: any[] = []
    const field = (fieldId: string): any => ({
      update: async (properties: any, type: string) => {
        sent.push({ fieldId, properties, type })
      },
      select: () => async () => ({
        SchemaXml: `<Field ID="{${id}}" Name="GtStatus" DisplayName="Status" Type="Text" Hidden="FALSE" Version="5" />`
      })
    })
    const list: any = { fields: { getById: field } }

    await handler.processFieldRef(
      list,
      { Title: 'Risikoer' },
      {
        ID: '{22222222-BBBB-2222-2222-222222222222}',
        Name: 'GtStatus',
        DisplayName: 'Status',
        AdditionalProperties: { ShowInNewForm: 'FALSE' }
      },
      [
        { Id: id, InternalName: 'GtStatus', SchemaXml: '<Field Version="4" />' }
      ],
      []
    )

    assert.equal(sent.length, 2)
    assert.deepEqual(sent[0].properties, {
      Hidden: undefined,
      Required: undefined,
      Title: 'Status'
    })
    const schemaXml: string = sent[1].properties.SchemaXml
    assert.match(schemaXml, /ShowInNewForm="FALSE"/)
    assert.match(schemaXml, /DisplayName="Status"/)
    assert.match(schemaXml, /Hidden="FALSE"/)
    assert.doesNotMatch(schemaXml, /Version=/)
    assert.equal(sent[1].type, 'SP.Field')
  })
})

describe('Lists: content types on lists made in this setup', () => {
  /**
   * A fake list with the given content types, recording what is removed.
   *
   * @param ids - The list content type ids
   */
  function createList(ids: string[]) {
    const removed: string[] = []
    const contentTypes: any = async () =>
      ids.map((id) => ({ Id: { StringValue: id } }))
    contentTypes.getById = (id: string) => ({
      delete: async () => {
        removed.push(id)
      }
    })
    contentTypes.addAvailableContentType = async () => undefined
    return { list: { contentTypes } as any, removed }
  }
  const ids = [
    '0x0100AAAA00112233',
    '0x0100BBBB00445566',
    '0x01',
    '0x012000ABCDEF'
  ]

  it('removes the unbound ones (prefix match, folders kept) on a list it created', async () => {
    const context = new ProvisioningContext()
    context.createdLists['Risikoer'] = true
    const handler = createHandler(context)
    const { list, removed } = createList(ids)

    await handler.processContentTypeBindings(
      { Title: 'Risikoer' },
      list,
      [{ ContentTypeID: '0x0100aaaa' }],
      true
    )

    assert.deepEqual(removed, ['0x0100BBBB00445566', '0x01'])
  })

  it('leaves the content types of a list it did not create as they are', async () => {
    const context = new ProvisioningContext()
    context.lists['Risikoer'] = 'list-guid'
    const handler = createHandler(context)
    const { list, removed } = createList(ids)
    const web: any = {
      lists: { getById: () => ({ ...list, update: async () => undefined }) }
    }

    const { entries } = await captureLogs(() =>
      handler.processList(web, {
        Title: 'Risikoer',
        RemoveExistingContentTypes: true,
        ContentTypeBindings: [{ ContentTypeID: '0x0100AAAA' }]
      })
    )

    assert.deepEqual(removed, [])
    assert.equal(
      levelOf(
        entries,
        /was not made in this setup: RemoveExistingContentTypes/
      ),
      LogLevel.Info
    )
  })
})

describe('Lists: views', () => {
  const views = [
    { Id: 'v1', Title: 'Alle elementer', DefaultView: true },
    { Id: 'v2', Title: 'Mine', PersonalView: true },
    {
      Id: 'v3',
      Title: '',
      Hidden: true,
      ServerRelativeUrl: '/sites/x/SitePages/Hjem.aspx'
    },
    {
      Id: 'v4',
      Title: 'Skjult i listen',
      Hidden: true,
      ServerRelativeUrl: '/sites/x/Lists/Risikoer/Skjult.aspx'
    },
    {
      Id: 'v5',
      Title: 'Brukerens',
      ServerRelativeUrl: '/sites/x/Lists/Risikoer/Brukerens.aspx'
    },
    {
      Id: 'v6',
      Title: 'Malens',
      ServerRelativeUrl: '/sites/x/Lists/Risikoer/Malens.aspx'
    }
  ]

  /**
   * A fake web whose list has the views above, recording deletes; `refuse` makes one refuse.
   */
  function createWeb(refuse?: string) {
    const deleted: string[] = []
    const web: any = {
      lists: {
        getByTitle: () => ({
          views: Object.assign(queryable(views), {
            getById: (id: string) => ({
              delete: async () => {
                if (id === refuse) throw new Error('refused')
                deleted.push(id)
              }
            })
          }),
          rootFolder: queryable({
            ServerRelativeUrl: '/sites/x/Lists/Risikoer'
          })
        })
      }
    }
    return { web, deleted }
  }
  const lc: any = {
    Title: 'Risikoer',
    RemoveExistingViews: true,
    Views: [{ Title: 'Malens' }]
  }

  it("removes, on a list it created, only views the template doesn't name and that are not personal, default or a page's", async () => {
    const context = new ProvisioningContext()
    context.createdLists['Risikoer'] = true
    const handler = createHandler(context)
    const { web, deleted } = createWeb()

    await handler.removeOtherViews(web, lc)

    assert.deepEqual(deleted, ['v4', 'v5'])
  })

  it('carries on past a view it may not remove, with a warning', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { web, deleted } = createWeb('v4')

    const { entries } = await captureLogs(() =>
      handler.removeOtherViews(web, lc)
    )

    assert.deepEqual(deleted, ['v5'])
    assert.equal(
      levelOf(entries, /Failed to remove view Skjult i listen/),
      LogLevel.Warning
    )
  })

  it('removes no view of a list it did not create', async () => {
    const handler = createHandler(new ProvisioningContext())
    handler.processView = async () => undefined
    const { web, deleted } = createWeb()

    await handler.processListViews(web, lc)

    assert.deepEqual(deleted, [])
  })

  it('sends no settings update for a view without AdditionalSettings', async () => {
    // An empty MERGE body is refused (400), and the view's fields were then not set either.
    const handler = createHandler(new ProvisioningContext())
    const updates: any[] = []
    const view: any = Object.assign(async () => ({ Title: 'Alle' }), {
      update: async (settings: any) => {
        updates.push(settings)
      },
      fields: Object.assign(async () => ({ Items: ['LinkTitle'] }), {
        removeAll: async () => undefined,
        add: async () => undefined
      })
    })
    const web: any = {
      lists: { getByTitle: () => ({ views: { getByTitle: () => view } }) }
    }

    await handler.processView(web, lc, {
      Title: 'Alle',
      ViewFields: ['LinkTitle']
    })
    await handler.processView(web, lc, {
      Title: 'Alle',
      ViewFields: ['LinkTitle'],
      AdditionalSettings: { RowLimit: 30 }
    })

    assert.deepEqual(updates, [{ RowLimit: 30 }])
  })

  it("leaves a view's fields alone when they are the template's, and adds the rest past one that fails", async () => {
    const handler = createHandler(new ProvisioningContext())
    const calls: string[] = []
    const view = (items: string[]): any => ({
      fields: Object.assign(async () => ({ Items: items }), {
        removeAll: async () => {
          calls.push('removeAll')
        },
        add: async (name: string) => {
          if (name === 'Mangler') throw new Error('no such field')
          calls.push(`add ${name}`)
        }
      })
    })

    await handler.processViewFields(view(['LinkTitle', 'GtStatus']), {
      Title: 'Alle',
      ViewFields: ['LinkTitle', 'GtStatus']
    })
    assert.deepEqual(calls, [])

    const { entries } = await captureLogs(() =>
      handler.processViewFields(view(['LinkTitle']), {
        Title: 'Alle',
        ViewFields: ['LinkTitle', 'Mangler', 'GtStatus']
      })
    )
    assert.deepEqual(calls, ['removeAll', 'add LinkTitle', 'add GtStatus'])
    assert.equal(
      levelOf(entries, /Failed to add field Mangler to view Alle/),
      LogLevel.Error
    )
  })
})

describe('Lists: list permissions', () => {
  /**
   * A fake web whose list has unique permissions or not (an Error: the read fails).
   */
  function createWeb(unique: boolean | Error) {
    const calls: string[] = []
    const web: any = {
      lists: {
        getByTitle: () => ({
          ...queryable(
            unique instanceof Error
              ? unique
              : { HasUniqueRoleAssignments: unique }
          ),
          select: () =>
            queryable(
              unique instanceof Error
                ? unique
                : { HasUniqueRoleAssignments: unique }
            ),
          breakRoleInheritance: async (copy: boolean, clear: boolean) => {
            calls.push(`break ${copy} ${clear}`)
          },
          roleAssignments: {
            add: async (principal: number, role: number) => {
              calls.push(`add ${principal} ${role}`)
            }
          }
        })
      },
      siteGroups: { getByName: () => queryable({ Id: 12 }) },
      roleDefinitions: { getByName: () => queryable({ Id: 3 }) }
    }
    return { web, calls }
  }
  const lc = (breakInheritance = true): any => ({
    Title: 'Risikoer',
    Security: {
      BreakRoleInheritance: breakInheritance,
      ClearSubscopes: true,
      RoleAssignments: [{ Principal: 'Styringsgruppe', RoleDefinition: 'Lese' }]
    }
  })

  it('breaks inheritance on a list that inherits, then adds the assignments', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { web, calls } = createWeb(false)

    await handler.processListSecurity(web, lc())

    assert.deepEqual(calls, ['break false true', 'add 12 3'])
  })

  it('does not break again on a list with unique permissions, whatever ClearSubscopes says', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { web, calls } = createWeb(true)

    await handler.processListSecurity(web, lc())

    assert.deepEqual(calls, ['add 12 3'])
  })

  it('does not break when the permissions cannot be read, and still adds the assignments', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { web, calls } = createWeb(new Error('throttled'))

    const { entries } = await captureLogs(() =>
      handler.processListSecurity(web, lc())
    )

    assert.deepEqual(calls, ['add 12 3'])
    assert.equal(
      levelOf(entries, /Failed to read the permissions of list Risikoer/),
      LogLevel.Error
    )
  })

  it('warns about assignments without BreakRoleInheritance only on a list that inherits', async () => {
    const handler = createHandler(new ProvisioningContext())
    const inherits = await captureLogs(() =>
      handler.processListSecurity(createWeb(false).web, lc(false))
    )
    const unique = await captureLogs(() =>
      handler.processListSecurity(createWeb(true).web, lc(false))
    )
    const warning = /BreakRoleInheritance is not set/
    assert.equal(levelOf(inherits.entries, warning), LogLevel.Warning)
    assert.equal(levelOf(unique.entries, warning), undefined)
  })
})

describe('Lists: DataRows', () => {
  /**
   * A fake web: a list with the given fields and items (by `$filter`), a lookup source list, and
   * `ensureUser`; it records the filters, writes and lookups.
   */
  function createWeb(options: {
    fields: any[]
    matches?: (filter: string) => Array<{ Id: number }>
    source?: Record<string, number>
  }) {
    const calls: string[] = []
    const filters: string[] = []
    const web: any = {
      lists: {
        getByTitle: () => ({
          fields: queryable(options.fields),
          items: Object.assign(
            queryable(
              (filter: string) =>
                options.matches ? options.matches(filter) : [],
              filters
            ),
            {
              add: async (values: any) => {
                calls.push(`add ${JSON.stringify(values)}`)
              },
              getById: (id: number) => ({
                update: async (values: any) => {
                  calls.push(`update ${id} ${JSON.stringify(values)}`)
                }
              })
            }
          )
        }),
        getById: () => ({
          items: queryable((filter: string) => {
            calls.push(`lookup ${filter}`)
            const value = /eq '(.*)'$/.exec(filter)[1]
            const id = (options.source || {})[value]
            return id ? [{ Id: id }] : []
          })
        })
      },
      ensureUser: async (login: string) => {
        calls.push(`ensureUser ${login}`)
        return { Id: 9 }
      }
    }
    return { web, calls, filters }
  }
  const lookupField = {
    InternalName: 'GtLinks',
    TypeAsString: 'LookupMulti',
    LookupList: '{11111111-1111-4111-8111-111111111111}',
    LookupField: 'Title'
  }

  it('matches a key holding a token by its resolved value, escaped after', async () => {
    const context = new ProvisioningContext()
    context.web = { Title: "Kari's prosjekt" }
    const handler = createHandler(context)
    const { web, calls, filters } = createWeb({
      fields: [{ InternalName: 'Title', TypeAsString: 'Text' }],
      matches: () => [{ Id: 4 }]
    })

    await handler.processListDataRows(web, {
      Title: 'Risikoer',
      DataRows: { KeyColumn: 'Title', Rows: [{ Title: '{sitename} rad' }] }
    })

    assert.deepEqual(filters, ["Title eq 'Kari''s prosjekt rad'"])
    assert.deepEqual(calls, ['update 4 {"Title":"Kari\'s prosjekt rad"}'])
  })

  it('builds nothing for a row it skips: no lookup, no ensureUser', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { web, calls } = createWeb({
      fields: [
        { InternalName: 'Title', TypeAsString: 'Text' },
        lookupField,
        { InternalName: 'GtOwner', TypeAsString: 'User' }
      ],
      matches: () => [{ Id: 4 }]
    })

    const { entries } = await captureLogs(() =>
      handler.processListDataRows(web, {
        Title: 'Risikoer',
        DataRows: {
          KeyColumn: 'Title',
          UpdateBehavior: 'Skip',
          Rows: [{ Title: 'A', GtLinks: ['X'], GtOwner: 'kari@contoso.com' }]
        }
      })
    )

    assert.deepEqual(calls, [])
    assert.equal(
      levelOf(entries, /Row 0 already exists \(Title=A\)/),
      LogLevel.Info
    )
  })

  it('warns about a key that matches more than one item, and a row without a key', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { web, calls } = createWeb({
      fields: [{ InternalName: 'Title', TypeAsString: 'Text' }],
      matches: () => [{ Id: 4 }, { Id: 5 }]
    })

    const { entries } = await captureLogs(() =>
      handler.processListDataRows(web, {
        Title: 'Risikoer',
        DataRows: {
          KeyColumn: 'Title',
          Rows: [{ Title: 'A' }, { Description: 'uten nøkkel' }]
        }
      })
    )

    assert.deepEqual(calls, [
      'update 4 {"Title":"A"}',
      'add {"Description":"uten nøkkel"}'
    ])
    assert.equal(
      levelOf(entries, /More than one item in list Risikoer has Title=A/),
      LogLevel.Warning
    )
    assert.equal(
      levelOf(entries, /Row 1 in list Risikoer has no value for its KeyColumn/),
      LogLevel.Warning
    )
  })

  it('describes a URL without a Description by its resolved Url, and resolves a given one', async () => {
    const handler = createHandler(new ProvisioningContext())
    handler.tokenHelper = new TokenHelper(
      Object.assign(new ProvisioningContext(), {
        web: { Title: 'Prosjekt A' }
      }),
      { parameters: { Slug: 'a' } }
    )
    const { web, calls } = createWeb({
      fields: [
        { InternalName: 'GtLink', TypeAsString: 'URL' },
        { InternalName: 'GtOther', TypeAsString: 'URL' }
      ]
    })

    await handler.processListDataRows(web, {
      Title: 'Risikoer',
      DataRows: {
        Rows: [
          {
            GtLink: { Url: 'https://contoso.com/{parameter:Slug}' },
            GtOther: {
              Url: 'https://contoso.com/b',
              Description: '{sitename} lenke'
            }
          }
        ]
      }
    })

    assert.deepEqual(calls, [
      'add {"GtLink":{"Url":"https://contoso.com/a","Description":"https://contoso.com/a"},"GtOther":{"Url":"https://contoso.com/b","Description":"Prosjekt A lenke"}}'
    ])
  })

  it('writes the lookups that resolve on a new item, and logs each that does not as an error', async () => {
    // An unresolved LookupMulti entry used to be sent as null: SharePoint refused the whole row.
    const handler = createHandler(new ProvisioningContext())
    const { web, calls } = createWeb({
      fields: [
        lookupField,
        { ...lookupField, InternalName: 'GtLink', TypeAsString: 'Lookup' }
      ],
      source: { Kjent: 3 }
    })

    const { entries } = await captureLogs(() =>
      handler.processListDataRows(web, {
        Title: 'Risikoer',
        DataRows: {
          Rows: [{ GtLinks: ['Kjent', 'Ukjent'], GtLink: 'Borte' }]
        }
      })
    )

    assert.equal(
      calls.filter((call) => call.startsWith('add')).join(),
      'add {"GtLinksId":[3]}'
    )
    assert.equal(
      levelOf(
        entries,
        /Value "Ukjent" for column GtLinks in row 0 in list Risikoer is not written: no item in its lookup list has Title = 'Ukjent'/
      ),
      LogLevel.Error
    )
    assert.equal(
      levelOf(entries, /Value "Borte" for column GtLink in row 0/),
      LogLevel.Error
    )
  })

  it('fails the row on a lookup request that may succeed later, so the next run writes it whole', async () => {
    // Taken as "matches nothing", a Skip row would be added without the value and never filled.
    const handler = createHandler(new ProvisioningContext())
    const { web, calls } = createWeb({ fields: [lookupField] })
    web.lists.getById = () => ({
      items: queryable(
        Object.assign(new Error('Too many requests'), { status: 429 })
      )
    })

    const { entries } = await captureLogs(() =>
      handler.processListDataRows(web, {
        Title: 'Risikoer',
        DataRows: { Rows: [{ GtLinks: ['Kjent'] }] }
      })
    )

    assert.deepEqual(calls, [])
    assert.equal(
      levelOf(
        entries,
        /Failed to provision row 0 in list Risikoer: Too many requests/
      ),
      LogLevel.Error
    )
  })

  it('logs a taxonomy term it cannot write', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { web } = createWeb({
      fields: [
        {
          Id: 'tax-id',
          InternalName: 'GtPhase',
          TypeAsString: 'TaxonomyFieldType',
          TextField: '{note-id}'
        },
        { Id: 'note-id', InternalName: 'GtPhaseNote', TypeAsString: 'Note' }
      ]
    })

    const { entries } = await captureLogs(() =>
      handler.processListDataRows(web, {
        Title: 'Risikoer',
        DataRows: { Rows: [{ GtPhase: { termId: 'x' } }] }
      })
    )

    assert.equal(
      levelOf(
        entries,
        /column GtPhase in row 0 in list Risikoer is not written in full: a term needs both termId and label/
      ),
      LogLevel.Error
    )
  })

  it('keeps the stored value of an existing item when a lookup does not resolve', async () => {
    const handler = createHandler(new ProvisioningContext())
    const { web, calls } = createWeb({
      fields: [{ InternalName: 'Title', TypeAsString: 'Text' }, lookupField],
      matches: () => [{ Id: 4 }],
      source: { Kjent: 3 }
    })

    const { entries } = await captureLogs(() =>
      handler.processListDataRows(web, {
        Title: 'Risikoer',
        DataRows: {
          KeyColumn: 'Title',
          Rows: [{ Title: 'A', GtLinks: ['Kjent', 'Ukjent'] }]
        }
      })
    )

    assert.equal(
      calls.filter((call) => call.startsWith('update')).join(),
      'update 4 {"Title":"A"}'
    )
    assert.equal(levelOf(entries, /the stored value is kept/), LogLevel.Error)
  })
})

describe('ClientSidePages: failure log levels', () => {
  it('logs a part that cannot be added to a page as an error', async () => {
    const handler: any = new ClientSidePages({})
    const column = {
      addControl: () => {
        throw new Error('Invalid part')
      }
    }

    const { entries } = await captureLogs(async () =>
      handler._addControl(column, {}, { Name: 'Hurtigkoblinger' }, 'Hjem.aspx')
    )

    assert.equal(
      levelOf(
        entries,
        /Failed adding part Hurtigkoblinger to client side page Hjem\.aspx: Invalid part/
      ),
      LogLevel.Error
    )
  })
})
