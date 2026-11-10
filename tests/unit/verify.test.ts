import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  IListState,
  ISiteState,
  IViewState
} from '../../debug/lib/readback/site'
import {
  compareValue,
  diff,
  expectations,
  formatResult,
  IVerifyContext,
  queryOf
} from '../../debug/lib/verify'
import { base64EncodeString } from '../../src/util'

/**
 * `verify` (debug/lib/verify.ts) is pure between reading the token values and reading the web: these
 * tests give it a template and a web state and check the findings, rule by rule
 * (docs/specs/integration-testing.md, Harness). The integration suites run it against SharePoint.
 */

const WEB_ID = '11111111-1111-1111-1111-111111111111'
const LIST_ID = '33333333-3333-3333-3333-333333333333'

const context: IVerifyContext = {
  web: {
    Id: WEB_ID,
    Title: 'Sub',
    ServerRelativeUrl: '/sites/demo/sub',
    Url: 'https://contoso.sharepoint.com/sites/demo/sub',
    Language: 1044
  },
  site: {
    Id: '22222222-2222-2222-2222-222222222222',
    Url: 'https://contoso.sharepoint.com/sites/demo',
    ServerRelativeUrl: '/sites/demo'
  },
  lists: { 'SPJS Itest Kilde': LIST_ID },
  listViews: {},
  parameters: {}
}

/** The findings of a template against a web state, as `path: expected -> actual`. */
function findings(
  schema: Record<string, any>,
  actual: Partial<ISiteState>,
  options: { exact?: boolean } = {}
): string[] {
  return diff(expectations(schema, context, options), {
    unreadable: [],
    ...actual
  }).map(
    (finding) =>
      `${finding.handler} ${finding.path}: ${JSON.stringify(
        finding.expected
      )} -> ${JSON.stringify(finding.actual)}`
  )
}

/** A view as `readListState` returns it. */
const view = (
  Title: string,
  members: Partial<IViewState> = {}
): IViewState => ({
  Id: Title,
  Title,
  PersonalView: false,
  DefaultView: false,
  Hidden: false,
  ServerRelativeUrl: `/sites/demo/sub/Lists/L/${Title}.aspx`,
  properties: {},
  fields: ['LinkTitle'],
  ...members
})

/** A list as `readListState` returns it. */
const list = (members: Partial<IListState> = {}): IListState => ({
  properties: { BaseTemplate: 100, Description: '' },
  contentTypes: [{ Id: '0x0100ABC', Name: 'Element' }],
  fields: [],
  views: [],
  rootFolder: '/sites/demo/sub/Lists/L',
  folders: [],
  hasUniqueRoleAssignments: false,
  roleAssignments: [],
  ...members
})

describe('verify: tokens, per the engine', () => {
  it('replaces TokenHelper tokens in field XML, URL tokens in navigation and web settings', () => {
    const expected = expectations(
      {
        SiteFields: [
          `<Field ID="{aaaaaaaa-0000-0000-0000-000000000001}" Name="SPJSItestLookup" DisplayName="L" Type="Lookup" List="{listid:SPJS Itest Kilde}" WebId="{webid}" ShowField="Title" />`
        ],
        Navigation: {
          QuickLaunch: [{ Title: 'Hjem', Url: '{site}/SitePages/Home.aspx' }]
        },
        WebSettings: {
          AlternateCssUrl: '{site}/Style.css',
          Description: 'web {webid}'
        }
      },
      context
    )
    const field = expected.siteFields[0]
    assert.match(field.xml, new RegExp(LIST_ID))
    assert.match(field.xml, new RegExp(WEB_ID))
    assert.deepEqual(field.unresolved, [])
    assert.equal(field.type, 'Lookup')
    assert.equal(
      expected.navigation.QuickLaunch[0].url,
      '/sites/demo/SitePages/Home.aspx'
    )
    // WebSettings takes the URL tokens only.
    assert.equal(expected.webSettings.AlternateCssUrl, '/sites/demo/Style.css')
    assert.equal(expected.webSettings.Description, 'web {webid}')
  })

  it('replaces both in page properties, and TokenHelper only in server-processed content', () => {
    const control = expectations(
      {
        ClientSidePages: [
          {
            Name: 'p.aspx',
            Title: 'P',
            PageLayoutType: 'Article',
            Sections: [
              {
                Columns: [
                  {
                    Factor: 12,
                    Controls: [
                      {
                        Id: 'ABC',
                        Properties: { url: '{site}', web: '{webid}' },
                        ServerProcessedContent: {
                          searchablePlainTexts: { t: '{webid} {site}' }
                        }
                      }
                    ]
                  }
                ]
              }
            ]
          }
        ]
      },
      context
    ).pages[0].sections[0].columns[0].controls[0] as any
    assert.equal(control.id, 'abc')
    assert.deepEqual(control.properties, { url: '/sites/demo', web: WEB_ID })
    assert.deepEqual(control.serverProcessedContent, {
      searchablePlainTexts: { t: `${WEB_ID} {site}` }
    })
  })

  it('notes a {parameter:} without a value and compares the row as it stands', () => {
    const expected = expectations(
      {
        Lists: [
          {
            Title: 'L',
            DataRows: { KeyColumn: 'Title', Rows: [{ Title: '{parameter:x}' }] }
          }
        ]
      },
      context
    )
    assert.ok(
      expected.notes.some(
        (note) =>
          note.path === 'L/DataRows' && /\{parameter:x\}/.test(note.reason)
      )
    )
    assert.equal(expected.lists[0].rows.rows[0].key, '{parameter:x}')
  })

  it('reads only what the template names, and every section unless asked for fewer', () => {
    const schema = {
      Lists: [
        {
          Title: 'L',
          AdditionalSettings: { EnableVersioning: true },
          Views: [{ Title: 'V', AdditionalSettings: { RowLimit: 30 } }],
          DataRows: { Rows: [{ Title: 'a', Num: 1 }, { Other: 2 }] }
        }
      ],
      Features: [{ id: 'x' }],
      Hooks: []
    }
    const query = queryOf(expectations(schema, context))
    assert.deepEqual(query.lists, [
      {
        title: 'L',
        settings: ['EnableVersioning'],
        viewSettings: { V: ['RowLimit'] },
        rowColumns: ['Title', 'Num', 'Other']
      }
    ])
    assert.equal(query.features, true)
    assert.deepEqual(
      expectations(schema, context, { handlers: ['Features'] }).handlers,
      ['Features']
    )
  })
})

describe('verify: lists', () => {
  const schema = (members: Record<string, any> = {}) => ({
    Lists: [
      {
        Title: 'L',
        Template: 100,
        Views: [{ Title: 'V', ViewFields: ['LinkTitle'] }],
        ...members
      }
    ]
  })

  it('compares views, bindings and folders as a superset', () => {
    const actual = {
      lists: {
        L: list({
          views: [view('V'), view('Others')],
          folders: ['A', 'A/B']
        })
      }
    }
    assert.deepEqual(findings(schema(), actual), [])
    assert.deepEqual(
      findings(
        schema({ Folders: [{ Name: 'A', Folders: [{ Name: 'C' }] }] }),
        actual
      ),
      ['Lists L/Folders/A/C: "a folder" -> "none"']
    )
    assert.deepEqual(
      findings(
        schema({ ContentTypeBindings: [{ ContentTypeID: '0x0101' }] }),
        actual
      ),
      ['Lists L/ContentTypeBindings/0X0101: "bound" -> "not bound"']
    )
  })

  it('reports the extra views and content types RemoveExisting* leaves no room for', () => {
    const actual = {
      lists: {
        L: list({
          views: [
            view('V'),
            view('Others'),
            view('Alle', { DefaultView: true })
          ],
          contentTypes: [
            { Id: '0x0100AAA', Name: 'Element' },
            { Id: '0x0120BBB', Name: 'Mappe' }
          ]
        })
      }
    }
    assert.deepEqual(
      findings(
        schema({
          RemoveExistingViews: true,
          RemoveExistingContentTypes: true,
          ContentTypeBindings: [{ ContentTypeID: '0x0100AAA' }]
        }),
        actual
      ),
      [
        'Lists L/views: "only the template\'s views (RemoveExistingViews)" -> ["Others"]'
      ]
    )
    assert.deepEqual(
      findings(
        schema({
          RemoveExistingContentTypes: true,
          ContentTypeBindings: [{ ContentTypeID: '0x0101' }]
        }),
        { lists: { L: list({ views: [view('V')] }) } }
      ),
      [
        'Lists L/ContentTypeBindings/0X0101: "bound" -> "not bound"',
        'Lists L/ContentTypes: "only the bound content types (RemoveExistingContentTypes)" -> ["Element"]'
      ]
    )
  })

  it('compares a view query without the whitespace SharePoint does not store', () => {
    const actual = {
      lists: {
        L: list({
          views: [
            view('V', {
              properties: {
                ViewQuery: '<OrderBy><FieldRef Name="ID" /></OrderBy>'
              }
            })
          ]
        })
      }
    }
    const withQuery = (query: string) =>
      schema({
        Views: [{ Title: 'V', AdditionalSettings: { ViewQuery: query } }]
      })
    assert.deepEqual(
      findings(
        withQuery('<OrderBy>\n  <FieldRef Name="ID" />\n</OrderBy>'),
        actual
      ),
      []
    )
    // As pp-enkel-prosjektmal writes its queries (2026-11-10): a space after `=`, no space
    // before `/>`. SharePoint stores the query without them.
    assert.deepEqual(
      findings(withQuery('<OrderBy><FieldRef Name= "ID"/></OrderBy>'), actual),
      []
    )
    // Text between tags is compared as it is.
    assert.equal(
      findings(
        withQuery(
          '<Where><Eq><FieldRef Name="A" /><Value Type="Text">x = "y"</Value></Eq></Where>'
        ),
        {
          lists: {
            L: list({
              views: [
                view('V', {
                  properties: {
                    ViewQuery:
                      '<Where><Eq><FieldRef Name="A" /><Value Type="Text">x="y"</Value></Eq></Where>'
                  }
                })
              ]
            })
          }
        }
      ).length,
      1
    )
    assert.equal(
      findings(
        withQuery('<OrderBy><FieldRef Name="Title" /></OrderBy>'),
        actual
      ).length,
      1
    )
  })

  it('notes a field ref without an ID instead of reporting the column as missing', () => {
    // pp-testprosjekt 1.1.0's hub template had three list field refs with only a Name (fixed in
    // 1.1.1, 2026-11-10). The handler finds a field ref by its ID, so it logs an error for each
    // on every run, in a browser too; the column itself came with the content type.
    const template = schema({
      FieldRefs: [
        { Name: 'GtSortOrder' },
        { ID: 'aaaaaaaa-0000-0000-0000-000000000003', Name: 'GtOther' }
      ]
    })
    const expected = expectations(template, context)
    assert.deepEqual(
      expected.lists[0].fieldRefs.map((ref) => ref.name),
      ['GtOther']
    )
    assert.deepEqual(
      expected.notes.map((note) => note.path),
      ['L/FieldRefs/GtSortOrder']
    )
    assert.match(expected.notes[0].reason, /without an ID is never applied/)
    assert.deepEqual(
      findings(template, { lists: { L: list({ views: [view('V')] }) } }),
      ['Lists L/FieldRefs/GtOther: "a field" -> "none"']
    )
  })

  it('compares a list field by id, ignoring what SharePoint manages', () => {
    const xml = (attributes: string) =>
      `<Field ID="{aaaaaaaa-0000-0000-0000-000000000002}" Name="SPJSItestText" DisplayName="Tekst" Type="Text" ${attributes} />`
    const stored = {
      Id: 'aaaaaaaa-0000-0000-0000-000000000002',
      InternalName: 'SPJSItestText',
      TypeAsString: 'Text',
      Title: 'Tekst',
      Required: false,
      Hidden: false,
      SchemaXml: xml(
        'MaxLength="80" Version="3" SourceID="{x}" ColName="nvarchar5"'
      )
    }
    const actual = {
      lists: { L: list({ views: [view('V')], fields: [stored] }) }
    }
    assert.deepEqual(
      findings(
        schema({ Fields: [xml('MaxLength="80" Required="FALSE"')] }),
        actual
      ),
      []
    )
    assert.deepEqual(
      findings(schema({ Fields: [xml('MaxLength="255"')] }), actual),
      ['Lists L/Fields/SPJSItestText/@MaxLength: "255" -> "80"']
    )
    assert.deepEqual(
      findings(
        schema({ Fields: [xml('').replace('Type="Text"', 'Type="Note"')] }),
        actual
      ),
      ['Lists L/Fields/SPJSItestText/Type: "Note" -> "Text"']
    )
  })
})

describe('verify: data rows', () => {
  const rows = (UpdateBehavior: string, Rows: any[]) => ({
    Lists: [
      {
        Title: 'L',
        DataRows: { KeyColumn: 'Title', UpdateBehavior, Rows }
      }
    ]
  })
  const items = (...values: Array<Record<string, any>>) => ({
    lists: {
      L: list({
        items: {
          types: { Title: 'Text', Num: 'Number' },
          missing: [],
          expanded: true,
          items: values.map((value, index) => ({
            Id: index + 1,
            values: value
          }))
        }
      })
    }
  })

  it('finds a row by its key; a Skip row only has to exist, an Overwrite row is compared', () => {
    const web = items({ Title: 'A', Num: 2 })
    assert.deepEqual(findings(rows('Skip', [{ Title: 'A', Num: 1 }]), web), [])
    assert.deepEqual(
      findings(rows('Skip', [{ Title: 'A', Num: 1 }]), web, { exact: true }),
      ['Lists L/DataRows/0 (Title=A)/Num: 1 -> 2']
    )
    assert.deepEqual(
      findings(rows('Overwrite', [{ Title: 'A', Num: 1 }]), web),
      ['Lists L/DataRows/0 (Title=A)/Num: 1 -> 2']
    )
  })

  it('reports a missing row and a key held by more than one item', () => {
    assert.deepEqual(
      findings(rows('Skip', [{ Title: 'B' }]), items({ Title: 'A' })),
      ['Lists L/DataRows/0 (Title=B): "an item" -> "none"']
    )
    assert.deepEqual(
      findings(
        rows('Skip', [{ Title: 'A' }]),
        items({ Title: 'A' }, { Title: 'A' })
      ),
      ['Lists L/DataRows/0 (Title=A): "1 item" -> "2 items"']
    )
  })

  it('compares values by the column type', () => {
    const same = (type: string, raw: any, actual: any, resolved = raw) =>
      compareValue(type, raw, resolved, actual, true) === undefined
    assert.ok(same('Boolean', 'true', true))
    assert.ok(!same('Boolean', false, true))
    assert.ok(same('Number', '10', 10))
    assert.ok(same('MultiChoice', ['b', 'a'], ['a', 'b']))
    assert.ok(same('Note', 'a  b', '<div class="x">a b</div>'))
    assert.ok(
      same(
        'URL',
        { Url: 'https://example.org' },
        {
          Url: 'https://example.org',
          Description: 'https://example.org'
        }
      )
    )
    assert.ok(same('Lookup', 3, { ids: [3], texts: ['C'] }))
    assert.ok(same('LookupMulti', ['C'], { ids: [3], texts: ['C'] }))
    assert.ok(!same('Lookup', 3, { ids: [], texts: [] }))
    assert.ok(same('DateTime', '2026-01-02T03:04:05Z', '2026-01-02T03:04:05Z'))
    // Without a time zone the value is stored in the web's: not compared, and said so.
    assert.deepEqual(
      Object.keys(
        compareValue('DateTime', '2026-01-02T03:04:05', undefined, 'x', true)
      ),
      ['unchecked']
    )
    assert.ok(same('Text', 'x', 'x'))
    assert.ok(compareValue('Text', null, null, 'anything', true) === undefined)
  })
})

describe('verify: navigation', () => {
  const base = {
    serverRelativeUrl: '/sites/demo/sub',
    origin: 'https://contoso.sharepoint.com'
  }
  const node = (Title: string, Url: string, children: any[] = []) => ({
    Title,
    Url,
    children
  })
  const template = {
    Navigation: {
      QuickLaunch: [
        { Title: 'A', Url: '/sites/demo/sub/a' },
        {
          Title: 'B',
          Url: '/sites/demo/sub/b',
          IgnoreExisting: true,
          Children: [{ Title: 'B1', Url: '/sites/demo/sub/b1' }]
        }
      ]
    }
  }
  const web = (quickLaunch: any[]) => ({
    navigation: {
      base,
      QuickLaunch: quickLaunch,
      TopNavigationBar: [node('X', '/x')]
    }
  })

  it('compares the titles in order, and a URL only with IgnoreExisting', () => {
    const kept = [
      node('A', '/sites/demo/sub/somewhere-else'),
      node('B', 'https://contoso.sharepoint.com/sites/demo/sub/b', [
        node('B1', '/moved')
      ])
    ]
    // TopNavigationBar is not in the template: left as it is.
    assert.deepEqual(findings(template, web(kept)), [])
    assert.deepEqual(findings(template, web(kept), { exact: true }), [
      'Navigation QuickLaunch/A/Url: "/sites/demo/sub/a" -> "/sites/demo/sub/somewhere-else"',
      'Navigation QuickLaunch/B/B1/Url: "/sites/demo/sub/b1" -> "/moved"'
    ])
    assert.deepEqual(
      findings(
        template,
        web([kept[0], node('B', '/sites/demo/sub/c', kept[1].children)])
      ),
      [
        'Navigation QuickLaunch/B/Url: "/sites/demo/sub/b" -> "/sites/demo/sub/c"'
      ]
    )
  })

  it('reports a node the template does not have, at every level', () => {
    assert.deepEqual(
      findings(
        template,
        web([
          node('A', '/a'),
          node('B', '/sites/demo/sub/b', [
            node('B1', '/b1'),
            node('B2', '/b2')
          ]),
          node('C', '/c')
        ])
      ),
      [
        'Navigation QuickLaunch: ["A","B"] -> ["A","B","C"]',
        'Navigation QuickLaunch/B: ["B1"] -> ["B1","B2"]'
      ]
    )
  })
})

describe('verify: security', () => {
  const user = {
    Id: 7,
    Title: 'Test User',
    LoginName: 'i:0#.f|membership|test@contoso.com',
    Email: 'test@contoso.com',
    PrincipalType: 1
  }
  const other = {
    ...user,
    Id: 8,
    LoginName: 'i:0#.f|membership|other@contoso.com',
    Email: ''
  }
  const owners = { Id: 3, Title: 'Sub Owners', LoginName: 'Sub Owners' }
  const group = (members: any[], Owner: any = owners) => ({
    Id: 20,
    Title: 'SPJS Itest Lesere',
    Description: '',
    Owner,
    AllowMembersEditMembership: false,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: false,
    OnlyAllowMembersViewMembership: true,
    RequestToJoinLeaveEmailSetting: '',
    members
  })
  const web = (members: any[], Owner?: any) => ({
    security: {
      groups: { 'SPJS Itest Lesere': group(members, Owner) },
      web: {
        hasUniqueRoleAssignments: true,
        roleAssignments: [
          {
            member: {
              Id: 20,
              Title: 'SPJS Itest Lesere',
              LoginName: 'SPJS Itest Lesere'
            },
            roles: [{ Name: 'Lese', RoleTypeKind: 2 }]
          }
        ]
      },
      associated: { owner: owners, member: null, visitor: null }
    }
  })
  const template = (members: Record<string, any> = {}) => ({
    Security: {
      BreakRoleInheritance: true,
      SiteGroups: [
        {
          Title: 'SPJS Itest Lesere',
          Owner: '{associatedownergroupid}',
          Members: ['test@contoso.com'],
          ...members
        }
      ],
      RoleAssignments: [
        { Principal: 'SPJS Itest Lesere', RoleDefinition: 'Read' }
      ]
    }
  })

  it('compares members as a superset, an owner through the associated group and a role by type', () => {
    assert.deepEqual(findings(template(), web([user, other])), [])
    assert.deepEqual(
      findings(template({ ClearExistingMembers: true }), web([user, other])),
      [
        'Security groups/SPJS Itest Lesere/members: ["test@contoso.com"] -> ["i:0#.f|membership|test@contoso.com","i:0#.f|membership|other@contoso.com"]'
      ]
    )
    assert.deepEqual(findings(template(), web([other], user)), [
      'Security groups/SPJS Itest Lesere/members/test@contoso.com: "a member" -> "not a member"',
      'Security groups/SPJS Itest Lesere/Owner: "{associatedownergroupid}" -> "Test User"'
    ])
  })
})

describe('verify: list permissions', () => {
  const owners = { Id: 3, Title: 'Sub Owners', LoginName: 'Sub Owners' }
  const template = {
    Lists: [
      {
        Title: 'L',
        Security: {
          BreakRoleInheritance: true,
          RoleAssignments: [
            {
              Principal: '{associatedownergroupid}',
              RoleDefinition: 'Full Control'
            }
          ]
        }
      }
    ]
  }
  const web = (associatedGroups?: any) => ({
    ...(associatedGroups ? { associatedGroups } : {}),
    lists: {
      L: list({
        hasUniqueRoleAssignments: true,
        roleAssignments: [
          {
            member: owners,
            roles: [{ Name: 'Full kontroll', RoleTypeKind: 5 }]
          }
        ]
      })
    }
  })

  it('reads the associated groups for a list that names them, without a Security section', () => {
    // Found by the Lists suite (2026-11-10): without `Security` they were not read, and every
    // `{associated...groupid}` on a list was "the web has no associated owner group".
    const query = queryOf(expectations(template, context))
    assert.equal(query.associatedGroups, true)
    assert.equal(query.security, undefined)
    assert.deepEqual(
      findings(template, web({ owner: owners, member: null, visitor: null })),
      []
    )
    assert.deepEqual(
      findings(template, web({ owner: null, member: null, visitor: null })),
      [
        'Lists L/Security/RoleAssignments/{associatedownergroupid}: "Full Control" -> null'
      ]
    )
  })
})

describe('verify: what a re-apply leaves as it is', () => {
  it('only requires a custom action and a page with Overwrite false to exist, unless exact', () => {
    const template = {
      CustomActions: [
        {
          Title: 'SPJS Itest A',
          Location: 'ClientSideExtension.ApplicationCustomizer',
          Sequence: 10
        }
      ],
      ClientSidePages: [
        {
          Name: 'p.aspx',
          Title: 'P',
          PageLayoutType: 'Article',
          Overwrite: false
        }
      ]
    }
    const action = (Sequence: number) => ({
      Id: 'x',
      Title: 'SPJS Itest A',
      Name: '',
      Description: '',
      Location: 'ClientSideExtension.ApplicationCustomizer',
      Sequence,
      ClientSideComponentId: '',
      ClientSideComponentProperties: '',
      Url: '',
      RegistrationId: '',
      RegistrationType: 0,
      ScriptBlock: '',
      ScriptSrc: ''
    })
    const page = {
      title: 'Edited',
      layout: 'Article',
      commentsDisabled: false,
      published: true,
      checkedOut: false,
      sections: [],
      verticalSection: null,
      uniqueId: 'u',
      modified: 'm'
    }
    const web = { customActions: [action(20)], pages: { 'p.aspx': page } }
    assert.deepEqual(findings(template, web), [])
    assert.deepEqual(findings(template, web, { exact: true }), [
      'CustomActions SPJS Itest A/Sequence: 10 -> 20',
      'ClientSidePages p.aspx/Title: "P" -> "Edited"'
    ])
    assert.deepEqual(
      findings(template, {
        customActions: [action(10), action(10)],
        pages: { 'p.aspx': page }
      }),
      ['CustomActions SPJS Itest A: "1 custom action" -> "2"']
    )
    assert.deepEqual(
      findings(template, { customActions: [], pages: { 'p.aspx': null } }),
      [
        'CustomActions SPJS Itest A: "a custom action" -> "none"',
        'ClientSidePages p.aspx: "a page" -> "none"'
      ]
    )
  })
})

describe('verify: features, property bag and web settings', () => {
  it('compares features, written property bag entries and settings', () => {
    const template = {
      Features: [
        { id: 'AAAA0000-0000-0000-0000-000000000001' },
        { id: 'aaaa0000-0000-0000-0000-000000000002', deactivate: true }
      ],
      PropertyBagEntries: [
        { Key: 'spjs_a', Value: '1', Overwrite: true, Indexed: true },
        { Key: 'spjs_b', Value: '2' }
      ],
      WebSettings: {
        WelcomePage: 'SitePages/Home.aspx',
        QuickLaunchEnabled: true
      }
    }
    const web = {
      features: ['aaaa0000-0000-0000-0000-000000000001'],
      propertyBag: {
        spjs_a: '1',
        vti_indexedpropertykeys: `${base64EncodeString('spjs_a')}|`
      },
      webSettings: {
        WelcomePage: 'sitepages/home.aspx',
        QuickLaunchEnabled: true
      }
    }
    assert.deepEqual(findings(template, web), [
      'PropertyBagEntries spjs_b: "a value" -> null'
    ])
    assert.deepEqual(
      findings(template, {
        features: ['aaaa0000-0000-0000-0000-000000000002'],
        propertyBag: { spjs_a: '0', spjs_b: 'x' },
        webSettings: {
          WelcomePage: 'SitePages/Other.aspx',
          QuickLaunchEnabled: false
        }
      }),
      [
        'Features aaaa0000-0000-0000-0000-000000000001: "active" -> "inactive"',
        'Features aaaa0000-0000-0000-0000-000000000002: "inactive" -> "active"',
        'PropertyBagEntries spjs_a: "1" -> "0"',
        'PropertyBagEntries spjs_a/Indexed: true -> false',
        'WebSettings WelcomePage: "SitePages/Home.aspx" -> "SitePages/Other.aspx"',
        'WebSettings QuickLaunchEnabled: true -> false'
      ]
    )
  })
})

describe('verify: taxonomy', () => {
  const SET = '518a580a-11e9-4f6d-b4b5-3099c454535f'
  const A = '3689460f-6803-4d01-bb0b-2dc423374adb'
  const B = '548b20a5-9c63-45de-8e5a-a96096ef484d'
  const template = (update = false) => ({
    Taxonomy: {
      TermGroup: { Id: '0a9151e2-a998-4a44-b71c-160ef44e81b1', Name: 'G' },
      TermSets: [
        {
          Id: `{${SET.toUpperCase()}}`,
          Name: 'Faser',
          Description: 'Fasene',
          IsOpenForTermCreation: false,
          UpdateExistingTerms: update,
          Terms: [
            {
              Id: A,
              Name: 'Konsept',
              SortOrder: 2,
              CustomProperties: { Level: 'Prosjekt', Text: 'Første' }
            },
            { Id: B, Name: 'Planlegge', SortOrder: 1 }
          ]
        }
      ]
    }
  })
  const store = (
    properties: Record<string, string>,
    customSortOrder = [B, A]
  ): Partial<ISiteState> => ({
    taxonomy: {
      group: { Id: '0a9151e2-a998-4a44-b71c-160ef44e81b1', Name: 'G' },
      sets: {
        [SET]: {
          Id: SET,
          names: ['Faser'],
          description: 'Endret',
          isOpen: true,
          customSortOrder,
          terms: {
            [A]: { Id: A, labels: ['Konsept'], properties },
            [B]: { Id: B, labels: ['Planlegge'], properties: {} }
          }
        }
      }
    }
  })

  it("compares the sort order; a term's custom properties only with UpdateExistingTerms, as keys", () => {
    assert.deepEqual(findings(template(), store({})), [])
    assert.deepEqual(findings(template(), store({}, [A, B])), [
      `Taxonomy TermSets/Faser/CustomSortOrder: ${JSON.stringify([
        B,
        A
      ])} -> ${JSON.stringify([A, B])}`
    ])
    assert.deepEqual(findings(template(true), store({ Level: 'Endret' })), [
      'Taxonomy TermSets/Faser/Terms/Konsept/CustomProperties/Text: "Første" -> null'
    ])
  })

  it('compares the values, the description and IsOpenForTermCreation only when exact', () => {
    assert.deepEqual(
      findings(template(), store({ Level: 'Endret' }), { exact: true }),
      [
        'Taxonomy TermSets/Faser/Description: "Fasene" -> "Endret"',
        'Taxonomy TermSets/Faser/IsOpenForTermCreation: false -> true',
        'Taxonomy TermSets/Faser/Terms/Konsept/CustomProperties/Level: "Prosjekt" -> "Endret"',
        'Taxonomy TermSets/Faser/Terms/Konsept/CustomProperties/Text: "Første" -> null'
      ]
    )
  })

  it('asks only for the term sets and terms of the template', () => {
    assert.deepEqual(queryOf(expectations(template(), context)).taxonomy, {
      groupId: '0a9151e2-a998-4a44-b71c-160ef44e81b1',
      sets: [{ id: SET, terms: [A, B] }]
    })
  })
})

describe('verify: the result', () => {
  it('prints a line per finding and note, and a summary', () => {
    const lines = formatResult({
      handlers: ['Lists'],
      findings: [
        {
          handler: 'Lists',
          path: 'L/Template',
          expected: 101,
          actual: 100,
          source: 'template'
        }
      ],
      notes: [{ handler: 'Hooks', path: '*', reason: 'nothing to read' }]
    })
    assert.deepEqual(lines, [
      'FINDING [Lists] (template) L/Template: expected 101, found 100',
      'not checked [Hooks] *: nothing to read',
      '1 finding(s) in Lists; 1 not checked.'
    ])
  })
})
