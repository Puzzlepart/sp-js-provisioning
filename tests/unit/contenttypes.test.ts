import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { LogLevel } from '@pnp/logging'
import { ContentTypes } from '../../src/handlers/contenttypes'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { IContentType } from '../../src/schema'
import {
  CsomArg,
  CsomFailure,
  CSOM_PATH,
  FakeCsom,
  ICsomAction,
  segmentOf
} from '../helpers/csom'
import { captureLogs, levelOf } from '../helpers/logs'
import { createWireWeb } from '../helpers/transport'

const TITLE = 'fa564e0f-0c70-4ab9-b863-0177e6ddd247'
const CONTENT_TYPE = 'c042a256-787d-4a6f-8a8a-cf6ab767f12d'
const FIELD_A = 'a0000000-0000-4000-8000-00000000000a'
const FIELD_B = 'b0000000-0000-4000-8000-00000000000b'
const FIELD_C = 'c0000000-0000-4000-8000-00000000000c'

const ID = '0x0100A1B2C3D4E5F60718293A4B5C6D7E8F90'
const CHILD_ID = `${ID}01`

interface ILink {
  Id: string
  Name: string
  Required: boolean
  Hidden: boolean
}

interface IStoredContentType {
  id: string
  Name: string
  Description: string
  Group: string
  links: ILink[]
}

interface IField {
  Id: string
  InternalName: string
  Title: string
}

const link = (Id: string, Name: string, flags: Partial<ILink> = {}): ILink => ({
  Id,
  Name,
  Required: false,
  Hidden: false,
  ...flags
})

/** The links every content type below Item has. */
const itemLinks = (): ILink[] => [
  link(CONTENT_TYPE, 'ContentType'),
  link(TITLE, 'Title', { Required: true })
]

const FIELDS: IField[] = [
  { Id: TITLE, InternalName: 'Title', Title: 'Tittel' },
  { Id: FIELD_A, InternalName: 'GtFieldA', Title: 'Felt A' },
  { Id: FIELD_B, InternalName: 'GtFieldB', Title: 'Felt B' },
  { Id: FIELD_C, InternalName: 'GtFieldC', Title: 'Felt C' }
]

/**
 * A web with content types and site columns, answering what the handler sends: REST reads, and
 * CSOM requests that change the content types as SharePoint would. What a request sets, adds and
 * reorders is kept only when the request calls `Update`.
 *
 * @param initial - The web's content types and columns, and actions SharePoint refuses
 */
function createSite(
  initial: {
    contentTypes?: IStoredContentType[]
    fields?: IField[]
    fail?: (action: ICsomAction) => CsomFailure | undefined
    storedId?: (asked: string) => string
  } = {}
) {
  const contentTypes: IStoredContentType[] = initial.contentTypes || []
  const fields = initial.fields || FIELDS
  /** The `Update` calls, as `<id> Update(<updateChildren>)`. */
  const updates: string[] = []
  const find = (id: string) =>
    contentTypes.find((ct) => ct.id.toUpperCase() === id.toUpperCase())
  const text = (arg: CsomArg) => arg as string
  const guidOf = (arg: CsomArg) => (arg as { guid: string }).guid

  const csom = new FakeCsom((action, scratch) => {
    const refused = initial.fail && initial.fail(action)
    if (refused) throw refused
    const add = segmentOf(action, 'Add')
    const collection = action.segments[1] && action.segments[1].name
    if (collection !== 'ContentTypes') return undefined

    // ContentTypes.Add(ContentTypeCreationInformation): made at once, once per request.
    if (add && action.segments.indexOf(add) === 2) {
      const info = (add.args[0] as any).properties
      const asked = text(info.Id)
      if (!scratch[`made ${asked}`]) {
        if (find(asked))
          throw new CsomFailure(
            `En innholdstype med IDen ${asked} finnes allerede.`
          )
        const parent = contentTypes
          .filter((ct) => asked.toUpperCase().startsWith(ct.id.toUpperCase()))
          .sort((a, b) => b.id.length - a.id.length)[0]
        const stored = initial.storedId ? initial.storedId(asked) : asked
        contentTypes.push({
          id: stored,
          Name: text(info.Name),
          Description: info.Description === null ? '' : text(info.Description),
          Group:
            info.Group === null
              ? 'Egendefinerte innholdstyper'
              : text(info.Group),
          links: (parent ? parent.links : itemLinks()).map((l) => ({ ...l }))
        })
        scratch[`made ${asked}`] = stored
      }
      if (action.kind === 'query') return { StringId: scratch[`made ${asked}`] }
      return undefined
    }

    // ContentTypes.GetById(id): a draft of it, kept when the request calls Update.
    const getById = action.segments[2]
    const id = text(getById.args[0])
    const stored = find(id)
    if (!stored) {
      if (action.kind === 'instantiate') return { IsNull: true }
      throw new CsomFailure('Innholdstypen finnes ikke.')
    }
    const draft: IStoredContentType =
      scratch[`draft ${id}`] ||
      (scratch[`draft ${id}`] = {
        ...stored,
        links: stored.links.map((l) => ({ ...l }))
      })
    const onLinks =
      action.segments[3] && action.segments[3].name === 'FieldLinks'
    if (!onLinks) {
      if (action.kind === 'set') (draft as any)[action.name] = action.args[0]
      if (action.kind === 'call' && action.name === 'Update') {
        Object.assign(stored, draft, {
          links: draft.links.map((l) => ({ ...l }))
        })
        updates.push(`${stored.id} Update(${action.args[0]})`)
      }
      return undefined
    }
    if (action.kind === 'call' && action.name === 'Reorder') {
      const names = action.args[0] as string[]
      draft.links = names
        .map((name) => draft.links.find((l) => l.Name === name))
        .concat(draft.links.filter((l) => !names.includes(l.Name)))
      return undefined
    }
    const made = action.segments[4]
    let target: ILink
    if (made.name === 'Add') {
      const fieldRef = (made.args[0] as any).properties.Field.ref
      const fieldId = guidOf(fieldRef[fieldRef.length - 1].args[0])
      assert.equal(fieldRef[1].name, 'AvailableFields')
      target = draft.links.find((l) => l.Id === fieldId)
      if (!target || !scratch[`added ${id} ${fieldId}`]) {
        const field = fields.find((f) => f.Id === fieldId)
        if (!field) throw new CsomFailure('Kolonnen finnes ikke.')
        if (target) throw new CsomFailure('Feltet er allerede koblet.')
        target = link(field.Id, field.InternalName)
        draft.links.push(target)
        scratch[`added ${id} ${fieldId}`] = true
      }
    } else {
      target = draft.links.find((l) => l.Id === guidOf(made.args[0]))
      if (!target) throw new CsomFailure('Feltkoblingen finnes ikke.')
    }
    if (action.kind === 'set') (target as any)[action.name] = action.args[0]
    return undefined
  })

  const rest = (ct: IStoredContentType) => ({
    Id: { StringValue: ct.id },
    Name: ct.Name,
    Description: ct.Description,
    Group: ct.Group,
    FieldLinks: ct.links.map((l) => ({ ...l }))
  })
  const { web, requests } = createWireWeb((request) => {
    if (request.path === CSOM_PATH) return csom.respond(request.body)
    if (request.method !== 'GET') return undefined
    if (request.path === '_api/web/contentTypes') return contentTypes.map(rest)
    if (request.path === '_api/web/availablefields') return fields
    const one = /^_api\/web\/contentTypes\('([^']+)'\)$/.exec(request.path)
    if (one) return rest(find(one[1]))
    return undefined
  })
  return { web, requests, csom, contentTypes, updates, find }
}

/**
 * Runs the handler and returns what it logged.
 *
 * @param site - The fake web
 * @param template - The template's content types
 * @param context - Provisioning context
 */
async function provision(
  site: ReturnType<typeof createSite>,
  template: Array<Partial<IContentType>>,
  context = new ProvisioningContext()
) {
  const { entries } = await captureLogs(() =>
    new ContentTypes({}).ProvisionObjects(
      site.web,
      template as IContentType[],
      context
    )
  )
  return {
    entries,
    context,
    errors: entries
      .filter((entry) => entry.level === LogLevel.Error)
      .map((entry) => entry.message)
  }
}

/** The fixture: a content type with three field refs, in an order of its own. */
const template = (): IContentType[] => [
  {
    ID,
    Name: 'Prosjekt',
    Description: 'Et prosjekt',
    Group: 'Prosjektportalen',
    FieldRefs: [
      { ID: FIELD_B, Name: 'GtFieldB', Required: true },
      { ID: TITLE, Name: 'Title', Hidden: true },
      { ID: FIELD_A, Name: 'GtFieldA' }
    ]
  }
]

/** A web that has the fixture's content type as a first run leaves it. */
const provisioned = (change: Partial<IStoredContentType> = {}) =>
  createSite({
    contentTypes: [
      {
        id: ID,
        Name: 'Prosjekt',
        Description: 'Et prosjekt',
        Group: 'Prosjektportalen',
        links: [
          link(FIELD_B, 'GtFieldB', { Required: true }),
          link(TITLE, 'Title', { Required: true, Hidden: true }),
          link(FIELD_A, 'GtFieldA'),
          link(CONTENT_TYPE, 'ContentType')
        ],
        ...change
      }
    ]
  })

const names = (ct: IStoredContentType) => ct.links.map((l) => l.Name)
const BY_ID = `Web.ContentTypes.GetById("${ID}")`

describe('ContentTypes: a new content type', () => {
  it('is created with the id of the template, then gets its links, flags and order in one request', async () => {
    const site = createSite()

    const { errors } = await provision(site, template())

    assert.deepEqual(errors, [])
    assert.deepEqual(site.csom.requests.length, 2)
    assert.deepEqual(site.csom.ran, [
      // No parent: SharePoint reads it out of the id.
      'instantiate Web.ContentTypes.Add(ContentTypeCreationInformation{ Description: "Et prosjekt", ' +
        `Group: "Prosjektportalen", Id: "${ID}", Name: "Prosjekt", ParentContentType: null })`,
      'query Web.ContentTypes.Add(ContentTypeCreationInformation{ Description: "Et prosjekt", ' +
        `Group: "Prosjektportalen", Id: "${ID}", Name: "Prosjekt", ParentContentType: null }) [StringId]`,
      `instantiate ${BY_ID}.FieldLinks.Add(FieldLinkCreationInformation{ Field: @Web.AvailableFields.GetById({${FIELD_B}}) })`,
      `set ${BY_ID}.FieldLinks.Add(FieldLinkCreationInformation{ Field: @Web.AvailableFields.GetById({${FIELD_B}}) }).Required = true`,
      `set ${BY_ID}.FieldLinks.GetById({${TITLE}}).Hidden = true`,
      `instantiate ${BY_ID}.FieldLinks.Add(FieldLinkCreationInformation{ Field: @Web.AvailableFields.GetById({${FIELD_A}}) })`,
      `call ${BY_ID}.FieldLinks.Reorder(["GtFieldB", "Title", "GtFieldA"])`,
      `call ${BY_ID}.Update(true)`
    ])
    const stored = site.find(ID)
    assert.deepEqual(
      [stored.id, stored.Name, stored.Description, stored.Group],
      [ID, 'Prosjekt', 'Et prosjekt', 'Prosjektportalen']
    )
    assert.deepEqual(stored.links, [
      link(FIELD_B, 'GtFieldB', { Required: true }),
      link(TITLE, 'Title', { Required: true, Hidden: true }),
      link(FIELD_A, 'GtFieldA'),
      link(CONTENT_TYPE, 'ContentType')
    ])
    assert.deepEqual(site.updates, [`${ID} Update(true)`])
  })

  it('is read back through REST with its inherited links, and the columns are read once without $top', async () => {
    const site = createSite()

    await provision(site, template())

    assert.deepEqual(
      site.requests.map((request) => [
        `${request.method} ${request.path}`,
        request.query.$select,
        request.query.$expand,
        request.query.$top
      ]),
      [
        [
          'GET _api/web/contentTypes',
          'Id,Name,Description,Group,FieldLinks',
          'FieldLinks',
          undefined
        ],
        [`POST ${CSOM_PATH}`, undefined, undefined, undefined],
        [
          `GET _api/web/contentTypes('${ID}')`,
          'Id,Name,Description,Group,FieldLinks',
          'FieldLinks',
          undefined
        ],
        [
          'GET _api/web/availablefields',
          'Id,InternalName,Title',
          undefined,
          undefined
        ],
        [`POST ${CSOM_PATH}`, undefined, undefined, undefined]
      ]
    )
  })

  it('is created without a description and a group when the template has none', async () => {
    const site = createSite()

    await provision(site, [{ ID, Name: 'Uten gruppe' }])

    assert.deepEqual(site.csom.writes, [
      'instantiate Web.ContentTypes.Add(ContentTypeCreationInformation{ Description: null, ' +
        `Group: null, Id: "${ID}", Name: "Uten gruppe", ParentContentType: null })`
    ])
    // Nothing differs from what SharePoint made: no second request, no Update.
    assert.equal(site.csom.requests.length, 1)
    assert.deepEqual(site.updates, [])
  })

  it('comes after its parent, whatever the order of the template, which is left as it was', async () => {
    const site = createSite()
    const reversed: IContentType[] = [
      {
        ID: CHILD_ID,
        Name: 'Barn',
        Description: '',
        Group: 'Prosjektportalen',
        FieldRefs: [{ ID: FIELD_C, Name: 'GtFieldC' }]
      },
      ...template()
    ]

    await provision(site, reversed)

    assert.deepEqual(
      site.contentTypes.map((ct) => ct.Name),
      ['Prosjekt', 'Barn']
    )
    // The child got its parent's links when it was made, and its own after them.
    assert.deepEqual(names(site.find(CHILD_ID)), [
      'GtFieldB',
      'Title',
      'GtFieldA',
      'ContentType',
      'GtFieldC'
    ])
    assert.deepEqual(
      reversed.map((ct) => ct.Name),
      ['Barn', 'Prosjekt']
    )
  })

  it('stops the template with what SharePoint said when it cannot be created', async () => {
    const site = createSite({
      fail: (action) =>
        action.kind === 'instantiate' && segmentOf(action, 'Add')
          ? new CsomFailure(
              'Ingen tilgang. Du har ikke tillatelse til å utføre denne handlingen.',
              'System.UnauthorizedAccessException'
            )
          : undefined
    })

    await assert.rejects(
      () =>
        provision(site, [
          ...template(),
          { ID: `${ID}02`, Name: 'Kommer ikke så langt' }
        ]),
      (error: Error) => {
        assert.equal(
          error.message,
          `Failed to create content type [Prosjekt] (${ID}): Ingen tilgang. Du har ikke tillatelse til å utføre denne handlingen. (System.UnauthorizedAccessException)`
        )
        return true
      }
    )
    assert.equal(site.csom.requests.length, 1)
  })

  it('stops when SharePoint made it with another id than the template has', async () => {
    const site = createSite({ storedId: () => '0x0100FFFF' })

    await assert.rejects(
      () => provision(site, template()),
      /Content type \[Prosjekt\] was created with the id 0x0100FFFF, not 0x0100A1B2C3D4E5F60718293A4B5C6D7E8F90 as the template says\./
    )
  })
})

describe('ContentTypes: a content type the web has', () => {
  it('is sent nothing when the template has nothing new, and the columns are not read', async () => {
    const site = provisioned()

    const { entries, errors } = await provision(site, template())

    assert.deepEqual(errors, [])
    assert.deepEqual(
      site.requests.map((request) => `${request.method} ${request.path}`),
      ['GET _api/web/contentTypes']
    )
    assert.deepEqual(site.updates, [])
    assert.equal(
      levelOf(
        entries,
        /Content type \[Prosjekt\] \(0x0100A1B2C3D4E5F60718293A4B5C6D7E8F90\) is unchanged\./
      ),
      LogLevel.Info
    )
  })

  it('is found whatever the case of the hex digits of its id', async () => {
    const site = provisioned()

    await provision(site, [{ ...template()[0], ID: ID.toLowerCase() }])

    assert.deepEqual(site.csom.requests, [])
  })

  it('gets a changed name, description and group, and nothing else', async () => {
    const site = provisioned()
    const changed = template()
    changed[0].Name = 'Prosjekt (ny)'
    changed[0].Description = 'Et prosjekt, beskrevet på nytt'
    changed[0].Group = 'Prosjektportalen innholdstyper'

    const { context } = await provision(site, changed)

    assert.deepEqual(site.csom.ran, [
      `set ${BY_ID}.Name = "Prosjekt (ny)"`,
      `set ${BY_ID}.Description = "Et prosjekt, beskrevet på nytt"`,
      `set ${BY_ID}.Group = "Prosjektportalen innholdstyper"`,
      `call ${BY_ID}.Update(true)`
    ])
    assert.equal(site.find(ID).Name, 'Prosjekt (ny)')
    // The context follows: the new name finds it, the old one no longer does.
    assert.equal(context.contentTypes['Prosjekt (ny)'].ID, ID)
    assert.equal(context.contentTypes[ID].Name, 'Prosjekt (ny)')
    assert.equal(context.contentTypes['Prosjekt'], undefined)
  })

  it('keeps its description and group when the template has none', async () => {
    const site = provisioned()

    const { context } = await provision(site, [
      { ID, Name: 'Prosjekt', FieldRefs: template()[0].FieldRefs }
    ])

    assert.deepEqual(site.csom.requests, [])
    assert.deepEqual(
      [context.contentTypes[ID].Description, context.contentTypes[ID].Group],
      ['Et prosjekt', 'Prosjektportalen']
    )
  })

  it('gets a link it lacks, and a changed Required and Hidden on the ones it has', async () => {
    const site = provisioned()
    const changed = template()
    changed[0].FieldRefs = [
      { ID: FIELD_B, Name: 'GtFieldB', Required: false },
      { ID: TITLE, Name: 'Title', Hidden: true },
      { ID: FIELD_A, Name: 'GtFieldA', Hidden: true },
      { ID: FIELD_C, Name: 'GtFieldC', Required: true, Hidden: false }
    ]

    const { context } = await provision(site, changed)

    const added = `${BY_ID}.FieldLinks.Add(FieldLinkCreationInformation{ Field: @Web.AvailableFields.GetById({${FIELD_C}}) })`
    assert.deepEqual(site.csom.ran, [
      `set ${BY_ID}.FieldLinks.GetById({${FIELD_B}}).Required = false`,
      `set ${BY_ID}.FieldLinks.GetById({${FIELD_A}}).Hidden = true`,
      `instantiate ${added}`,
      `set ${added}.Required = true`,
      `set ${added}.Hidden = false`,
      // The new link is last, the others first and in order: `Reorder` puts it fourth.
      `call ${BY_ID}.FieldLinks.Reorder(["GtFieldB", "Title", "GtFieldA", "GtFieldC"])`,
      `call ${BY_ID}.Update(true)`
    ])
    assert.deepEqual(site.find(ID).links, [
      link(FIELD_B, 'GtFieldB'),
      link(TITLE, 'Title', { Required: true, Hidden: true }),
      link(FIELD_A, 'GtFieldA', { Hidden: true }),
      link(FIELD_C, 'GtFieldC', { Required: true }),
      link(CONTENT_TYPE, 'ContentType')
    ])
    assert.deepEqual(
      context.contentTypes[ID].FieldRefs.map((fieldRef) => fieldRef.Name),
      ['GtFieldB', 'Title', 'GtFieldA', 'GtFieldC', 'ContentType']
    )
  })

  it('leaves Required and Hidden alone where the template does not give them', async () => {
    const site = provisioned()
    const changed = template()
    // Title is Required and Hidden on the web; the template now says neither.
    changed[0].FieldRefs = [
      { ID: FIELD_B, Name: 'GtFieldB' },
      { ID: TITLE, Name: 'Title' },
      { ID: FIELD_A, Name: 'GtFieldA' }
    ]

    await provision(site, changed)

    assert.deepEqual(site.csom.requests, [])
  })

  it('has its links reordered only when they are not first and in the order of the template', async () => {
    const site = provisioned()
    const changed = template()
    changed[0].FieldRefs = [
      { ID: TITLE, Name: 'Title' },
      { ID: FIELD_A, Name: 'GtFieldA' },
      { ID: FIELD_B, Name: 'GtFieldB' }
    ]

    await provision(site, changed)

    assert.deepEqual(site.csom.ran, [
      `call ${BY_ID}.FieldLinks.Reorder(["Title", "GtFieldA", "GtFieldB"])`,
      `call ${BY_ID}.Update(true)`
    ])
    assert.deepEqual(names(site.find(ID)), [
      'Title',
      'GtFieldA',
      'GtFieldB',
      'ContentType'
    ])

    // The same template again: nothing.
    await provision(site, changed)
    assert.equal(site.csom.requests.length, 1)
  })

  it('is not reordered for a single field ref', async () => {
    const site = provisioned()

    await provision(site, [
      { ID, Name: 'Prosjekt', FieldRefs: [{ ID: FIELD_A, Name: 'GtFieldA' }] }
    ])

    assert.deepEqual(site.csom.requests, [])
    assert.equal(names(site.find(ID))[0], 'GtFieldB')
  })

  it('names a field ref the template has twice once', async () => {
    const site = provisioned()
    const changed = template()
    changed[0].FieldRefs = [
      { ID: FIELD_A, Name: 'GtFieldA' },
      { ID: FIELD_B, Name: 'GtFieldB' },
      { ID: FIELD_A, Name: 'GtFieldA' }
    ]

    await provision(site, changed)

    assert.deepEqual(site.csom.ran, [
      `call ${BY_ID}.FieldLinks.Reorder(["GtFieldA", "GtFieldB"])`,
      `call ${BY_ID}.Update(true)`
    ])
  })

  it('is matched by name when the template gives no id', async () => {
    const site = provisioned()

    const { context } = await provision(site, [
      { Name: 'Prosjekt', Description: 'Funnet på navnet' }
    ])

    assert.deepEqual(site.csom.ran, [
      `set ${BY_ID}.Description = "Funnet på navnet"`,
      `call ${BY_ID}.Update(true)`
    ])
    assert.equal(context.contentTypes[ID].Description, 'Funnet på navnet')
  })

  it('stops the template when a content type without id is not on the web', async () => {
    const site = provisioned()

    await assert.rejects(
      () => provision(site, [{ Name: 'Finnes ikke' }]),
      /Content type with name 'Finnes ikke' does not exist in the web\./
    )
    assert.deepEqual(site.csom.requests, [])
  })

  it('stops the template with what SharePoint said when it cannot be updated', async () => {
    const site = provisioned()
    const failing = createSite({
      contentTypes: site.contentTypes,
      fail: (action) =>
        action.name === 'Update'
          ? new CsomFailure('Innholdstypen er skrivebeskyttet.')
          : undefined
    })
    const changed = template()
    changed[0].Name = 'Nytt navn'

    const error: Error = await provision(failing, changed).then(
      () => assert.fail('the run did not stop'),
      (thrown) => thrown
    )

    assert.equal(
      error.message,
      `Failed to update content type [Nytt navn] (${ID}): Innholdstypen er skrivebeskyttet. (Microsoft.SharePoint.SPException)`
    )
    assert.equal(failing.find(ID).Name, 'Prosjekt')
  })
})

describe('ContentTypes: the site column of a field ref', () => {
  it('is found by internal name, else by id, else by title', async () => {
    const site = createSite()

    await provision(site, [
      {
        ID,
        Name: 'Prosjekt',
        FieldRefs: [
          // The name wins over an id that is another column's.
          { ID: FIELD_C, Name: 'GtFieldA' },
          // A name the web does not know: the id.
          { ID: `{${FIELD_B.toUpperCase()}}`, Name: 'GtRenamed' },
          // No id: the title, as getByInternalNameOrTitle did.
          { ID: undefined, Name: 'Felt C' }
        ]
      }
    ])

    assert.deepEqual(names(site.find(ID)), [
      'GtFieldA',
      'GtFieldB',
      'GtFieldC',
      'ContentType',
      'Title'
    ])
  })

  it('is an error that names it when the web does not have it, and the other links are added', async () => {
    const site = createSite()

    const { errors, entries } = await provision(site, [
      {
        ID,
        Name: 'Prosjekt',
        FieldRefs: [
          { ID: FIELD_A, Name: 'GtFieldA' },
          { ID: 'd0000000-0000-4000-8000-00000000000d', Name: 'GtMissing' },
          { ID: undefined, Name: 'GtAlsoMissing' },
          { ID: 'e0000000-0000-4000-8000-00000000000e' },
          { ID: FIELD_B, Name: 'GtFieldB' }
        ]
      }
    ])

    assert.deepEqual(
      errors.map((message) => message.replace(/^.*\): /, '')),
      [
        `Field ref GtMissing of content type [Prosjekt] (${ID}) is left out: the web has no site column with that name or id.`,
        `Field ref GtAlsoMissing of content type [Prosjekt] (${ID}) is left out: the web has no site column with that name.`,
        `Field ref e0000000-0000-4000-8000-00000000000e of content type [Prosjekt] (${ID}) is left out: the web has no site column with that id.`
      ]
    )
    assert.equal(levelOf(entries, /GtMissing/), LogLevel.Error)
    assert.deepEqual(names(site.find(ID)), [
      'GtFieldA',
      'GtFieldB',
      'ContentType',
      'Title'
    ])
    assert.deepEqual(site.updates, [`${ID} Update(true)`])
  })
})
