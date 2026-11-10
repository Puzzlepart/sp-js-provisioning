import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { LogLevel } from '@pnp/logging'
import { Taxonomy } from '../../src/handlers/taxonomy'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { ITaxonomy } from '../../src/schema'
import {
  CsomArg,
  CsomFailure,
  CSOM_PATH,
  FakeCsom,
  ICsomAction,
  ICsomSegment
} from '../helpers/csom'
import { captureLogs, levelOf } from '../helpers/logs'
import { createWireWeb } from '../helpers/transport'

const GROUP = '11111111-0000-4000-8000-000000000001'
const SET = '22222222-0000-4000-8000-000000000002'
const TERM_A = '33333333-0000-4000-8000-00000000000a'
const TERM_B = '33333333-0000-4000-8000-00000000000b'
const TERM_C = '33333333-0000-4000-8000-00000000000c'

interface IStore {
  groups: Record<string, { name: string }>
  sets: Record<
    string,
    {
      name: string
      group: string
      lcid: number
      description: string
      open: boolean
      customSortOrder: string | null
    }
  >
  terms: Record<
    string,
    {
      name: string
      set: string
      lcid: number
      customProperties: Record<string, string>
    }
  >
}

const empty = (): IStore => ({ groups: {}, sets: {}, terms: {} })
const clone = (store: IStore): IStore => JSON.parse(JSON.stringify(store))
const guidArg = (arg: CsomArg) => (arg as { guid: string }).guid

/**
 * A default term store answering the handler's CSOM requests as SharePoint does: what a request
 * creates or sets is kept only when it calls `CommitAll`; a failing action fails the request
 * there, and nothing of it is kept.
 *
 * @param initial - What the store holds, and actions SharePoint refuses
 */
function createTermStore(
  initial: {
    store?: IStore
    noStore?: boolean
    fail?: (action: ICsomAction) => CsomFailure | undefined
  } = {}
) {
  let committed = initial.store ? clone(initial.store) : empty()
  /** `CommitAll` calls. */
  const commits: number[] = []

  const csom = new FakeCsom((action, scratch) => {
    const refused = initial.fail && initial.fail(action)
    if (refused) throw refused
    const pending: IStore =
      scratch.pending || (scratch.pending = clone(committed))
    const [type, session, opened] = action.segments
    assert.deepEqual(
      [type.name, session.name, opened && opened.name],
      [
        'TaxonomySession',
        'GetTaxonomySession',
        'GetDefaultSiteCollectionTermStore'
      ]
    )

    /** The object a path ends at: `store`, a group, a term set, a term, or null. */
    const resolve = (
      segments: ICsomSegment[]
    ): { kind: string; id?: string } | null => {
      if (initial.noStore) return null
      let at: { kind: string; id?: string } = { kind: 'store' }
      for (const segment of segments.slice(3)) {
        const id = guidArg(
          segment.args.find((arg) => typeof arg === 'object' && 'guid' in arg)
        )
        switch (segment.name) {
          case 'GetGroup':
            at = pending.groups[id] ? { kind: 'group', id } : null
            break
          case 'GetTermSet':
            at = pending.sets[id] ? { kind: 'set', id } : null
            break
          case 'GetTerm':
            at = pending.terms[id] ? { kind: 'term', id } : null
            break
          case 'CreateGroup':
            if (!scratch[`group ${id}`]) {
              if (pending.groups[id])
                throw new CsomFailure('Gruppen finnes allerede.')
              pending.groups[id] = { name: segment.args[0] as string }
              scratch[`group ${id}`] = true
            }
            at = { kind: 'group', id }
            break
          case 'CreateTermSet':
            if (!scratch[`set ${id}`]) {
              if (pending.sets[id])
                throw new CsomFailure('Termsettet finnes allerede.')
              pending.sets[id] = {
                name: segment.args[0] as string,
                group: at.id,
                lcid: segment.args[2] as number,
                description: '',
                open: true,
                customSortOrder: null
              }
              scratch[`set ${id}`] = true
            }
            at = { kind: 'set', id }
            break
          case 'CreateTerm':
            if (!scratch[`term ${id}`]) {
              if (pending.terms[id])
                throw new CsomFailure('Termen finnes allerede.')
              pending.terms[id] = {
                name: segment.args[0] as string,
                set: at.id,
                lcid: segment.args[1] as number,
                customProperties: {}
              }
              scratch[`term ${id}`] = true
            }
            at = { kind: 'term', id }
            break
          default:
            throw new Error(`The fake has no ${segment.name}`)
        }
        if (!at) return null
      }
      return at
    }

    const target = resolve(action.segments)
    if (action.kind === 'instantiate') return { IsNull: !target }
    if (action.kind === 'query') {
      if (!target) return null
      if (target.kind === 'store') return { DefaultLanguage: 1044 }
      if (target.kind === 'set')
        return { CustomSortOrder: pending.sets[target.id].customSortOrder }
      if (target.kind === 'term')
        return {
          CustomProperties: { ...pending.terms[target.id].customProperties }
        }
      return {}
    }
    if (!target) throw new CsomFailure('Objektet finnes ikke.')
    if (action.kind === 'set') {
      const set = pending.sets[target.id]
      const value = action.args[0]
      if (action.name === 'Description') set.description = value as string
      else if (action.name === 'IsOpenForTermCreation')
        set.open = value as boolean
      else if (action.name === 'CustomSortOrder')
        set.customSortOrder = value as string
      else throw new Error(`The fake cannot set ${action.name}`)
      return undefined
    }
    if (action.name === 'SetCustomProperty') {
      pending.terms[target.id].customProperties[
        action.args[0] as string
      ] = action.args[1] as string
      return undefined
    }
    if (action.name === 'CommitAll') {
      committed = clone(pending)
      commits.push(commits.length + 1)
      return undefined
    }
    throw new Error(`The fake has no ${action.name}`)
  })
  const { web, requests } = createWireWeb((request) =>
    request.path === CSOM_PATH ? csom.respond(request.body) : undefined
  )
  return { web, requests, csom, commits, store: () => committed }
}

/**
 * Runs the handler and returns what it logged.
 *
 * @param site - The fake term store
 * @param taxonomy - The template's Taxonomy section
 */
async function provision(
  site: ReturnType<typeof createTermStore>,
  taxonomy: ITaxonomy
) {
  const { entries } = await captureLogs(() =>
    new Taxonomy({}).ProvisionObjects(
      site.web,
      taxonomy,
      new ProvisioningContext()
    )
  )
  return { entries }
}

/** A term group with one term set of three terms, sorted B, A, C. */
const template = (): ITaxonomy => ({
  TermGroup: { Id: GROUP, Name: 'Prosjektportalen' },
  TermSets: [
    {
      Id: `{${SET.toUpperCase()}}`,
      Name: 'Fase',
      Description: 'Prosjektfaser',
      IsOpenForTermCreation: false,
      Terms: [
        {
          Id: TERM_A,
          Name: 'Konsept',
          SortOrder: 2,
          CustomProperties: { PhaseLevel: 'Prosjekt' }
        },
        { Id: TERM_B, Name: 'Initiering', SortOrder: 1 },
        {
          Id: TERM_C,
          Name: 'Gjennomføre',
          SortOrder: 3,
          CustomProperties: { PhaseLevel: 'Prosjekt', PhaseSubText: 'Gjør' }
        }
      ]
    }
  ]
})

/** The store a first run of the template leaves. */
const provisioned = (): IStore => ({
  groups: { [GROUP]: { name: 'Prosjektportalen' } },
  sets: {
    [SET]: {
      name: 'Fase',
      group: GROUP,
      lcid: 1044,
      description: 'Prosjektfaser',
      open: false,
      customSortOrder: `${TERM_B}:${TERM_A}:${TERM_C}`
    }
  },
  terms: {
    [TERM_A]: {
      name: 'Konsept',
      set: SET,
      lcid: 1044,
      customProperties: { PhaseLevel: 'Prosjekt' }
    },
    [TERM_B]: {
      name: 'Initiering',
      set: SET,
      lcid: 1044,
      customProperties: {}
    },
    [TERM_C]: {
      name: 'Gjennomføre',
      set: SET,
      lcid: 1044,
      customProperties: { PhaseLevel: 'Prosjekt', PhaseSubText: 'Gjør' }
    }
  }
})

const STORE =
  'TaxonomySession.GetTaxonomySession().GetDefaultSiteCollectionTermStore()'

describe('Taxonomy: an empty term store', () => {
  it('gets the group, then the term set with its terms, custom properties and sort order, with the ids of the template', async () => {
    const site = createTermStore()

    await provision(site, template())

    assert.deepEqual(site.store(), provisioned())
    const createdSet = `${STORE}.GetGroup({${GROUP}}).CreateTermSet("Fase", {${SET}}, 1044)`
    const term = (name: string, id: string) =>
      `${createdSet}.CreateTerm("${name}", 1044, {${id}})`
    assert.deepEqual(site.csom.writes, [
      `instantiate ${STORE}.CreateGroup("Prosjektportalen", {${GROUP}})`,
      `call ${STORE}.CommitAll()`,
      `instantiate ${createdSet}`,
      `set ${createdSet}.Description = "Prosjektfaser"`,
      `set ${createdSet}.IsOpenForTermCreation = false`,
      `instantiate ${term('Konsept', TERM_A)}`,
      `call ${term(
        'Konsept',
        TERM_A
      )}.SetCustomProperty("PhaseLevel", "Prosjekt")`,
      `instantiate ${term('Initiering', TERM_B)}`,
      `instantiate ${term('Gjennomføre', TERM_C)}`,
      `call ${term(
        'Gjennomføre',
        TERM_C
      )}.SetCustomProperty("PhaseLevel", "Prosjekt")`,
      `call ${term(
        'Gjennomføre',
        TERM_C
      )}.SetCustomProperty("PhaseSubText", "Gjør")`,
      `set ${createdSet}.CustomSortOrder = "${TERM_B}:${TERM_A}:${TERM_C}"`,
      `call ${STORE}.CommitAll()`
    ])
    // A read, the group, the term set.
    assert.equal(site.csom.requests.length, 3)
    assert.deepEqual(
      site.requests.map((request) => `${request.method} ${request.path}`),
      [`POST ${CSOM_PATH}`, `POST ${CSOM_PATH}`, `POST ${CSOM_PATH}`]
    )
  })

  it('reads everything in one request: the language, the group, each term set and each term', async () => {
    const site = createTermStore()

    await provision(site, template())

    assert.deepEqual(
      site.csom.requests[0].map((action) => action.text),
      [
        `instantiate ${STORE}`,
        `query ${STORE} [DefaultLanguage]`,
        `instantiate ${STORE}.GetGroup({${GROUP}})`,
        `instantiate ${STORE}.GetTermSet({${SET}})`,
        `query ${STORE}.GetTermSet({${SET}}) [CustomSortOrder]`,
        `instantiate ${STORE}.GetTerm({${TERM_A}})`,
        `instantiate ${STORE}.GetTerm({${TERM_B}})`,
        `instantiate ${STORE}.GetTerm({${TERM_C}})`
      ]
    )
  })

  it('leaves the template as it was', async () => {
    const site = createTermStore()
    const taxonomy = template()

    await provision(site, taxonomy)

    assert.deepEqual(taxonomy, template())
  })
})

describe('Taxonomy: a term store that has the template', () => {
  it('is read once and sent nothing else', async () => {
    const site = createTermStore({ store: provisioned() })

    const { entries } = await provision(site, template())

    assert.equal(site.csom.requests.length, 1)
    assert.deepEqual(site.csom.writes, [])
    assert.equal(
      levelOf(entries, /Term set Fase is unchanged\./),
      LogLevel.Info
    )
  })

  it('gets a missing term with its custom properties, and leaves the others as they are', async () => {
    const store = provisioned()
    delete store.terms[TERM_C]
    store.sets[SET].customSortOrder = `${TERM_B}:${TERM_A}`
    store.terms[TERM_A].customProperties = {}
    const site = createTermStore({ store })

    await provision(site, template())

    const termSet = `${STORE}.GetTermSet({${SET}})`
    const created = `${termSet}.CreateTerm("Gjennomføre", 1044, {${TERM_C}})`
    assert.deepEqual(site.csom.writes, [
      `instantiate ${created}`,
      `call ${created}.SetCustomProperty("PhaseLevel", "Prosjekt")`,
      `call ${created}.SetCustomProperty("PhaseSubText", "Gjør")`,
      `set ${termSet}.CustomSortOrder = "${TERM_B}:${TERM_A}:${TERM_C}"`,
      `call ${STORE}.CommitAll()`
    ])
    // Without UpdateExistingTerms an existing term is not touched.
    assert.deepEqual(site.store().terms[TERM_A].customProperties, {})
  })

  it('sets a changed sort order, and compares it without regard to case and braces', async () => {
    const store = provisioned()
    store.sets[
      SET
    ].customSortOrder = `{${TERM_B.toUpperCase()}}:${TERM_A}:${TERM_C}`
    const same = createTermStore({ store })

    await provision(same, template())
    assert.deepEqual(same.csom.writes, [])

    const changed = template()
    changed.TermSets[0].Terms[2].SortOrder = 0
    const site = createTermStore({ store: provisioned() })
    await provision(site, changed)
    assert.deepEqual(site.csom.writes, [
      `set ${STORE}.GetTermSet({${SET}}).CustomSortOrder = "${TERM_C}:${TERM_B}:${TERM_A}"`,
      `call ${STORE}.CommitAll()`
    ])
  })

  it('sets no sort order when fewer than two terms have one', async () => {
    const store = provisioned()
    store.sets[SET].customSortOrder = null
    const site = createTermStore({ store })
    const one = template()
    delete one.TermSets[0].Terms[0].SortOrder
    delete one.TermSets[0].Terms[1].SortOrder

    await provision(site, one)

    assert.deepEqual(site.csom.writes, [])
  })

  it('adds the custom properties an existing term lacks with UpdateExistingTerms, and never overwrites one', async () => {
    const store = provisioned()
    store.terms[TERM_C].customProperties = {
      PhaseLevel: 'Satt av en administrator'
    }
    const site = createTermStore({ store })
    const update = template()
    update.TermSets[0].UpdateExistingTerms = true

    await provision(site, update)

    assert.deepEqual(site.csom.writes, [
      `call ${STORE}.GetTerm({${TERM_C}}).SetCustomProperty("PhaseSubText", "Gjør")`,
      `call ${STORE}.CommitAll()`
    ])
    assert.deepEqual(site.store().terms[TERM_C].customProperties, {
      PhaseLevel: 'Satt av en administrator',
      PhaseSubText: 'Gjør'
    })
    // The read asked for the custom properties of the terms that have some in the template.
    assert.deepEqual(
      site.csom.requests[0]
        .filter((action) => action.kind === 'query')
        .map((action) => action.text),
      [
        `query ${STORE} [DefaultLanguage]`,
        `query ${STORE}.GetTermSet({${SET}}) [CustomSortOrder]`,
        `query ${STORE}.GetTerm({${TERM_A}}) [CustomProperties]`,
        `query ${STORE}.GetTerm({${TERM_C}}) [CustomProperties]`
      ]
    )
  })

  it('creates a missing term set in the group the store has', async () => {
    const store = provisioned()
    store.sets = {}
    store.terms = {}
    const site = createTermStore({ store })

    await provision(site, template())

    assert.ok(
      !site.csom.writes.some((line) => line.includes('CreateGroup')),
      site.csom.writes.join('\n')
    )
    assert.deepEqual(site.store(), provisioned())
  })
})

describe('Taxonomy: what stops the template', () => {
  it('a term store that is not there', async () => {
    const site = createTermStore({ noStore: true })

    await assert.rejects(
      () => provision(site, template()),
      /^Error: No default site collection term store is available\.$/
    )
    assert.equal(site.csom.requests.length, 1)
  })

  it("a refused write, in SharePoint's words, and nothing after it", async () => {
    const site = createTermStore({
      fail: (action) =>
        action.text.includes('CreateGroup')
          ? new CsomFailure(
              'Ingen tilgang.',
              'System.UnauthorizedAccessException'
            )
          : undefined
    })

    await assert.rejects(
      () => provision(site, template()),
      (error: Error) => {
        assert.equal(
          error.message,
          `Failed to create term group Prosjektportalen (${GROUP}): Ingen tilgang. (System.UnauthorizedAccessException)`
        )
        return true
      }
    )
    assert.deepEqual(site.store(), empty())
    assert.equal(site.csom.requests.length, 2)
  })

  it('a term set SharePoint refuses: nothing of it is kept', async () => {
    const site = createTermStore({
      store: { ...empty(), groups: { [GROUP]: { name: 'Prosjektportalen' } } },
      fail: (action) =>
        action.kind === 'call' && action.name === 'SetCustomProperty'
          ? new CsomFailure('Ugyldig egenskap.')
          : undefined
    })

    await assert.rejects(
      () => provision(site, template()),
      new RegExp(
        `^Error: Failed to provision term set Fase \\(${SET}\\): Ugyldig egenskap\\. \\(Microsoft\\.SharePoint\\.SPException\\)$`
      )
    )
    assert.deepEqual(site.store().sets, {})
    assert.deepEqual(site.store().terms, {})
  })

  it('an id that is not a GUID, before anything is sent', async () => {
    const site = createTermStore()
    const broken = template()
    broken.TermSets[0].Terms[1].Id = 'not-a-guid'

    await assert.rejects(
      () => provision(site, broken),
      /^Error: Term Initiering of term set Fase has no valid Id \(not-a-guid\)\.$/
    )
    assert.deepEqual(site.requests, [])
  })

  it('a Taxonomy section without a term group', async () => {
    const site = createTermStore()

    await assert.rejects(
      () => provision(site, ({ TermSets: [] } as unknown) as ITaxonomy),
      /The Taxonomy section has no TermGroup\./
    )
  })
})
