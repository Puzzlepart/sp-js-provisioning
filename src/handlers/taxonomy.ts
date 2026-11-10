/* eslint-disable unicorn/prevent-abbreviations */
import { IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { ITaxonomy, ITerm, ITermSet } from '../schema'
import {
  CsomRequest,
  describeCsomError,
  executeQuery,
  guid,
  guidOf,
  isNull
} from '../util/csom'
import { HandlerBase } from './handlerbase'

/** `SP.Taxonomy.TaxonomySession`, whose static `GetTaxonomySession` opens the term store. */
const TAXONOMY_SESSION = '{981cbc68-9edc-4f8d-872f-71146fcbb84f}'

/** A term as the term store has it. */
interface ITermState {
  exists: boolean
  /** Read only when the term set has `UpdateExistingTerms`. */
  customProperties: Record<string, string>
}

/** A term set as the term store has it, with the template's terms. */
interface ITermSetState {
  exists: boolean
  customSortOrder: string
  terms: Record<string, ITermState>
}

/** What the default term store holds of a template. */
interface ITermStoreState {
  defaultLanguage: number
  groupExists: boolean
  sets: Record<string, ITermSetState>
}

/**
 * A sort order (`id:id:…`) as ids in lower case without braces, to compare.
 *
 * @param order - The sort order
 */
const sortKey = (order: string): string =>
  (order || '')
    .split(':')
    .map((id) => guidOf(id) || id.trim().toLowerCase())
    .join(':')

/**
 * Describes the Taxonomy Object Handler.
 *
 * Provisions a term group, its term sets and terms into the default site collection term store,
 * with the ids the template gives them, through CSOM over HTTP (`src/util/csom.ts`). The store is
 * read first, in one request; then a term set that needs anything gets one request that ends with
 * `CommitAll`: the term set when it is missing, the terms that are missing with their custom
 * properties, the custom properties an existing term lacks (`UpdateExistingTerms`, never
 * overwriting), and the custom sort order when it differs. What exists is otherwise left as it is,
 * and a term set that needs nothing gets no request.
 */
export class Taxonomy extends HandlerBase {
  public context: ProvisioningContext

  /**
   * Creates a new instance of the Taxonomy handler.
   */
  constructor(config: IProvisioningConfig) {
    super('Taxonomy', config)
  }

  /**
   * Provisioning taxonomy. A term store that cannot be read or written stops the template: the
   * site columns that come next would point at term sets that are not there.
   *
   * @param web - The web
   * @param taxonomy - The taxonomy part of the template
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: IWeb,
    taxonomy: ITaxonomy,
    context: ProvisioningContext
  ): Promise<void> {
    this.context = context
    super.scope_started()
    try {
      this.validate(taxonomy)
      const termSets = taxonomy.TermSets || []
      const state = await this.read(web, taxonomy)
      if (!state.groupExists) await this.createGroup(web, taxonomy)
      else
        super.log_info(
          'ensureGroup',
          `Term group ${taxonomy.TermGroup.Name} already exists`
        )
      for (const termSet of termSets)
        await this.processTermSet(
          web,
          taxonomy,
          termSet,
          state.sets[termSet.Id],
          state.defaultLanguage
        )
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Checks the ids the template gives before anything is sent: each is written into a request.
   *
   * @param taxonomy - The taxonomy part of the template
   */
  private validate(taxonomy: ITaxonomy): void {
    const check = (id: string, what: string) => {
      if (!guidOf(id) || !/^{?[\da-f-]{36}}?$/i.test(String(id).trim()))
        throw new Error(`${what} has no valid Id (${String(id)}).`)
    }
    if (!taxonomy || !taxonomy.TermGroup)
      throw new Error('The Taxonomy section has no TermGroup.')
    check(taxonomy.TermGroup.Id, `Term group ${taxonomy.TermGroup.Name}`)
    for (const termSet of taxonomy.TermSets || []) {
      check(termSet.Id, `Term set ${termSet.Name}`)
      for (const term of termSet.Terms || [])
        check(term.Id, `Term ${term.Name} of term set ${termSet.Name}`)
    }
  }

  /**
   * The default site collection term store, as an object path of a request.
   *
   * @param request - The request
   */
  private termStore(request: CsomRequest): number {
    return request.method(
      request.staticMethod(TAXONOMY_SESSION, 'GetTaxonomySession'),
      'GetDefaultSiteCollectionTermStore'
    )
  }

  /**
   * Reads, in one request, the term store's default language and whether the group, each term set
   * and each term is there; a term set's custom sort order, and a term's custom properties when
   * its term set has `UpdateExistingTerms`. Terms and term sets are found by id anywhere in the
   * store, as before.
   *
   * @param web - The web
   * @param taxonomy - The taxonomy part of the template
   */
  private async read(web: IWeb, taxonomy: ITaxonomy): Promise<ITermStoreState> {
    const request = new CsomRequest()
    const store = this.termStore(request)
    const storeFound = request.instantiate(store)
    const storeRead = request.query(store, ['DefaultLanguage'])
    const groupFound = request.instantiate(
      request.method(store, 'GetGroup', guid(taxonomy.TermGroup.Id))
    )
    const reads = (taxonomy.TermSets || []).map((termSet) => {
      const termSetPath = request.method(store, 'GetTermSet', guid(termSet.Id))
      return {
        termSet,
        found: request.instantiate(termSetPath),
        read: request.query(termSetPath, ['CustomSortOrder']),
        terms: (termSet.Terms || []).map((term) => {
          const termPath = request.method(store, 'GetTerm', guid(term.Id))
          return {
            term,
            found: request.instantiate(termPath),
            read:
              termSet.UpdateExistingTerms && hasProperties(term)
                ? request.query(termPath, ['CustomProperties'])
                : undefined
          }
        })
      }
    })
    let results: Record<number, any>
    try {
      results = await executeQuery(web, request)
    } catch (error) {
      throw new Error(
        `Failed to read the default term store: ${describeCsomError(error)}`
      )
    }
    if (isNull(results[storeFound]))
      throw new Error('No default site collection term store is available.')
    const state: ITermStoreState = {
      defaultLanguage: results[storeRead] && results[storeRead].DefaultLanguage,
      groupExists: !isNull(results[groupFound]),
      sets: {}
    }
    for (const read of reads) {
      const terms: Record<string, ITermState> = {}
      for (const term of read.terms)
        terms[term.term.Id] = {
          exists: !isNull(results[term.found]),
          customProperties:
            (term.read !== undefined &&
              results[term.read] &&
              results[term.read].CustomProperties) ||
            {}
        }
      state.sets[read.termSet.Id] = {
        exists: !isNull(results[read.found]),
        customSortOrder:
          (results[read.read] && results[read.read].CustomSortOrder) || '',
        terms
      }
    }
    super.log_info(
      'read',
      `Default term store: language ${state.defaultLanguage}, term group ${
        state.groupExists ? 'exists' : 'is missing'
      }`
    )
    return state
  }

  /**
   * Creates the term group with the template's id.
   *
   * @param web - The web
   * @param taxonomy - The taxonomy part of the template
   */
  private async createGroup(web: IWeb, taxonomy: ITaxonomy): Promise<void> {
    const { Id, Name } = taxonomy.TermGroup
    super.log_info('ensureGroup', `Creating term group ${Name}`)
    const request = new CsomRequest()
    const store = this.termStore(request)
    request.instantiate(request.method(store, 'CreateGroup', Name, guid(Id)))
    request.call(store, 'CommitAll')
    try {
      await executeQuery(web, request)
    } catch (error) {
      throw new Error(
        `Failed to create term group ${Name} (${guidOf(
          Id
        )}): ${describeCsomError(error)}`
      )
    }
  }

  /**
   * Writes what a term set needs, in one request that ends with `CommitAll`; nothing when it needs
   * nothing.
   *
   * @param web - The web
   * @param taxonomy - The taxonomy part of the template
   * @param termSet - The term set
   * @param current - What the term store has of it
   * @param lcid - The term store's default language
   */
  private async processTermSet(
    web: IWeb,
    taxonomy: ITaxonomy,
    termSet: ITermSet,
    current: ITermSetState,
    lcid: number
  ): Promise<void> {
    const request = new CsomRequest()
    const store = this.termStore(request)
    let termSetPath: number
    if (current.exists) {
      super.log_info('ensureTermSet', `Term set ${termSet.Name} already exists`)
      termSetPath = request.method(store, 'GetTermSet', guid(termSet.Id))
    } else {
      super.log_info('ensureTermSet', `Creating term set ${termSet.Name}`)
      termSetPath = request.method(
        request.method(store, 'GetGroup', guid(taxonomy.TermGroup.Id)),
        'CreateTermSet',
        termSet.Name,
        guid(termSet.Id),
        lcid
      )
      request.instantiate(termSetPath)
      if (termSet.Description)
        request.set(termSetPath, 'Description', termSet.Description)
      if (typeof termSet.IsOpenForTermCreation === 'boolean')
        request.set(
          termSetPath,
          'IsOpenForTermCreation',
          termSet.IsOpenForTermCreation
        )
    }

    const terms = termSet.Terms || []
    for (const term of terms) {
      const stored = current.terms[term.Id]
      if (!stored.exists) {
        super.log_info('ensureTerms', `Creating term ${term.Name}`)
        const termPath = request.method(
          termSetPath,
          'CreateTerm',
          term.Name,
          lcid,
          guid(term.Id)
        )
        request.instantiate(termPath)
        for (const key of Object.keys(term.CustomProperties || {}))
          request.call(
            termPath,
            'SetCustomProperty',
            key,
            term.CustomProperties[key]
          )
        continue
      }
      const missing = termSet.UpdateExistingTerms
        ? Object.keys(term.CustomProperties || {}).filter(
            (key) => stored.customProperties[key] === undefined
          )
        : []
      if (missing.length === 0) {
        super.log_info('ensureTerms', `Term ${term.Name} already exists`)
        continue
      }
      super.log_info(
        'ensureTerms',
        `Adding missing custom properties to existing term ${
          term.Name
        }: ${missing.join(', ')}`
      )
      const termPath = request.method(store, 'GetTerm', guid(term.Id))
      for (const key of missing)
        request.call(
          termPath,
          'SetCustomProperty',
          key,
          term.CustomProperties[key]
        )
    }

    const order = sortOrder(terms)
    if (order && sortKey(order) !== sortKey(current.customSortOrder)) {
      super.log_info(
        'applySortOrder',
        `Setting the custom sort order of term set ${termSet.Name}`
      )
      request.set(termSetPath, 'CustomSortOrder', order)
    }

    if (request.actionCount === 0) {
      super.log_info('processTermSet', `Term set ${termSet.Name} is unchanged.`)
      return
    }
    request.call(store, 'CommitAll')
    try {
      await executeQuery(web, request)
    } catch (error) {
      throw new Error(
        `Failed to provision term set ${termSet.Name} (${guidOf(
          termSet.Id
        )}): ${describeCsomError(error)}`
      )
    }
  }
}

/**
 * Whether a term has custom properties in the template.
 *
 * @param term - The term
 */
function hasProperties(term: ITerm): boolean {
  return (
    !!term.CustomProperties && Object.keys(term.CustomProperties).length > 0
  )
}

/**
 * The custom sort order the template's `SortOrder` values give (`id:id:…`), or `undefined` when
 * fewer than two terms have one.
 *
 * @param terms - The term set's terms
 */
function sortOrder(terms: ITerm[]): string | undefined {
  const ordered = terms
    .filter((term) => typeof term.SortOrder === 'number')
    .sort((a, b) => a.SortOrder - b.SortOrder)
  if (ordered.length > 1)
    return ordered.map((term) => guidOf(term.Id)).join(':')
}
