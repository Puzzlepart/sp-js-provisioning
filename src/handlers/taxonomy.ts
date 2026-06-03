/* eslint-disable unicorn/prevent-abbreviations */
import { IWeb } from '@pnp/sp/presets/all'
import initSpfxJsom, { ExecuteJsomQuery, JsomContext } from 'spfx-jsom'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { ITaxonomy, ITerm, ITermGroup, ITermSet } from '../schema'
import { HandlerBase } from './handlerbase'

/**
 * Describes the Taxonomy Object Handler.
 *
 * Provisions a term group, its term sets and terms into the default site
 * collection term store using JSOM (so explicit term-set/term GUIDs from the
 * template are preserved). Idempotent: anything that already exists by ID is
 * left untouched.
 */
export class Taxonomy extends HandlerBase {
  public jsomContext: JsomContext
  public context: ProvisioningContext
  private termStore: SP.Taxonomy.TermStore

  /**
   * Creates a new instance of the Taxonomy handler.
   */
  constructor(config: IProvisioningConfig) {
    super('Taxonomy', config)
  }

  /**
   * Provisioning taxonomy
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
    const spfxJsom = await initSpfxJsom(context.web.ServerRelativeUrl, {
      loadTaxonomy: true
    })
    this.jsomContext = spfxJsom.jsomContext
    this.termStore = spfxJsom.defaultTermStore
    super.scope_started()
    try {
      if (!this.termStore) {
        throw new Error('No default site collection term store is available.')
      }
      this.jsomContext.clientContext.load(this.termStore)
      await ExecuteJsomQuery(this.jsomContext)
      const lcid = this.termStore.get_defaultLanguage()

      const group = await this.ensureGroup(taxonomy.TermGroup)
      for (const termSetDef of taxonomy.TermSets ?? []) {
        const termSet = await this.ensureTermSet(group, termSetDef, lcid)
        await this.ensureTerms(termSet, termSetDef, lcid)
      }
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Ensure the term group exists, creating it with the specified ID if missing.
   *
   * @param groupDef - Term group definition
   */
  private async ensureGroup(
    groupDef: ITermGroup
  ): Promise<SP.Taxonomy.TermGroup> {
    const groupId = new SP.Guid(groupDef.Id)
    const existing = this.termStore.getGroup(groupId)
    if (await this.exists(existing)) {
      super.log_info('ensureGroup', `Term group ${groupDef.Name} already exists`)
      return existing
    }
    super.log_info('ensureGroup', `Creating term group ${groupDef.Name}`)
    const group = this.termStore.createGroup(groupDef.Name, groupId)
    this.termStore.commitAll()
    await ExecuteJsomQuery(this.jsomContext)
    return group
  }

  /**
   * Ensure the term set exists under the group, creating it with the specified
   * ID if missing.
   *
   * @param group - Parent term group
   * @param termSetDef - Term set definition
   * @param lcid - Default language of the term store
   */
  private async ensureTermSet(
    group: SP.Taxonomy.TermGroup,
    termSetDef: ITermSet,
    lcid: number
  ): Promise<SP.Taxonomy.TermSet> {
    const termSetId = new SP.Guid(termSetDef.Id)
    const existing = this.termStore.getTermSet(termSetId)
    if (await this.exists(existing)) {
      super.log_info(
        'ensureTermSet',
        `Term set ${termSetDef.Name} already exists`
      )
      return existing
    }
    super.log_info('ensureTermSet', `Creating term set ${termSetDef.Name}`)
    const termSet = group.createTermSet(termSetDef.Name, termSetId, lcid)
    if (termSetDef.Description) {
      termSet.set_description(termSetDef.Description)
    }
    if (typeof termSetDef.IsOpenForTermCreation === 'boolean') {
      termSet.set_isOpenForTermCreation(termSetDef.IsOpenForTermCreation)
    }
    this.termStore.commitAll()
    await ExecuteJsomQuery(this.jsomContext)
    return termSet
  }

  /**
   * Ensure all terms exist in the term set (creating missing ones with their
   * specified IDs and custom properties) and apply the custom sort order.
   *
   * @param termSet - Parent term set
   * @param termSetDef - Term set definition
   * @param lcid - Default language of the term store
   */
  private async ensureTerms(
    termSet: SP.Taxonomy.TermSet,
    termSetDef: ITermSet,
    lcid: number
  ): Promise<void> {
    const terms = termSetDef.Terms ?? []
    let created = false
    for (const termDef of terms) {
      const termId = new SP.Guid(termDef.Id)
      const existing = this.termStore.getTerm(termId)
      if (await this.exists(existing)) {
        super.log_info('ensureTerms', `Term ${termDef.Name} already exists`)
        continue
      }
      super.log_info('ensureTerms', `Creating term ${termDef.Name}`)
      const term = termSet.createTerm(termDef.Name, lcid, termId)
      this.applyCustomProperties(term, termDef)
      created = true
    }
    if (created) {
      this.termStore.commitAll()
      await ExecuteJsomQuery(this.jsomContext)
    }
    await this.applySortOrder(termSet, terms)
  }

  /**
   * Apply custom properties to a newly created term.
   *
   * @param term - The term
   * @param termDef - Term definition
   */
  private applyCustomProperties(term: SP.Taxonomy.Term, termDef: ITerm): void {
    if (!termDef.CustomProperties) return
    for (const key of Object.keys(termDef.CustomProperties)) {
      term.setCustomProperty(key, termDef.CustomProperties[key])
    }
  }

  /**
   * Apply the custom sort order of a term set based on the `SortOrder` of its
   * terms.
   *
   * @param termSet - The term set
   * @param terms - Term definitions
   */
  private async applySortOrder(
    termSet: SP.Taxonomy.TermSet,
    terms: ITerm[]
  ): Promise<void> {
    const ordered = terms
      .filter((term) => typeof term.SortOrder === 'number')
      .sort((a, b) => a.SortOrder - b.SortOrder)
    if (ordered.length < 2) return
    termSet.set_customSortOrder(ordered.map((term) => term.Id).join(':'))
    this.termStore.commitAll()
    await ExecuteJsomQuery(this.jsomContext)
  }

  /**
   * Load a taxonomy client object and return whether it exists on the server.
   *
   * @param clientObject - The taxonomy client object (group/set/term)
   */
  private async exists(clientObject: SP.ClientObject): Promise<boolean> {
    this.jsomContext.clientContext.load(clientObject)
    try {
      await ExecuteJsomQuery(this.jsomContext)
      return !clientObject.get_serverObjectIsNull()
    } catch {
      return false
    }
  }
}
