import { IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import {
  declaredType,
  fieldAttributes,
  fieldTarget,
  guidOf,
  mergeFieldXml,
  sameFieldXml,
  typeChange
} from '../util/fieldxml'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'

/**
 * A site column as the handler reads it; `TypeAsString` and `SchemaXml` only for the web's own.
 */
export interface ISiteField {
  Id: string
  InternalName: string
  TypeAsString?: string
  SchemaXml?: string
}

/**
 * What SharePoint stores as a taxonomy column's `SspId` when the one sent is not a GUID (a literal
 * `{sitecollectiontermstoreid}`), probe SF3: a taxonomy target holding it is unset.
 */
const EMPTY_GUID = '00000000-0000-0000-0000-000000000000'

/**
 * A taxonomy column's targets. A `SchemaXml` MERGE ignores its whole `Customization` (probe SF3),
 * so they are set as `TaxonomyField` properties.
 */
const TAXONOMY_TARGETS = ['SspId', 'TermSetId']

/**
 * What SharePoint manages on a taxonomy column, kept as stored so an unchanged column compares as
 * the same: the `Customization` (a MERGE ignores it), and the hidden list, its web and the
 * `ShowField` it sets in the site's language (Term1033 becomes Term1044).
 */
const TAXONOMY_MANAGED = ['Customization', 'List', 'WebId', 'ShowField']

/**
 * The message of what a request rejected with.
 *
 * @param error - The error
 */
const messageOf = (error: any): string => (error && error.message) || error

/**
 * Describes the Site Fields Object Handler: site columns from field XML, created when missing and
 * updated in place with the rules list fields have (docs/specs/sitefields-re-apply.md). A column
 * that fails is logged and the others are still processed.
 */
export class SiteFields extends HandlerBase {
  private context: ProvisioningContext
  private tokenHelper: TokenHelper

  /**
   * Creates a new instance of the ObjectSiteFields class
   */
  constructor(config: IProvisioningConfig) {
    super('SiteFields', config)
  }

  /**
   * Provisioning site fields: the web's columns are read once, then each template column is
   * created, updated in place, or left as it is.
   *
   * @param web - The web
   * @param siteFields - The site fields
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: IWeb,
    siteFields: string[],
    context?: ProvisioningContext
  ): Promise<void> {
    this.context = context
    this.tokenHelper = new TokenHelper(this.context, this.config)
    super.scope_started()
    try {
      const read = await this.readSiteFields(web, siteFields)
      if (read) {
        const [own, available] = read
        this.context.siteFields = own.reduce((object, field) => {
          object[field.InternalName] = field.Id
          return object
        }, {})
        for (const schemaXml of siteFields)
          await this.processSiteField(web, schemaXml, own, available)
      }
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Reads the web's own columns (`Id`, `InternalName`, `TypeAsString`, `SchemaXml`) and the columns
   * available to it, which on a subweb include its parent webs' (`web.fields` lists only the web's
   * own). Without `top()`: the collection is not paged, and `$top` truncates it (probe SF4).
   *
   * @param web - The web
   * @param siteFields - The template's columns, named when the read fails
   */
  private async readSiteFields(
    web: IWeb,
    siteFields: string[]
  ): Promise<[ISiteField[], ISiteField[]] | undefined> {
    try {
      const [own, available] = await Promise.all([
        web.fields.select('Id', 'InternalName', 'TypeAsString', 'SchemaXml')<
          ISiteField[]
        >(),
        web.availablefields.select('Id', 'InternalName')<ISiteField[]>()
      ])
      return [own, available]
    } catch (error) {
      const names = siteFields.map((xml) => {
        try {
          return fieldAttributes(xml).Name
        } catch {
          return '?'
        }
      })
      super.log_error(
        'readSiteFields',
        `Failed to read the site columns of the web, so none is created or changed (${names.join(
          ', '
        )}): ${messageOf(error)}`
      )
      return
    }
  }

  /**
   * Provisions a site column, identified by its internal name (`Name`):
   *
   * - on the web: updated in place (`updateSiteField`); another id is a warning, the stored id stays;
   * - not on the web, but its id is, under another name: an error, not created (SharePoint would
   *   refuse it with "The local device name is already in use");
   * - only in a parent web: a warning, neither created nor changed (the template changes no other
   *   web than the one it is applied to);
   * - nowhere: created.
   *
   * A known token that cannot be resolved is an error naming the column; the column is still
   * written with the token as it stands, except into a target it already has.
   *
   * @param web - The web
   * @param schemaXml - The template's field XML
   * @param own - The web's own columns
   * @param available - The columns available to the web, its parents' included
   */
  private async processSiteField(
    web: IWeb,
    schemaXml: string,
    own: ISiteField[],
    available: ISiteField[]
  ): Promise<void> {
    const template = this.tokenHelper.replaceTokens(schemaXml)
    let attributes: Record<string, string>
    try {
      attributes = fieldAttributes(template)
    } catch (error) {
      super.log_error(
        'processSiteField',
        `Failed to read a site column's XML, it is skipped: ${messageOf(error)}`
      )
      return
    }
    const { Name: name, DisplayName: displayName, ID, Type: type } = attributes
    const id = guidOf(ID)
    const label = `${name} (${displayName})`
    const unresolved = this.tokenHelper.unresolvedTokens(schemaXml)
    if (unresolved.length > 0)
      super.log_error(
        'processSiteField',
        `Site column ${label}: ${unresolved.join(
          ', '
        )} could not be resolved. It is written as it stands, except into a list, web, term store or term set the column already points at.`
      )
    if ((type || '').startsWith('TaxonomyFieldType'))
      for (const target of TAXONOMY_TARGETS) {
        const value = fieldTarget(template, target)
        if (
          value !== undefined &&
          !guidOf(value) &&
          this.tokenHelper.unresolvedTokens(value).length === 0
        )
          super.log_error(
            'processSiteField',
            `Site column ${label}: its ${target} ${value} is not a GUID, so it is not set (SharePoint stores the empty GUID for it).`
          )
      }
    const stored = own.find((field) => field.InternalName === name)
    try {
      if (stored) {
        await this.updateSiteField(web, label, template, attributes, stored)
        return
      }
      const sameId = id && own.find((field) => guidOf(field.Id) === id)
      if (sameId) {
        super.log_error(
          'processSiteField',
          `Site column ${label} is not created: the web has its id ${id} as the column ${sameId.InternalName}.`
        )
        return
      }
      const inParent =
        available.find((field) => field.InternalName === name) ||
        (id && available.find((field) => guidOf(field.Id) === id))
      if (inParent) {
        const message = `Site column ${label} is not created: a parent web has ${
          inParent.InternalName === name
            ? 'it'
            : `its id ${id} as the column ${inParent.InternalName}`
        }, and the template changes no other web than this one.`
        if (inParent.InternalName === name)
          super.log_warn('processSiteField', message)
        else super.log_error('processSiteField', message)
        return
      }
      super.log_info('processSiteField', `Adding site column ${label}.`)
      const created = await web.fields.createFieldAsXml(template)
      const createdId = guidOf(created && created.Id) || id
      own.push({
        Id: createdId,
        InternalName: name,
        TypeAsString: declaredType(attributes),
        SchemaXml: template
      })
      this.context.siteFields[name] = createdId
      this.context.createdSiteFields[name] = true
      super.log_info('processSiteField', `Site column ${label} added.`)
    } catch (error) {
      super.log_error(
        'processSiteField',
        `Failed to ${
          stored ? 'update' : 'create'
        } site column ${label}: ${messageOf(error)}`
      )
    }
  }

  /**
   * Updates a site column in place with the rules list fields have (`Lists.processField`): the
   * template merged onto the stored `SchemaXml` (attributes and children it leaves out are kept:
   * SharePoint would drop them, probe SF2) and sent only when it differs; a widening type change is
   * applied, another is a warning and the column is left as it is; a target follows `targetRule`.
   * Nothing is pushed to the column's list copies: a `SchemaXml` MERGE does not reach them (probes
   * SF1, SF7).
   *
   * @param web - The web
   * @param label - The column, for the log
   * @param template - The template's field XML, tokens replaced
   * @param attributes - Its attributes
   * @param stored - The column on the web
   */
  private async updateSiteField(
    web: IWeb,
    label: string,
    template: string,
    attributes: Record<string, string>,
    stored: ISiteField
  ): Promise<void> {
    const id = guidOf(attributes.ID)
    if (id && guidOf(stored.Id) !== id)
      super.log_warn(
        'updateSiteField',
        `Site column ${label} has the id ${stored.Id} on the web, the template says ${id}: it keeps its id, and a list FieldRef or content type that names ${id} does not find it.`
      )
    const type = declaredType(attributes)
    const change = typeChange(stored.TypeAsString, type)
    if (change === 'other') {
      super.log_warn(
        'updateSiteField',
        `Site column ${label} is ${stored.TypeAsString}, the template says ${type}: only a widening change is applied, so the column is left as it is.`
      )
      return
    }
    const taxonomy = (stored.TypeAsString || '').startsWith('TaxonomyFieldType')
    const keep = taxonomy ? [...TAXONOMY_MANAGED] : []
    const properties: Record<string, string> = {}
    for (const target of taxonomy ? TAXONOMY_TARGETS : ['List', 'WebId']) {
      const rule = this.targetRule(label, target, stored, template)
      if (rule === 'keep') keep.push(target)
      else if (rule === 'follow' && taxonomy) {
        const value = guidOf(fieldTarget(template, target))
        if (value) properties[target] = value
      }
    }
    const merged = mergeFieldXml(stored.SchemaXml, template, keep)
    const field = () => web.fields.getById(guidOf(stored.Id))
    const changed = !sameFieldXml(merged, stored.SchemaXml)
    if (changed) await field().update({ SchemaXml: merged }, 'SP.Field')
    if (Object.keys(properties).length > 0)
      await field().update(properties, 'SP.Taxonomy.TaxonomyField')
    if (!changed && Object.keys(properties).length === 0)
      super.log_info('updateSiteField', `Site column ${label} is unchanged.`)
    else
      super.log_info(
        'updateSiteField',
        `Site column ${label} updated in place${
          change === 'widening' ? ` (${stored.TypeAsString} to ${type})` : ''
        }${
          Object.keys(properties).length > 0
            ? ` (${Object.keys(properties).join(', ')} set)`
            : ''
        }.`
      )
  }

  /**
   * What happens to one target of a site column the web has (a lookup's `List` or `WebId`, a
   * taxonomy column's `SspId` or `TermSetId`) when the template says another:
   *
   * - `same`: the template leaves it out or says what is stored;
   * - `keep`: the template's value still holds an unresolved token, or is no GUID for a taxonomy
   *   target (both already logged), or the column points at another list or term set it was not
   *   made with in this setup (a warning: new list copies would differ from those made before);
   * - `follow`: the stored value is not a GUID (a token an earlier run stored literally) or the empty
   *   GUID a literal term store id leaves, so the column heals; or the column was made in this setup.
   *
   * @param label - The column, for the log
   * @param target - The target
   * @param stored - The column on the web
   * @param template - The template's field XML, tokens replaced
   */
  private targetRule(
    label: string,
    target: string,
    stored: ISiteField,
    template: string
  ): 'same' | 'keep' | 'follow' {
    const before = fieldTarget(stored.SchemaXml, target)
    const after = fieldTarget(template, target)
    const comparable = (value: string) => guidOf(value) || value
    if (
      after === undefined ||
      (before !== undefined && comparable(before) === comparable(after))
    )
      return 'same'
    if (this.tokenHelper.unresolvedTokens(after).length > 0) return 'keep'
    if (TAXONOMY_TARGETS.includes(target) && !guidOf(after)) return 'keep'
    const storedId = guidOf(before)
    if (
      !storedId ||
      (storedId === EMPTY_GUID && TAXONOMY_TARGETS.includes(target))
    ) {
      if (before !== undefined)
        super.log_info(
          'updateSiteField',
          `Site column ${label} has ${target} ${before}, which is not one: it follows the template.`
        )
      return 'follow'
    }
    if (this.context.createdSiteFields[stored.InternalName]) {
      super.log_info(
        'updateSiteField',
        `Site column ${label} points at another ${target} than the template says; it was made in this setup, so it follows the template.`
      )
      return 'follow'
    }
    super.log_warn(
      'updateSiteField',
      `Site column ${label} points at another ${target} (${before}) than the template says (${after}): it keeps it, as the list copies made from it do.`
    )
    return 'keep'
  }
}
