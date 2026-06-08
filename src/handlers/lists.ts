/* eslint-disable unicorn/no-array-for-each */
import { IList, IWeb } from '@pnp/sp/presets/all'
import initSpfxJsom, { JsomContext } from 'spfx-jsom'
import * as xmljs from 'xml-js'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import {
  IContentTypeBinding,
  IDataRow,
  IListInstance,
  IListInstanceFieldReference,
  IListView
} from '../schema'
import { addFieldAttributes } from '../util'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'

export interface ISPField {
  Id: string
  InternalName: string
  SchemaXml: string
}

/**
 * Describes the Lists Object Handler
 */
export class Lists extends HandlerBase {
  public tokenHelper: TokenHelper
  public jsomContext: JsomContext
  public context: ProvisioningContext

  /**
   * Creates a new instance of the Lists class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super('Lists', config)
  }

  /**
   * Provisioning lists
   *
   * @param web - The web
   * @param lists - The lists to provision
   */
  public async ProvisionObjects(
    web: IWeb,
    lists: IListInstance[],
    context: ProvisioningContext
  ): Promise<void> {
    this.jsomContext = (
      await initSpfxJsom(context.web.ServerRelativeUrl)
    ).jsomContext
    this.context = context
    this.tokenHelper = new TokenHelper(this.context, this.config)
    super.scope_started()
    try {
      this.context.lists = (
        await web.lists.select('Id', 'Title')<
          Array<{ Id: string; Title: string }>
        >()
      ).reduce((object, l) => {
        object[l.Title] = l.Id
        return object
      }, {})
      await lists.reduce(
        (chain: any, list) => chain.then(() => this.processList(web, list)),
        Promise.resolve()
      )
      await lists.reduce(
        (chain: any, list) =>
          chain.then(() => this.processListFields(web, list)),
        Promise.resolve()
      )
      await lists.reduce(
        (chain: any, list) =>
          chain.then(() => this.processListFieldRefs(web, list)),
        Promise.resolve()
      )
      await lists.reduce(
        (chain: any, list) =>
          chain.then(() => this.processListViews(web, list)),
        Promise.resolve()
      )
      await lists.reduce(
        (chain: any, list) =>
          chain.then(() => this.processListDataRows(web, list)),
        Promise.resolve()
      )
      this.context.lists = (
        await web.lists.select('Id', 'Title')<
          Array<{ Id: string; Title: string }>
        >()
      ).reduce((object, l) => {
        object[l.Title] = l.Id
        return object
      }, {})
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Processes a list
   *
   * @param web - The web
   * @param lc - The list
   */
  private async processList(web: IWeb, lc: IListInstance): Promise<void> {
    super.log_info('processList', `Processing list ${lc.Title}`)
    let list: IList
    if (this.context.lists[lc.Title]) {
      super.log_info(
        'processList',
        `List ${lc.Title} already exists. Ensuring...`
      )
      const listEnsure = await web.lists.ensure(
        lc.Title,
        lc.Description,
        lc.Template,
        lc.ContentTypesEnabled,
        lc.AdditionalSettings
      )
      list = listEnsure.list
    } else {
      super.log_info(
        'processList',
        `List ${lc.Title} doesn't exist. Creating...`
      )
      const listAdd = await web.lists.add(
        lc.Title,
        lc.Description,
        lc.Template,
        lc.ContentTypesEnabled,
        lc.AdditionalSettings
      )
      list = listAdd.list
      this.context.lists[listAdd.data.Title] = listAdd.data.Id
    }
    if (lc.ContentTypeBindings) {
      await this.processContentTypeBindings(
        lc,
        list,
        lc.ContentTypeBindings,
        lc.RemoveExistingContentTypes
      )
    }
  }

  /**
   * Processes content type bindings for a list
   *
   * @param lc - The list configuration
   * @param list - The pnp list
   * @param contentTypeBindings - Content type bindings
   * @param removeExisting - Remove existing content type bindings
   */
  private async processContentTypeBindings(
    lc: IListInstance,
    list: IList,
    contentTypeBindings: IContentTypeBinding[],
    removeExisting: boolean
  ): Promise<any> {
    super.log_info(
      'processContentTypeBindings',
      `Processing content types for list ${lc.Title}.`
    )
    await contentTypeBindings.reduce(
      (chain, ct) =>
        chain.then(() =>
          this.processContentTypeBinding(lc, list, ct.ContentTypeID)
        ),
      Promise.resolve()
    )
    if (removeExisting) {
      const promises = []
      const contentTypes = await list.contentTypes()
      contentTypes.forEach(({ Id: { StringValue: ContentTypeId } }) => {
        const shouldRemove =
          contentTypeBindings.filter((ctb) =>
            ContentTypeId.includes(ctb.ContentTypeID)
          ).length === 0 && !ContentTypeId.includes('0x0120')
        if (shouldRemove) {
          super.log_info(
            'processContentTypeBindings',
            `Removing content type ${ContentTypeId} from list ${lc.Title}`
          )
          promises.push(list.contentTypes.getById(ContentTypeId).delete())
        } else {
          super.log_info(
            'processContentTypeBindings',
            `Skipping removal of content type ${ContentTypeId} from list ${lc.Title}`
          )
        }
      })
      await Promise.all(promises)
    }
  }

  /**
   * Processes a content type binding for a list
   *
   * @param lc - The list configuration
   * @param list - The pnp list
   * @param contentTypeID - The Content Type ID
   */
  private async processContentTypeBinding(
    lc: IListInstance,
    list: IList,
    contentTypeID: string
  ): Promise<any> {
    try {
      super.log_info(
        'processContentTypeBinding',
        `Adding content Type ${contentTypeID} to list ${lc.Title}.`
      )
      await list.contentTypes.addAvailableContentType(contentTypeID)
      super.log_info(
        'processContentTypeBinding',
        `Content Type ${contentTypeID} added successfully to list ${lc.Title}.`
      )
    } catch {
      super.log_info(
        'processContentTypeBinding',
        `Failed to add Content Type ${contentTypeID} to list ${lc.Title}.`
      )
    }
  }

  /**
   * Processes fields for a list
   *
   * @param web - The web
   * @param list - The pnp list
   */
  private async processListFields(
    web: IWeb,
    list: IListInstance
  ): Promise<any> {
    if (list.Fields) {
      await list.Fields.reduce(
        (chain, field) => chain.then(() => this.processField(web, list, field)),
        Promise.resolve()
      )
    }
  }

  /**
   * Processes a field for a lit
   *
   * @param web - The web
   * @param lc - The list configuration
   * @param fieldXml - Field XML
   */
  private async processField(
    web: IWeb,
    lc: IListInstance,
    fieldXml: string
  ): Promise<any> {
    const list = web.lists.getByTitle(lc.Title)
    const fieldXmlJson = JSON.parse(xmljs.xml2json(fieldXml))
    const fieldAttribute = fieldXmlJson.elements[0].attributes
    const fieldName = fieldAttribute.Name
    const fieldDisplayName = fieldAttribute.DisplayName

    super.log_info(
      'processField',
      `Processing field ${fieldName} (${fieldDisplayName}) for list ${lc.Title}.`
    )
    fieldXmlJson.elements[0].attributes.DisplayName = fieldName
    fieldXml = xmljs.json2xml(fieldXmlJson)

    // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
    try {
      await list.fields.getById(fieldAttribute.ID).delete()
      super.log_info(
        'processField',
        `Field ${fieldName} (${fieldDisplayName}) successfully deleted from list ${lc.Title}.`
      )
    } catch {
      super.log_info(
        'processField',
        `Field ${fieldName} (${fieldDisplayName}) does not exist in list ${lc.Title}.`
      )
    }

    // Looks like e.g. lookup fields can't be updated, so we'll need to re-create the field
    try {
      const fieldAddResult = await list.fields.createFieldAsXml(
        this.tokenHelper.replaceTokens(fieldXml)
      )
      await fieldAddResult.field.update({ Title: fieldDisplayName })
      super.log_info(
        'processField',
        `Field '${fieldDisplayName}' added successfully to list ${lc.Title}.`
      )
    } catch {
      super.log_info(
        'processField',
        `Failed to add field '${fieldDisplayName}' to list ${lc.Title}.`
      )
    }
  }

  /**
   * Processes field refs for a list
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async processListFieldRefs(
    web: IWeb,
    lc: IListInstance
  ): Promise<any> {
    if (lc.FieldRefs) {
      super.log_info(
        'processListFieldRefs',
        `Retrieving fields for list ${lc.Title} and web.`
      )
      const list = web.lists.getByTitle(lc.Title)
      const [listFields, webFields] = await Promise.all([
        list.fields.select('Id', 'InternalName', 'SchemaXml')<ISPField[]>(),
        web.fields.select('Id', 'InternalName', 'SchemaXml')<ISPField[]>()
      ])
      super.log_info(
        'processListFieldRefs',
        `Fields for list ${lc.Title} and web retrieved. Processing field refs.`
      )
      await await lc.FieldRefs.reduce((chain: any, fieldReference) => {
        return chain.then(() =>
          this.processFieldRef(list, lc, fieldReference, listFields, webFields)
        )
      }, Promise.resolve())
    }
  }

  /**
   *
   * Processes a field ref for a list
   *
   * @param list - The list
   * @param lc - The list configuration
   * @param fieldRef - The list field ref
   * @param listFields - The list fields
   * @param webFields - The web fields
   */
  private async processFieldRef(
    list: IList,
    lc: IListInstance,
    fieldReference: IListInstanceFieldReference,
    listFields: ISPField[],
    webFields: ISPField[]
  ): Promise<void> {
    const listFld = listFields.find((f) => f.Id === fieldReference.ID)
    const webFld = webFields.find((f) => f.Id === fieldReference.ID)
    super.log_info(
      'processFieldRef',
      `Processing field ref '${fieldReference.ID}' for list ${lc.Title}.`
    )
    if (Boolean(listFld)) {
      await list.fields.getById(fieldReference.ID).update({
        Hidden: fieldReference.Hidden,
        Required: fieldReference.Required,
        Title: fieldReference.DisplayName
      })
      if (fieldReference.AdditionalProperties) {
        const schemaXml = addFieldAttributes(listFld.SchemaXml, {
          Hidden: fieldReference.Hidden,
          Required: fieldReference.Required,
          DisplayName: fieldReference.Name,
          ...(fieldReference.AdditionalProperties ?? {})
        })
        super.log_info(
          'processFieldRef',
          `Additional properties set for field ref '${fieldReference.ID}' for list ${lc.Title}. Attempting to generate schema XML.`,
          { schemaXml }
        )
        await list.fields
          .getById(fieldReference.ID)
          .update({ SchemaXml: schemaXml })
      }
      super.log_info(
        'processFieldRef',
        `Field '${fieldReference.ID}' updated for list ${lc.Title}.`
      )
    } else if (Boolean(webFld)) {
      super.log_info(
        'processFieldRef',
        `Adding field '${fieldReference.ID}' to list ${lc.Title}.`
      )
      const schemaXml = addFieldAttributes(webFld.SchemaXml, {
        DisplayName: fieldReference.Name,
        SourceID: `{${this.context.lists[lc.Title]}}`,
        ...(fieldReference.AdditionalProperties ?? {})
      })
      const fieldAddResult = await list.fields.createFieldAsXml(schemaXml)
      await fieldAddResult.field.update({
        Title: fieldReference.DisplayName,
        Required: fieldReference.Required,
        Hidden: fieldReference.Hidden
      })
      super.log_info(
        'processFieldRef',
        `Field '${fieldReference.ID}' added from web.`
      )
    }
  }

  /**
   * Processes views for a list
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async processListViews(web: IWeb, lc: IListInstance): Promise<any> {
    if (lc.Views) {
      if (lc.RemoveExistingViews) {
        await this.removeExistingViews(web, lc)
      }
      await lc.Views.reduce(
        (chain: any, view) => chain.then(() => this.processView(web, lc, view)),
        Promise.resolve()
      )
    }
    this.context.listViews = (
      await web.lists.getByTitle(lc.Title).views.select('Id', 'Title')<
        Array<{ Id: string; Title: string }>
      >()
    ).reduce((object, view) => {
      object[`${lc.Title}|${view.Title}`] = view.Id
      return object
    }, this.context.listViews)
  }

  /**
   * Removes existing views for a list
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async removeExistingViews(
    web: IWeb,
    lc: IListInstance
  ): Promise<void> {
    const views = await web.lists
      .getByTitle(lc.Title)
      .views.select('Id', 'Title')<Array<{ Id: string; Title: string }>>()
    super.log_info(
      '_removeExistingViews',
      `Removing existing views for list ${lc.Title}.`,
      views.map((view) => view.Title)
    )
    const promises = views.map((view) =>
      web.lists.getByTitle(lc.Title).views.getById(view.Id).delete()
    )
    await Promise.all(promises)
    super.log_info(
      '_removeExistingViews',
      `Existing views removed for list ${lc.Title}.`
    )
  }

  /**
   * Processes a view for a list
   *
   * @param web - The web
   * @param lc - The list configuration
   * @param lvc - The view configuration
   */
  private async processView(
    web: IWeb,
    lc: IListInstance,
    lvc: IListView
  ): Promise<void> {
    super.log_info(
      'processView',
      `Processing view ${lvc.Title} for list ${lc.Title}.`
    )
    const existingView = web.lists
      .getByTitle(lc.Title)
      .views.getByTitle(lvc.Title)
    let viewExists = false
    try {
      await existingView()
      viewExists = true
    } catch {}
    try {
      if (viewExists) {
        super.log_info(
          'processView',
          `View ${lvc.Title} for list ${lc.Title} already exists, updating.`
        )
        await existingView.update(lvc.AdditionalSettings)
        super.log_info(
          'processView',
          `View ${lvc.Title} successfully updated for list ${lc.Title}.`
        )
        await this.processViewFields(existingView, lvc)
      } else {
        super.log_info(
          'processView',
          `View ${lvc.Title} for list ${lc.Title} doesn't exists, creating.`
        )
        await web.lists
          .getByTitle(lc.Title)
          .views.add(lvc.Title, lvc.PersonalView, lvc.AdditionalSettings)
        // SP REST's view-create endpoint silently ignores properties such as
        // CustomFormatter, ViewType2 and Scope. Re-apply AdditionalSettings
        // via update so newly created views match updated views. Re-resolve
        // the view by title rather than reusing the IViewAddResult, since the
        // returned reference is bound to the parent views collection and the
        // update can silently no-op on some PnPjs/SharePoint combinations.
        const newView = web.lists
          .getByTitle(lc.Title)
          .views.getByTitle(lvc.Title)
        try {
          await newView.update(lvc.AdditionalSettings)
        } catch (error) {
          super.log_warn(
            'processView',
            `Failed to re-apply AdditionalSettings on newly created view ${lvc.Title}: ${error}`
          )
        }
        super.log_info(
          'processView',
          `View ${lvc.Title} added successfully to list ${lc.Title}.`
        )
        await this.processViewFields(newView, lvc)
      }
    } catch (error) {
      super.log_error(
        'processView',
        `Failed to process view ${lvc.Title}: ${error}`
      )
    }
  }

  /**
   * Processes view fields for a view
   *
   * @param view - The pnp view
   * @param lvc - The view configuration
   */
  private async processViewFields(view: any, lvc: IListView): Promise<void> {
    try {
      super.log_info(
        'processViewFields',
        `Processing view fields for view ${lvc.Title}.`
      )
      await view.fields.removeAll()
      await lvc.ViewFields.reduce(
        (chain, viewField) => chain.then(() => view.fields.add(viewField)),
        Promise.resolve()
      )
      super.log_info(
        'processViewFields',
        `View fields successfully processed for view ${lvc.Title}.`
      )
    } catch {
      super.log_info(
        'processViewFields',
        `Failed to process view fields for view ${lvc.Title}.`
      )
    }
  }

  /**
   * Provisions seed data rows for a list. Runs last (after fields, content
   * types and views exist). Upserts by KeyColumn so re-running is idempotent,
   * and resolves Lookup/User/Taxonomy values. A failing row is logged and
   * skipped rather than aborting the whole list.
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async processListDataRows(web: IWeb, lc: IListInstance): Promise<void> {
    const dataRows = lc.DataRows
    if (!dataRows || !dataRows.Rows || dataRows.Rows.length === 0) return
    const { KeyColumn, UpdateBehavior = 'Overwrite', Rows } = dataRows
    super.log_info(
      'processListDataRows',
      `Provisioning ${Rows.length} data row(s) for list ${lc.Title}.`
    )
    const list = web.lists.getByTitle(lc.Title)
    const fields = await list.fields.select(
      'Id',
      'InternalName',
      'TypeAsString',
      'LookupList',
      'LookupField',
      'TextField'
    )<any[]>()
    const byName = new Map<string, any>(fields.map((f) => [f.InternalName, f]))
    const byId = new Map<string, any>(fields.map((f) => [Lists._normGuid(f.Id), f]))

    for (const [index, row] of Rows.entries()) {
      try {
        const values = await this._buildItemValues(web, byName, byId, row)
        let existingId: number
        const keyValue = KeyColumn ? row[KeyColumn] : undefined
        if (KeyColumn && keyValue !== undefined && keyValue !== null) {
          const matches = await list.items
            .filter(`${KeyColumn} eq '${Lists._escapeOData(String(keyValue))}'`)
            .select('Id')
            .top(1)<Array<{ Id: number }>>()
          existingId = matches[0] ? matches[0].Id : undefined
        }
        if (existingId !== undefined) {
          if (UpdateBehavior === 'Skip') {
            super.log_info(
              'processListDataRows',
              `Row ${index} already exists (${KeyColumn}=${keyValue}) — skipped.`
            )
            continue
          }
          await list.items.getById(existingId).update(values)
        } else {
          await list.items.add(values)
        }
      } catch (error) {
        super.log_info(
          'processListDataRows',
          `Failed to provision row ${index} in list ${lc.Title}: ${
            (error && error.message) || error
          }`
        )
      }
    }
  }

  /**
   * Builds the REST item payload from a data row, mapping each field by type
   * (Lookup/User → `<Field>Id`, Taxonomy → the hidden note field, URL, etc.).
   */
  private async _buildItemValues(
    web: IWeb,
    byName: Map<string, any>,
    byId: Map<string, any>,
    row: IDataRow
  ): Promise<Record<string, any>> {
    const values: Record<string, any> = {}
    for (const fieldName of Object.keys(row)) {
      const raw = row[fieldName]
      if (raw === null || raw === undefined) continue
      const def = byName.get(fieldName)
      const type: string = (def && def.TypeAsString) || 'Text'
      switch (type) {
        case 'Lookup':
          values[`${fieldName}Id`] = await this._resolveLookupId(web, def, raw)
          break
        case 'LookupMulti':
          values[`${fieldName}Id`] = await Promise.all(
            Lists._toArray(raw).map((value) => this._resolveLookupId(web, def, value))
          )
          break
        case 'User':
          values[`${fieldName}Id`] = await this._resolveUserId(web, raw)
          break
        case 'UserMulti':
          values[`${fieldName}Id`] = await Promise.all(
            Lists._toArray(raw).map((value) => this._resolveUserId(web, value))
          )
          break
        case 'TaxonomyFieldType':
        case 'TaxonomyFieldTypeMulti':
          this._applyTaxonomy(values, byId, def, Lists._toArray(raw))
          break
        case 'URL':
          values[fieldName] =
            typeof raw === 'string'
              ? { Url: raw }
              : { Url: raw.Url, Description: raw.Description || raw.Url }
          break
        case 'MultiChoice':
          values[fieldName] = Lists._toArray(raw)
          break
        case 'Boolean':
          values[fieldName] = raw === true || raw === 'true' || raw === 1 || raw === '1'
          break
        case 'DateTime':
          values[fieldName] = raw instanceof Date ? raw.toISOString() : raw
          break
        default:
          values[fieldName] = raw
      }
    }
    return values
  }

  /**
   * Resolves a lookup value to an item id: a number (or `{ lookupId }`) is used
   * directly; a string (or `{ lookupValue }`) is looked up in the target list by
   * its show field.
   */
  private async _resolveLookupId(web: IWeb, def: any, raw: any): Promise<number> {
    if (typeof raw === 'number') return raw
    if (raw && typeof raw === 'object' && raw.lookupId !== undefined) return raw.lookupId
    const value = typeof raw === 'string' ? raw : raw && raw.lookupValue
    if (value === undefined || value === null || !def || !def.LookupList) return
    const listId = Lists._normGuid(def.LookupList)
    const showField = def.LookupField || 'Title'
    const items = await web.lists
      .getById(listId)
      .items.filter(`${showField} eq '${Lists._escapeOData(String(value))}'`)
      .select('Id')
      .top(1)<Array<{ Id: number }>>()
    return items[0] ? items[0].Id : undefined
  }

  /**
   * Resolves a user value (login/email string, `{ login }`/`{ email }`, or a
   * numeric id) to a user id via `ensureUser`.
   */
  private async _resolveUserId(web: IWeb, raw: any): Promise<number> {
    if (typeof raw === 'number') return raw
    const login = typeof raw === 'string' ? raw : raw && (raw.login || raw.email)
    if (!login) return
    const result = await web.ensureUser(login)
    return result && result.data ? result.data.Id : undefined
  }

  /**
   * Writes taxonomy value(s) by setting the field's hidden note field to the
   * `-1;#Label|TermGuid` form (joined with `;#` for multi-value fields). Terms
   * without a label are skipped (the label is required to write the note field).
   */
  private _applyTaxonomy(
    values: Record<string, any>,
    byId: Map<string, any>,
    def: any,
    terms: any[]
  ): void {
    if (!def || !def.TextField) return
    const noteField = byId.get(Lists._normGuid(def.TextField))
    if (!noteField) return
    const parts = terms
      .filter((term) => term && term.termId && term.label)
      .map((term) => `-1;#${term.label}|${Lists._normGuid(term.termId)}`)
    if (parts.length === 0) return
    values[noteField.InternalName] = parts.join(';#')
  }

  private static _normGuid(value?: string): string {
    return (value || '').replace(/[{}]/g, '').toLowerCase()
  }

  private static _escapeOData(value: string): string {
    return value.replace(/'/g, '\'\'')
  }

  private static _toArray<T>(value: T | T[]): T[] {
    return Array.isArray(value) ? value : [value]
  }
}
