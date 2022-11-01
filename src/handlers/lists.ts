/* eslint-disable unicorn/no-array-for-each */
import initSpfxJsom, { JsomContext } from 'spfx-jsom'
import * as xmljs from 'xml-js'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import {
  IContentTypeBinding,
  IListInstance,
  IListInstanceFieldReference,
  IListView
} from '../schema'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'
import { addFieldAttributes } from '../util'
import { IWeb } from '@pnp/sp/webs'
import { IList } from '@pnp/sp/lists'

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
    super(Lists.name, config)
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
        await web.lists
          .select('Id', 'Title')()
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
      this.context.lists = (
        await web.lists
          .select('Id', 'Title')()
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
  private async processListFields(web: IWeb, list: IListInstance): Promise<any> {
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
  private async processListFieldRefs(web: IWeb, lc: IListInstance): Promise<any> {
    if (lc.FieldRefs) {
      super.log_info(
        'processListFieldRefs',
        `Retrieving fields for list ${lc.Title} and web.`
      )
      const list = web.lists.getByTitle(lc.Title)
      const [listFields, webFields] = await Promise.all([
        list.fields
          .select('Id', 'InternalName', 'SchemaXml')(),
        web.fields
          .select('Id', 'InternalName', 'SchemaXml')()
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
   * @param fieldReference - The list field ref
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
    if (listFld) {
      await list.fields.getById(fieldReference.ID).update({
        Hidden: fieldReference.Hidden,
        Required: fieldReference.Required,
        Title: fieldReference.DisplayName
      })
      super.log_info(
        'processFieldRef',
        `Field '${fieldReference.ID}' updated for list ${lc.Title}.`
      )
    } else if (webFld) {
      super.log_info(
        'processFieldRef',
        `Adding field '${fieldReference.ID}' to list ${lc.Title}.`
      )
      const schemaXml = addFieldAttributes(webFld.SchemaXml, {
        DisplayName: fieldReference.Name,
        SourceID: `{${this.context.lists[lc.Title]}}`
      })
      const fieldAddResult = await list.fields.createFieldAsXml(schemaXml)
      fieldAddResult.field.update({
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
      await lc.Views.reduce(
        (chain: any, view) => chain.then(() => this.processView(web, lc, view)),
        Promise.resolve()
      )
    }
    this.context.listViews = (
      await web.lists
        .getByTitle(lc.Title)
        .views.select('Id', 'Title')()
    ).reduce((object, view) => {
      object[`${lc.Title}|${view.Title}`] = view.Id
      return object
    }, this.context.listViews)
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
        const result = await web.lists
          .getByTitle(lc.Title)
          .views.add(lvc.Title, lvc.PersonalView, lvc.AdditionalSettings)
        super.log_info(
          'processView',
          `View ${lvc.Title} added successfully to list ${lc.Title}.`
        )
        await this.processViewFields(result.view, lvc)
      }
    } catch {
      super.log_info(
        'processViewFields',
        `Failed to process view for view ${lvc.Title}.`
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
}
