/* eslint-disable unicorn/no-array-for-each */
import { IList, IWeb } from '@pnp/sp/presets/all'
import initSpfxJsom, { JsomContext } from 'spfx-jsom'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import {
  IContentTypeBinding,
  IDataRow,
  IFolder as IFolderConfig,
  IListInstance,
  IListInstanceFieldReference,
  IListView
} from '../schema'
import {
  declaredType,
  fieldAttributes,
  mergeFieldXml,
  sameFieldXml,
  setFieldAttributes,
  TARGET,
  targetChanged,
  typeChange
} from '../util/fieldxml'
import { resolvePrincipalId, resolveRoleDefinitionId } from '../util/security'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'

export interface ISPField {
  Id: string
  InternalName: string
  SchemaXml: string
  TypeAsString?: string
}

/**
 * A GUID attribute of a field XML (`ID` by default), normalised; empty when the XML cannot be read.
 *
 * @param xml - The field XML
 * @param attribute - The attribute
 */
function fieldId(xml: string, attribute = 'ID'): string {
  try {
    return (fieldAttributes(xml)[attribute] || '')
      .replace(/[{}]/g, '')
      .toLowerCase()
  } catch {
    return ''
  }
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
        await web.lists.select('Id', 'Title').top(5000)<
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
          chain.then(() => this.processListFolders(web, list)),
        Promise.resolve()
      )
      await lists.reduce(
        (chain: any, list) =>
          chain.then(() => this.processListDataRows(web, list)),
        Promise.resolve()
      )
      await lists.reduce(
        (chain: any, list) =>
          chain.then(() => this.processListSecurity(web, list)),
        Promise.resolve()
      )
      this.context.lists = (
        await web.lists.select('Id', 'Title').top(5000)<
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
      list = web.lists.getById(this.context.lists[lc.Title])
      // Only what the template defines: `lists.ensure` sent PnPjs' defaults for what it leaves out
      // (an empty Description, content types off) on every apply.
      const properties: Record<string, any> = {
        ...(lc.Description !== undefined
          ? { Description: lc.Description }
          : {}),
        ...(lc.ContentTypesEnabled !== undefined
          ? { ContentTypesEnabled: lc.ContentTypesEnabled }
          : {}),
        ...(lc.AdditionalSettings || {})
      }
      if (Object.keys(properties).length === 0) {
        super.log_info('processList', `List ${lc.Title} already exists.`)
      } else {
        try {
          await list.update(properties)
          super.log_info(
            'processList',
            `List ${lc.Title} already exists, its settings updated.`
          )
        } catch (error) {
          super.log_error(
            'processList',
            `Failed to update the settings of list ${lc.Title}: ${
              (error && error.message) || error
            }`
          )
        }
      }
    } else {
      super.log_info(
        'processList',
        `List ${lc.Title} doesn't exist. Creating...`
      )
      const listInfo = await web.lists.add(
        lc.Title,
        lc.Description,
        lc.Template,
        lc.ContentTypesEnabled,
        lc.AdditionalSettings
      )
      list = web.lists.getById(listInfo.Id)
      this.context.lists[listInfo.Title] = listInfo.Id
      this.context.createdLists[lc.Title] = true
    }
    if (lc.ContentTypeBindings) {
      const removeExisting =
        !!lc.RemoveExistingContentTypes && !!this.context.createdLists[lc.Title]
      if (lc.RemoveExistingContentTypes && !removeExisting)
        super.log_info(
          'processList',
          `List ${lc.Title} was not made in this setup: RemoveExistingContentTypes leaves its content types as they are.`
        )
      await this.processContentTypeBindings(
        lc,
        list,
        lc.ContentTypeBindings,
        removeExisting
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
      const contentTypes = await list.contentTypes()
      for (const {
        Id: { StringValue: ContentTypeId }
      } of contentTypes) {
        const shouldRemove =
          !contentTypeBindings.some((ctb) =>
            Lists._boundTo(ContentTypeId, ctb.ContentTypeID)
          ) && !Lists._boundTo(ContentTypeId, '0x0120')
        if (!shouldRemove) {
          super.log_info(
            'processContentTypeBindings',
            `Skipping removal of content type ${ContentTypeId} from list ${lc.Title}`
          )
          continue
        }
        super.log_info(
          'processContentTypeBindings',
          `Removing content type ${ContentTypeId} from list ${lc.Title}`
        )
        // A content type that is still in use by existing items can't be
        // removed (SharePoint returns 500 "content type still in use"). Don't
        // let that abort the whole provisioning — log it and carry on.
        try {
          await list.contentTypes.getById(ContentTypeId).delete()
        } catch (error) {
          // SharePoint refuses a content type still in use with a 500 (its message is localised):
          // a warning. Anything else is a failure.
          const inUse = error && error.status === 500
          const message = `Failed to remove content type ${ContentTypeId} from list ${
            lc.Title
          }${inUse ? ' (likely still in use)' : ''}: ${
            (error && error.message) || error
          }`
          if (inUse) super.log_warn('processContentTypeBindings', message)
          else super.log_error('processContentTypeBindings', message)
        }
      }
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
      // addAvailableContentType is NOT idempotent — it returns a 500 when the
      // content type is already bound to the list (e.g. on re-import). A list
      // content type derived from a site content type has an id that starts with
      // the site content type's id, so skip the add when one already matches.
      const existing = await list.contentTypes()
      const alreadyBound = (existing || []).some((ct: any) =>
        Lists._boundTo(`${ct && ct.Id && ct.Id.StringValue}`, contentTypeID)
      )
      if (alreadyBound) {
        super.log_info(
          'processContentTypeBinding',
          `Content Type ${contentTypeID} already bound to list ${lc.Title} — skipping.`
        )
        return
      }
      super.log_info(
        'processContentTypeBinding',
        `Adding content Type ${contentTypeID} to list ${lc.Title}.`
      )
      await list.contentTypes.addAvailableContentType(contentTypeID)
      super.log_info(
        'processContentTypeBinding',
        `Content Type ${contentTypeID} added successfully to list ${lc.Title}.`
      )
    } catch (error) {
      super.log_error(
        'processContentTypeBinding',
        `Failed to add Content Type ${contentTypeID} to list ${lc.Title}: ${
          (error && error.message) || error
        }`
      )
    }
  }

  /**
   * Processes the fields (`Fields`, schema XML) of a list: the list's fields are read once, then
   * each template field is created when it is missing and updated in place when it exists. A list
   * field is never deleted (docs/specs/lists-re-apply.md): deleting it lost the values in the
   * column, its place in views and content types, and an admin's settings, on every apply.
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async processListFields(web: IWeb, lc: IListInstance): Promise<any> {
    if (!lc.Fields || lc.Fields.length === 0) return
    const list = web.lists.getByTitle(lc.Title)
    let existing: ISPField[]
    try {
      existing = await web.lists
        .getByTitle(lc.Title)
        .fields.select('Id', 'InternalName', 'TypeAsString', 'SchemaXml')<
        ISPField[]
      >()
    } catch (error) {
      super.log_error(
        'processListFields',
        `Failed to read the fields of list ${
          lc.Title
        }, its fields are left as they are: ${
          (error && error.message) || error
        }`
      )
      return
    }
    for (const fieldXml of Lists._primariesFirst(lc.Fields))
      await this.processField(list, lc, fieldXml, existing)
  }

  /**
   * Creates a list field, or updates it in place: the template merged onto the stored schema XML
   * (`mergeFieldXml`), sent only when it differs. A widening type change (Text to Note, single to
   * multi) is applied; another type change, or a new target (a lookup's list, a taxonomy field's
   * term set) on a list that was not made in this setup, is a warning, because SharePoint would
   * accept it and lose or misdirect the stored values.
   *
   * @param list - The list
   * @param lc - The list configuration
   * @param fieldXml - The template's field XML
   * @param existing - The list's fields, read before the first template field
   */
  private async processField(
    list: IList,
    lc: IListInstance,
    fieldXml: string,
    existing: ISPField[]
  ): Promise<void> {
    const template = this.tokenHelper.replaceTokens(fieldXml)
    let attributes: Record<string, string>
    try {
      attributes = fieldAttributes(template)
    } catch (error) {
      super.log_error(
        'processField',
        `Failed to read a field of list ${lc.Title}: ${
          (error && error.message) || error
        }`
      )
      return
    }
    const { Name: name, DisplayName: displayName, ID } = attributes
    const id = Lists._normGuid(ID)
    const label = `${name} (${displayName}) in list ${lc.Title}`
    super.log_info('processField', `Processing field ${label}.`)
    const stored = existing.find((field) => Lists._normGuid(field.Id) === id)
    try {
      if (!stored) {
        const sameName = existing.find((field) => field.InternalName === name)
        if (sameName) {
          super.log_warn(
            'processField',
            `Field ${label} is not created: the list has a field ${name} with another id (${sameName.Id}).`
          )
          return
        }
        // Created with the internal name as display name, so SharePoint derives the internal name
        // from it, then titled.
        const created = await list.fields.createFieldAsXml(
          setFieldAttributes(template, { DisplayName: name })
        )
        await list.fields
          .getById(created.Id ?? id)
          .update({ Title: displayName }, 'SP.Field')
        existing.push({
          Id: created.Id ?? id,
          InternalName: name,
          SchemaXml: template
        })
        super.log_info('processField', `Field ${label} added.`)
        return
      }
      const change = typeChange(stored.TypeAsString, declaredType(attributes))
      if (change === 'other') {
        super.log_warn(
          'processField',
          `Field ${label} is ${
            stored.TypeAsString
          }, the template says ${declaredType(
            attributes
          )}: SharePoint would lose values, so the field is left as it is.`
        )
        return
      }
      // A new target (a lookup's list, a taxonomy field's term set) follows the template on a
      // list made in this setup, which holds no values yet (a content type can bring the field
      // with another target); elsewhere the stored values belong to the target the field has.
      const keep: string[] = []
      if (targetChanged(stored.SchemaXml, template)) {
        if (this.context.createdLists[lc.Title]) {
          super.log_info(
            'processField',
            `Field ${label} points at another list or term set than the template says; the list was made in this setup, so it follows the template.`
          )
        } else {
          super.log_warn(
            'processField',
            `Field ${label} points at another list or term set than the template says: its stored values belong to the one it has, so it keeps it.`
          )
          keep.push(...TARGET)
        }
      }
      // The type as SharePoint writes it: `Type="Lookup" Mult="TRUE"` is a LookupMulti column, and
      // merged as written it would set Type="Lookup" on one.
      const merged = mergeFieldXml(
        stored.SchemaXml,
        setFieldAttributes(template, { Type: declaredType(attributes) }),
        keep
      )
      if (sameFieldXml(merged, stored.SchemaXml)) {
        super.log_info('processField', `Field ${label} is unchanged.`)
        return
      }
      await list.fields
        .getById(Lists._normGuid(stored.Id))
        .update({ SchemaXml: merged }, 'SP.Field')
      super.log_info(
        'processField',
        `Field ${label} updated in place${
          change === 'widening'
            ? ` (${stored.TypeAsString} to ${declaredType(attributes)})`
            : ''
        }.`
      )
    } catch (error) {
      super.log_error(
        'processField',
        `Failed to ${stored ? 'update' : 'add'} field '${displayName}' ${
          stored ? 'in' : 'to'
        } list ${lc.Title}: ${(error && error.message) || error}`
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
      for (const fieldReference of lc.FieldRefs) {
        try {
          await this.processFieldRef(
            list,
            lc,
            fieldReference,
            listFields,
            webFields
          )
        } catch (error) {
          super.log_error(
            'processFieldRef',
            `Failed to process field ref '${fieldReference.ID}' for list ${
              lc.Title
            }: ${(error && error.message) || error}`
          )
        }
      }
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
    const id = Lists._normGuid(fieldReference.ID)
    const listFld = listFields.find((f) => Lists._normGuid(f.Id) === id)
    const webFld = webFields.find((f) => Lists._normGuid(f.Id) === id)
    super.log_info(
      'processFieldRef',
      `Processing field ref '${fieldReference.ID}' for list ${lc.Title}.`
    )
    if (Boolean(listFld)) {
      const field = list.fields.getById(id)
      await field.update(
        {
          Hidden: fieldReference.Hidden,
          Required: fieldReference.Required,
          Title: fieldReference.DisplayName
        },
        'SP.Field'
      )
      if (fieldReference.AdditionalProperties) {
        // Read after the update above: its SchemaXml (and Version) is not the one read before.
        // Only the additional properties are set; the title stays what the update made it.
        const { SchemaXml } = await list.fields
          .getById(id)
          .select('SchemaXml')<{ SchemaXml: string }>()
        const schemaXml = setFieldAttributes(
          SchemaXml,
          fieldReference.AdditionalProperties
        )
        super.log_info(
          'processFieldRef',
          `Additional properties set for field ref '${fieldReference.ID}' for list ${lc.Title}. Attempting to generate schema XML.`,
          { schemaXml }
        )
        await list.fields
          .getById(id)
          .update({ SchemaXml: schemaXml }, 'SP.Field')
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
      const schemaXml = setFieldAttributes(webFld.SchemaXml, {
        DisplayName: fieldReference.Name,
        SourceID: `{${this.context.lists[lc.Title]}}`,
        ...(fieldReference.AdditionalProperties ?? {})
      })
      const addedField = await list.fields.createFieldAsXml(schemaXml)
      await list.fields.getById(addedField.Id ?? id).update(
        {
          Title: fieldReference.DisplayName,
          Required: fieldReference.Required,
          Hidden: fieldReference.Hidden
        },
        'SP.Field'
      )
      super.log_info(
        'processFieldRef',
        `Field '${fieldReference.ID}' added from web.`
      )
    }
  }

  /**
   * Processes views for a list: the template's views first (updated in place when they exist), then,
   * with `RemoveExistingViews` on a list made in this setup, the other views are removed.
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async processListViews(web: IWeb, lc: IListInstance): Promise<any> {
    if (lc.Views) {
      for (const view of lc.Views) await this.processView(web, lc, view)
      if (lc.RemoveExistingViews) {
        if (this.context.createdLists[lc.Title])
          await this.removeOtherViews(web, lc)
        else
          super.log_info(
            'processListViews',
            `List ${lc.Title} was not made in this setup: RemoveExistingViews leaves its other views as they are.`
          )
      }
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
   * Removes the views of a list made in this setup that the template does not name, after the
   * template's own exist. Never a view the template names, a personal view, the default view, or a
   * hidden view outside the list's root folder (the view behind a page's list web part, as in PnP).
   * One at a time; a refusal is a warning.
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async removeOtherViews(web: IWeb, lc: IListInstance): Promise<void> {
    let read: [
      Array<{
        Id: string
        Title: string
        Hidden: boolean
        PersonalView: boolean
        DefaultView: boolean
        ServerRelativeUrl: string
      }>,
      string
    ]
    try {
      read = await Promise.all([
        web.lists
          .getByTitle(lc.Title)
          .views.select(
            'Id',
            'Title',
            'Hidden',
            'PersonalView',
            'DefaultView',
            'ServerRelativeUrl'
          )<typeof read[0]>(),
        web.lists
          .getByTitle(lc.Title)
          .rootFolder.select('ServerRelativeUrl')<{
            ServerRelativeUrl: string
          }>()
          .then((folder) => `${folder.ServerRelativeUrl}/`.toLowerCase())
      ])
    } catch (error) {
      super.log_error(
        'removeOtherViews',
        `Failed to read the views of list ${lc.Title}: ${
          (error && error.message) || error
        }`
      )
      return
    }
    const [views, root] = read
    const named = new Set(lc.Views.map((view) => view.Title))
    for (const view of views) {
      const pageView =
        view.Hidden &&
        `${view.ServerRelativeUrl}`.toLowerCase().indexOf(root) !== 0
      if (
        named.has(view.Title) ||
        view.PersonalView ||
        view.DefaultView ||
        pageView
      )
        continue
      try {
        await web.lists.getByTitle(lc.Title).views.getById(view.Id).delete()
        super.log_info(
          'removeOtherViews',
          `View ${view.Title} removed from list ${lc.Title}.`
        )
      } catch (error) {
        super.log_warn(
          'removeOtherViews',
          `Failed to remove view ${view.Title} from list ${lc.Title}: ${
            (error && error.message) || error
          }`
        )
      }
    }
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
        // No settings, no update: an empty MERGE body is refused (400).
        if (Lists._hasSettings(lvc))
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
          if (Lists._hasSettings(lvc))
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
   * Gives a view the template's fields, in order, when it does not have them already: an unchanged
   * view is not touched. A field that cannot be added is logged and the others are still added.
   *
   * @param view - The pnp view
   * @param lvc - The view configuration
   */
  private async processViewFields(view: any, lvc: IListView): Promise<void> {
    if (!lvc.ViewFields) return
    try {
      super.log_info(
        'processViewFields',
        `Processing view fields for view ${lvc.Title}.`
      )
      const current: { Items?: string[] } = await view.fields()
      if (
        current.Items &&
        current.Items.length === lvc.ViewFields.length &&
        current.Items.every((name, index) => name === lvc.ViewFields[index])
      ) {
        super.log_info(
          'processViewFields',
          `View ${lvc.Title} already has the template's fields.`
        )
        return
      }
      await view.fields.removeAll()
    } catch (error) {
      super.log_error(
        'processViewFields',
        `Failed to process view fields for view ${lvc.Title}: ${
          (error && error.message) || error
        }`
      )
      return
    }
    for (const viewField of lvc.ViewFields) {
      try {
        await view.fields.add(viewField)
      } catch (error) {
        super.log_error(
          'processViewFields',
          `Failed to add field ${viewField} to view ${lvc.Title}: ${
            (error && error.message) || error
          }`
        )
      }
    }
    super.log_info(
      'processViewFields',
      `View fields processed for view ${lvc.Title}.`
    )
  }

  /**
   * Provisions a folder hierarchy inside a list / document library. Runs after
   * the list, its fields and views exist. Creation is idempotent: an existing
   * folder is left in place and recursion continues into its children, so
   * re-running the template is safe. A folder that fails to create is logged
   * and skipped (along with its subtree) rather than aborting the whole list.
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async processListFolders(
    web: IWeb,
    lc: IListInstance
  ): Promise<void> {
    const folders = lc.Folders
    if (!folders || folders.length === 0) return
    super.log_info(
      'processListFolders',
      `Provisioning ${folders.length} top-level folder(s) for list ${lc.Title}.`
    )
    const list = web.lists.getByTitle(lc.Title)
    let rootServerRelativeUrl: string
    try {
      const rootInfo = await list.rootFolder.select('ServerRelativeUrl')()
      rootServerRelativeUrl = rootInfo.ServerRelativeUrl
    } catch (error) {
      super.log_error(
        'processListFolders',
        `Could not resolve root folder for list ${lc.Title}: ${
          (error && error.message) || error
        }`
      )
      return
    }
    await this._processFolderLevel(
      list.rootFolder,
      rootServerRelativeUrl,
      folders,
      lc.Title
    )
  }

  /**
   * Creates each folder in `folders` beneath `parentFolder` (idempotent) then
   * recurses into its children. `parentServerRelativeUrl` is used only for
   * logging the full path.
   *
   * @param parentFolder - The PnP parent folder
   * @param parentServerRelativeUrl - Server-relative URL of the parent folder
   * @param folders - Folder configs to create at this level
   * @param listTitle - The list title (for logging)
   */
  private async _processFolderLevel(
    parentFolder: any,
    parentServerRelativeUrl: string,
    folders: IFolderConfig[],
    listTitle: string
  ): Promise<void> {
    // List the existing child folders once so creation is idempotent. Note
    // `addSubFolderUsingPath` is NOT idempotent — it fails with a 500 ("a folder
    // with that name already exists") when the folder is already there, e.g. when
    // a package is re-imported. We therefore only create the ones that are missing.
    let existingNames: string[] = []
    try {
      const existingFolders = await parentFolder.folders.select('Name')()
      existingNames = (existingFolders || []).map((f: any) =>
        `${f.Name}`.toLowerCase()
      )
    } catch {
      // Best-effort — if listing fails, fall back to create-and-catch below.
    }
    for (const folder of folders) {
      const name = folder.Name
      if (!name) continue
      const childServerRelativeUrl = `${parentServerRelativeUrl}/${name}`
      try {
        let childFolder: any
        if (existingNames.includes(`${name}`.toLowerCase())) {
          childFolder = parentFolder.folders.getByUrl(name)
          super.log_info(
            'processListFolders',
            `Folder '${childServerRelativeUrl}' already exists in list ${listTitle} — skipping creation.`
          )
        } else {
          childFolder = await parentFolder.addSubFolderUsingPath(name)
          super.log_info(
            'processListFolders',
            `Created folder '${childServerRelativeUrl}' in list ${listTitle}.`
          )
        }
        if (folder.Folders && folder.Folders.length > 0) {
          await this._processFolderLevel(
            childFolder,
            childServerRelativeUrl,
            folder.Folders,
            listTitle
          )
        }
      } catch (error) {
        super.log_error(
          'processListFolders',
          `Failed to provision folder '${childServerRelativeUrl}' in list ${listTitle}: ${
            (error && error.message) || error
          }`
        )
      }
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
  private async processListDataRows(
    web: IWeb,
    lc: IListInstance
  ): Promise<void> {
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
    const byId = new Map<string, any>(
      fields.map((f) => [Lists._normGuid(f.Id), f])
    )

    for (const [index, row] of Rows.entries()) {
      // The key is matched resolved, as the item holds it, and escaped after: a resolved
      // {sitename} can hold an apostrophe.
      const rawKey = KeyColumn ? row[KeyColumn] : undefined
      const keyValue =
        typeof rawKey === 'string'
          ? this.tokenHelper.replaceTokens(rawKey)
          : rawKey
      const where = `row ${index}${
        KeyColumn ? ` (${KeyColumn}=${keyValue})` : ''
      } in list ${lc.Title}`
      try {
        let existingId: number
        if (KeyColumn) {
          if (keyValue === undefined || keyValue === null || keyValue === '') {
            super.log_warn(
              'processListDataRows',
              `Row ${index} in list ${lc.Title} has no value for its KeyColumn ${KeyColumn}: it is added, not matched.`
            )
          } else {
            const matches = await list.items
              .filter(
                `${KeyColumn} eq '${Lists._escapeOData(String(keyValue))}'`
              )
              .select('Id')
              .top(2)<Array<{ Id: number }>>()
            if (matches.length > 1)
              super.log_warn(
                'processListDataRows',
                `More than one item in list ${lc.Title} has ${KeyColumn}=${keyValue}: the first is used.`
              )
            existingId = matches[0] ? matches[0].Id : undefined
          }
        }
        if (existingId !== undefined && UpdateBehavior === 'Skip') {
          super.log_info(
            'processListDataRows',
            `Row ${index} already exists (${KeyColumn}=${keyValue}) — skipped.`
          )
          continue
        }
        // Built only for a row that is written: a skipped one queries no lookup list and ensures
        // no user.
        const values = await this._buildItemValues(web, byName, byId, row, {
          where,
          existing: existingId !== undefined
        })
        await (existingId === undefined
          ? list.items.add(values)
          : list.items.getById(existingId).update(values))
      } catch (error) {
        super.log_error(
          'processListDataRows',
          `Failed to provision ${where}: ${(error && error.message) || error}`
        )
      }
    }
  }

  /**
   * Builds the REST item payload from a data row, mapping each field by type
   * (Lookup/User → `<Field>Id`, Taxonomy → the hidden note field, URL, etc.).
   * A Lookup or User value that matches nothing is logged as an error naming it
   * and left out: a new item gets the values that resolved, an existing one
   * keeps what it holds in that column.
   *
   * @param web - The web
   * @param byName - The list's fields by internal name
   * @param byId - The list's fields by id
   * @param row - The data row
   * @param target - The row, for the log, and whether it updates an existing item
   */
  private async _buildItemValues(
    web: IWeb,
    byName: Map<string, any>,
    byId: Map<string, any>,
    row: IDataRow,
    target: { where: string; existing: boolean }
  ): Promise<Record<string, any>> {
    const values: Record<string, any> = {}
    /**
     * Resolves the values of a Lookup or User column, logs those that resolve to nothing and sets
     * what is left (`multi`: an array; an existing item keeps its stored value when one is missing).
     */
    const setResolved = async (
      fieldName: string,
      raws: any[],
      multi: boolean,
      resolve: (raw: any) => Promise<{ id?: number; reason?: string }>
    ) => {
      // A request that may succeed later (throttled, unavailable, no answer) fails the row, so the
      // next run writes it whole; a Skip row would otherwise never get the value.
      const results = await Promise.all(
        raws.map((raw) =>
          resolve(raw).catch((error) => {
            if (Lists._transient(error)) throw error
            return {
              id: undefined,
              reason: `${(error && error.message) || error}`
            }
          })
        )
      )
      const missing = results
        .map((result, index) => ({ ...result, raw: raws[index] }))
        .filter((result) => result.id === undefined)
      for (const result of missing)
        super.log_error(
          'processListDataRows',
          `Value ${JSON.stringify(result.raw)} for column ${fieldName} in ${
            target.where
          } is not written: ${result.reason}${
            target.existing ? '; the stored value is kept' : ''
          }.`
        )
      const ids = results
        .map((result) => result.id)
        .filter((id) => id !== undefined)
      if (target.existing && missing.length > 0) return
      if (multi) values[`${fieldName}Id`] = ids
      else if (ids.length > 0) values[`${fieldName}Id`] = ids[0]
    }
    for (const fieldName of Object.keys(row)) {
      const raw = row[fieldName]
      if (raw === null || raw === undefined) continue
      const def = byName.get(fieldName)
      const type: string = (def && def.TypeAsString) || 'Text'
      switch (type) {
        case 'Lookup':
        case 'LookupMulti':
          await setResolved(
            fieldName,
            type === 'Lookup' ? [raw] : Lists._toArray(raw),
            type === 'LookupMulti',
            (value) => this._resolveLookup(web, def, value)
          )
          break
        case 'User':
        case 'UserMulti':
          await setResolved(
            fieldName,
            type === 'User' ? [raw] : Lists._toArray(raw),
            type === 'UserMulti',
            async (value) => {
              const id = await this._resolveUserId(web, value)
              return id === undefined
                ? { reason: 'it names no user (login, email or id)' }
                : { id }
            }
          )
          break
        case 'TaxonomyFieldType':
        case 'TaxonomyFieldTypeMulti':
          {
            const problem = this._applyTaxonomy(
              values,
              byId,
              def,
              Lists._toArray(raw)
            )
            if (problem)
              super.log_error(
                'processListDataRows',
                `Value ${JSON.stringify(raw)} for column ${fieldName} in ${
                  target.where
                } is not written in full: ${problem}.`
              )
          }
          break
        case 'URL': {
          const url = this.tokenHelper.replaceTokens(
            typeof raw === 'string' ? raw : raw.Url
          )
          values[fieldName] =
            typeof raw === 'string'
              ? { Url: url }
              : {
                  Url: url,
                  Description: raw.Description
                    ? this.tokenHelper.replaceTokens(raw.Description)
                    : url
                }
          break
        }
        case 'MultiChoice':
          values[fieldName] = Lists._toArray(raw)
          break
        case 'Boolean':
          values[fieldName] =
            raw === true || raw === 'true' || raw === 1 || raw === '1'
          break
        case 'DateTime':
          values[fieldName] = raw instanceof Date ? raw.toISOString() : raw
          break
        default:
          values[fieldName] =
            typeof raw === 'string' ? this.tokenHelper.replaceTokens(raw) : raw
      }
    }
    return values
  }

  /**
   * Applies list security (`Security`): optionally breaks role inheritance and
   * adds role assignments. A principal may be a PnP-style token for one of the
   * web's associated groups (`{associatedownergroupid}`,
   * `{associatedmembergroupid}`, `{associatedvisitorgroupid}`), a principal id,
   * a site group name or a user login name; other tokens in the principal
   * (e.g. `{sitename}`, `{parameter:Name}`) are replaced first. Role definitions
   * are resolved by their (localized) name, falling back to the well-known role
   * type for the English names. Runs after data rows. Errors are logged, not
   * thrown.
   *
   * @param web - The web
   * @param lc - The list configuration
   */
  private async processListSecurity(
    web: IWeb,
    lc: IListInstance
  ): Promise<void> {
    const security = lc.Security
    if (!security) return
    super.log_info(
      'processListSecurity',
      `Processing security for list ${lc.Title}.`
    )
    try {
      const list = web.lists.getByTitle(lc.Title)
      // Inheritance is broken only on a list that still inherits, as on the web: re-applying a
      // template with ClearSubscopes must not clear unique permissions set since.
      let inherits: boolean | undefined
      try {
        inherits = !(
          await web.lists
            .getByTitle(lc.Title)
            .select('HasUniqueRoleAssignments')<{
            HasUniqueRoleAssignments: boolean
          }>()
        ).HasUniqueRoleAssignments
      } catch (error) {
        super.log_error(
          'processListSecurity',
          `Failed to read the permissions of list ${
            lc.Title
          }, role inheritance is left as is: ${
            (error && error.message) || error
          }`
        )
      }
      if (security.BreakRoleInheritance && inherits === false) {
        super.log_info(
          'processListSecurity',
          `List ${lc.Title} already has unique permissions, role inheritance is left as is.`
        )
      } else if (security.BreakRoleInheritance && inherits) {
        await list.breakRoleInheritance(
          !!security.CopyRoleAssignments,
          !!security.ClearSubscopes
        )
        super.log_info(
          'processListSecurity',
          `Role inheritance broken for list ${
            lc.Title
          } (copyRoleAssignments=${!!security.CopyRoleAssignments}, clearSubscopes=${!!security.ClearSubscopes}).`
        )
      }
      if (
        !security.BreakRoleInheritance &&
        inherits &&
        (security.RoleAssignments || []).length > 0
      ) {
        super.log_warn(
          'processListSecurity',
          `List ${lc.Title} has role assignments but BreakRoleInheritance is not set. SharePoint rejects role assignments on a list that still inherits permissions.`
        )
      }
      for (const assignment of security.RoleAssignments || []) {
        const principal =
          typeof assignment.Principal === 'string'
            ? this.tokenHelper.replaceTokens(assignment.Principal)
            : assignment.Principal
        try {
          const principalId = await resolvePrincipalId(
            web,
            principal,
            (message) => this.log_info('resolvePrincipalId', message)
          )
          const roleDefinitionId = await resolveRoleDefinitionId(
            web,
            assignment.RoleDefinition
          )
          await list.roleAssignments.add(principalId, roleDefinitionId)
          super.log_info(
            'processListSecurity',
            `Role assignment added on list ${lc.Title}: ${principal} (${principalId}) -> ${assignment.RoleDefinition} (${roleDefinitionId}).`
          )
        } catch (error) {
          super.log_error(
            'processListSecurity',
            `Failed to add role assignment ${principal} -> ${
              assignment.RoleDefinition
            } on list ${lc.Title}: ${(error && error.message) || error}`
          )
        }
      }
    } catch (error) {
      super.log_error(
        'processListSecurity',
        `Failed to process security for list ${lc.Title}: ${
          (error && error.message) || error
        }`
      )
    }
  }

  /**
   * Resolves a lookup value to an item id: a number (or `{ lookupId }`) is used
   * directly; a string (or `{ lookupValue }`) is looked up in the target list by
   * its show field. Without an id, `reason` says why.
   */
  private async _resolveLookup(
    web: IWeb,
    def: any,
    raw: any
  ): Promise<{ id?: number; reason?: string }> {
    if (typeof raw === 'number') return { id: raw }
    if (raw && typeof raw === 'object' && raw.lookupId !== undefined)
      return { id: raw.lookupId }
    const value = typeof raw === 'string' ? raw : raw && raw.lookupValue
    if (value === undefined || value === null)
      return { reason: 'it is not an id, a lookupId or a lookupValue' }
    if (!def || !def.LookupList)
      return { reason: 'the column has no lookup list' }
    const listId = Lists._normGuid(def.LookupList)
    const showField = def.LookupField || 'Title'
    const items = await web.lists
      .getById(listId)
      .items.filter(`${showField} eq '${Lists._escapeOData(String(value))}'`)
      .select('Id')
      .top(1)<Array<{ Id: number }>>()
    return items[0]
      ? { id: items[0].Id }
      : { reason: `no item in its lookup list has ${showField} = '${value}'` }
  }

  /**
   * Resolves a user value (login/email string, `{ login }`/`{ email }`, or a
   * numeric id) to a user id via `ensureUser`.
   */
  private async _resolveUserId(web: IWeb, raw: any): Promise<number> {
    if (typeof raw === 'number') return raw
    const login =
      typeof raw === 'string' ? raw : raw && (raw.login || raw.email)
    if (!login) return
    const user = await web.ensureUser(login)
    return user ? user.Id : undefined
  }

  /**
   * Writes taxonomy value(s) by setting the field's hidden note field to the
   * `-1;#Label|TermGuid` form (joined with `;#` for multi-value fields). Terms
   * without a term id or a label are skipped (the label is required to write the
   * note field); what could not be written is returned, for the log.
   */
  private _applyTaxonomy(
    values: Record<string, any>,
    byId: Map<string, any>,
    def: any,
    terms: any[]
  ): string | undefined {
    const noteField =
      def && def.TextField && byId.get(Lists._normGuid(def.TextField))
    if (!noteField) return 'the column has no hidden note field'
    const valid = terms.filter((term) => term && term.termId && term.label)
    if (valid.length > 0)
      values[noteField.InternalName] = valid
        .map((term) => `-1;#${term.label}|${Lists._normGuid(term.termId)}`)
        .join(';#')
    if (valid.length < terms.length) return 'a term needs both termId and label'
  }

  /**
   * Whether a list content type belongs to a binding: its id is the binding's or a child of it.
   *
   * @param listContentTypeId - The list content type's id
   * @param contentTypeId - The bound (site) content type's id
   */
  private static _boundTo(
    listContentTypeId: string,
    contentTypeId: string
  ): boolean {
    return (
      listContentTypeId.toLowerCase().indexOf(contentTypeId.toLowerCase()) === 0
    )
  }

  /**
   * Field XMLs with each dependent lookup (`FieldRef` to another field of the list) after the field
   * it depends on, the rest in template order: a dependent cannot be made before its primary.
   *
   * @param fields - The template's field XMLs
   */
  private static _primariesFirst(fields: string[]): string[] {
    const ids = new Set(fields.map((xml) => fieldId(xml)))
    const ordered: string[] = []
    const visit = (xml: string, seen: string[]) => {
      if (ordered.includes(xml) || seen.includes(xml)) return
      const primary = fieldId(xml, 'FieldRef')
      if (primary && ids.has(primary))
        for (const other of fields)
          if (fieldId(other) === primary) visit(other, [...seen, xml])
      ordered.push(xml)
    }
    for (const xml of fields) visit(xml, [])
    return ordered
  }

  /**
   * Whether a view configuration has settings to send.
   *
   * @param lvc - The view configuration
   */
  private static _hasSettings(lvc: IListView): boolean {
    return (
      !!lvc.AdditionalSettings && Object.keys(lvc.AdditionalSettings).length > 0
    )
  }

  /**
   * Whether a failed request may succeed later: throttled, unavailable, timed out, or no answer.
   *
   * @param error - What the request rejected with
   */
  private static _transient(error: any): boolean {
    const status = error && error.status
    return !status || status === 429 || status === 503 || status === 504
  }

  private static _normGuid(value?: string): string {
    return (value || '').replace(/[{}]/g, '').toLowerCase()
  }

  private static _escapeOData(value: string): string {
    return value.replace(/'/g, "''")
  }

  private static _toArray<T>(value: T | T[]): T[] {
    return Array.isArray(value) ? value : [value]
  }
}
