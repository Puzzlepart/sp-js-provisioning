/* eslint-disable unicorn/prevent-abbreviations */
import { IWeb } from '@pnp/sp/presets/all'
import initSpfxJsom, { ExecuteJsomQuery, JsomContext } from 'spfx-jsom'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { IContentType, IFieldReference } from '../schema'
import { HandlerBase } from './handlerbase'

interface ContentTypeCreationInformationWithId
  extends SP.ContentTypeCreationInformation {
  set_id(value: string): void
}

interface CurrentFieldLink {
  ID: string
  Name: string
  Required: boolean
  Hidden: boolean
}

/**
 * Describes the Content Types Object Handler
 */
export class ContentTypes extends HandlerBase {
  public jsomContext: JsomContext
  public context: ProvisioningContext

  /**
   * Creates a new instance of the ObjectSiteFields class
   */
  constructor(config: IProvisioningConfig) {
    super('ContentTypes', config)
  }

  /**
   * Provisioning Content Types
   *
   * @param web - The web
   * @param contentTypes - The content types
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: IWeb,
    contentTypes: IContentType[],
    context: ProvisioningContext
  ): Promise<void> {
    this.jsomContext = (
      await initSpfxJsom(context.web.ServerRelativeUrl)
    ).jsomContext
    this.context = context
    super.scope_started()
    try {
      await this._initContext(web)
      await contentTypes
        .sort((a, b) => {
          if (a.ID < b.ID) {
            return -1
          }
          if (a.ID > b.ID) {
            return 1
          }
          return 0
        })
        .reduce(
          (chain: any, contentType) =>
            chain.then(() => this.processContentType(contentType)),
          Promise.resolve()
        )
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Provision a content type
   *
   * @param contentType - Content type
   */
  private async processContentType(
    contentType: IContentType
  ): Promise<void> {
    try {
      const contentTypeId =
        contentType.ID ?? this.context.contentTypes[contentType.Name]?.ID
      if (!contentTypeId)
        throw new Error(
          `Content type with name '${contentType.Name}' does not exist in the web.`
        )

      const existingContentType = this.context.contentTypes[contentTypeId]
      const spContentType = existingContentType
        ? this.jsomContext.web.get_contentTypes().getById(contentTypeId)
        : await this.createContentType(contentType, contentTypeId)

      super.log_info(
        'processContentType',
        `Processing content type [${contentType.Name}] (${contentTypeId})`
      )

      spContentType.set_name(contentType.Name)
      if (contentType.Description) {
        spContentType.set_description(contentType.Description)
      }
      if (contentType.Group) {
        spContentType.set_group(contentType.Group)
      }
      spContentType.update(true)
      await ExecuteJsomQuery(this.jsomContext)
      if (contentType.FieldRefs) {
        await this.processContentTypeFieldRefs(contentType, spContentType)
      }

      this.context.contentTypes[contentTypeId] = {
        ID: contentTypeId,
        Name: contentType.Name,
        Description:
          contentType.Description ?? existingContentType?.Description ?? '',
        Group: contentType.Group ?? existingContentType?.Group ?? '',
        FieldRefs: contentType.FieldRefs ?? []
      }
      this.context.contentTypes[contentType.Name] =
        this.context.contentTypes[contentTypeId]
    } catch (error) {
      throw error
    }
  }

  private async createContentType(
    contentType: IContentType,
    contentTypeId: string
  ): Promise<SP.ContentType> {
    super.log_info(
      'createContentType',
      `Creating content type [${contentType.Name}] (${contentTypeId})`
    )

    const ctInfo = new SP.ContentTypeCreationInformation()
    ctInfo.set_name(contentType.Name)
    const ctInfoWithId = ctInfo as ContentTypeCreationInformationWithId
    ctInfoWithId.set_id(contentTypeId)
    if (contentType.Description) {
      ctInfo.set_description(contentType.Description)
    }
    if (contentType.Group) {
      ctInfo.set_group(contentType.Group)
    }
    // NB: do NOT call ctInfo.set_parentContentType() when ctInfo.set_id() is
    // used — SharePoint rejects the combination with
    //   "parameters.Id, parameters.ParentContentType cannot be used together".
    // The parent is inferred from the explicit ID's structure.

    const spContentType = this.jsomContext.web.get_contentTypes().add(ctInfo)
    await ExecuteJsomQuery(this.jsomContext)
    return spContentType
  }

  private getExistingFieldLink(
    currentFieldLinks: CurrentFieldLink[],
    fieldReference: IFieldReference
  ): CurrentFieldLink {
    const fieldReferenceId = this.normalizeGuid(fieldReference.ID)
    return currentFieldLinks.find((fieldLink) => {
      return (
        fieldLink.Name === fieldReference.Name ||
        this.normalizeGuid(fieldLink.ID) === fieldReferenceId
      )
    })
  }

  private normalizeGuid(guid: string): string {
    return (guid ?? '').replace(/[{}]/g, '').toLowerCase()
  }

  private async getCurrentFieldLinks(
    spContentType: SP.ContentType
  ): Promise<CurrentFieldLink[]> {
    const fieldLinks = spContentType.get_fieldLinks()
    await ExecuteJsomQuery(this.jsomContext, [{ clientObject: fieldLinks }])

    const currentFieldLinks: CurrentFieldLink[] = []
    const enumerator = fieldLinks.getEnumerator()
    while (enumerator.moveNext()) {
      const fieldLink = enumerator.get_current()
      currentFieldLinks.push({
        ID: fieldLink.get_id().toString(),
        Name: fieldLink.get_name(),
        Required: fieldLink.get_required(),
        Hidden: fieldLink.get_hidden()
      })
    }
    return currentFieldLinks
  }

  /**
   * Adding content type field refs
   *
   * @param contentType - Content type
   * @param spContentType - SP content type
   */
  private async processContentTypeFieldRefs(
    contentType: IContentType,
    spContentType: SP.ContentType
  ): Promise<void> {
    const fieldRefs = contentType.FieldRefs
    const currentFieldLinks = await this.getCurrentFieldLinks(spContentType)
    for (const [index, fieldReference] of fieldRefs.entries()) {
      const existingFieldLink = this.getExistingFieldLink(
        currentFieldLinks,
        fieldReference
      )
      let fieldLink: SP.FieldLink
      if (existingFieldLink) {
        fieldLink = spContentType
          .get_fieldLinks()
          .getById(new SP.Guid(existingFieldLink.ID))
      } else {
        super.log_info(
          'processContentTypeFieldRefs',
          `Adding field ref ${fieldReference.Name} to content type [${contentType.Name}] (${contentType.ID})`
        )
        const siteField = this.jsomContext.web
          .get_fields()
          .getByInternalNameOrTitle(fieldReference.Name)
        const fieldLinkCreationInformation = new SP.FieldLinkCreationInformation()
        fieldLinkCreationInformation.set_field(siteField)
        fieldLink = spContentType
          .get_fieldLinks()
          .add(fieldLinkCreationInformation)
      }
      if (contentType.FieldRefs[index].hasOwnProperty('Required')) {
        fieldLink.set_required(contentType.FieldRefs[index].Required)
      }
      if (contentType.FieldRefs[index].hasOwnProperty('Hidden')) {
        fieldLink.set_hidden(contentType.FieldRefs[index].Hidden)
      }
    }
    spContentType.update(true)
    await ExecuteJsomQuery(this.jsomContext)

    // Reorder field links to match the order specified in the schema
    const fieldRefNames = fieldRefs
      .map((fr) => fr.Name)
      .filter(Boolean) as string[]
    if (fieldRefNames.length > 1) {
      super.log_info(
        'processContentTypeFieldRefs',
        `Reordering field refs for content type [${contentType.Name}] (${contentType.ID})`
      )
      spContentType.get_fieldLinks().reorder(fieldRefNames)
      spContentType.update(true)
      await ExecuteJsomQuery(this.jsomContext)
    }

    super.log_info(
      'processContentTypeFieldRefs',
      `Successfully processed field refs for content type [${contentType.Name}] (${contentType.ID})`
    )
  }

  private async _initContext(web: IWeb): Promise<void> {
    this.context.contentTypes = (
      await web.contentTypes
        .select('Id', 'Name', 'Description', 'Group', 'FieldLinks')
        .expand('FieldLinks')()
    ).reduce((object, contentType) => {
      const ct = {
        ID: contentType.Id.StringValue,
        Name: contentType.Name,
        Description: contentType.Description ?? '',
        Group: contentType.Group ?? '',
        FieldRefs: contentType['FieldLinks'].map((fieldLink: any) => ({
          ID: fieldLink.Id,
          Name: fieldLink.Name,
          Required: fieldLink.Required,
          Hidden: fieldLink.Hidden
        }))
      }
      object[ct.Name] = ct
      object[ct.ID] = ct
      return object
    }, {})
  }
}
