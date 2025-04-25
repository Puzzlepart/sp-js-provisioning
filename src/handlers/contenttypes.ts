/* eslint-disable unicorn/prevent-abbreviations */
import { IContentTypeAddResult, IWeb } from '@pnp/sp/presets/all'
import initSpfxJsom, { ExecuteJsomQuery, JsomContext } from 'spfx-jsom'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { IContentType, IFieldReference } from '../schema'
import { HandlerBase } from './handlerbase'

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
      this._initContext(web)
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
      const contentTypeId = contentType.ID ?? this.context.contentTypes[contentType.Name]?.ID
      if (!contentTypeId)
        throw new Error(
          `Content type with name '${contentType.Name}' does not exist in the web.`
        )
      super.log_info(
        'processContentType',
        `Processing content type [${contentType.Name}] (${contentTypeId})`
      )
      const spContentType = this.jsomContext.web
        .get_contentTypes()
        .getById(contentTypeId)
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
    } catch (error) {
      throw error
    }
  }

  /**
   * Add a content type
   *
   * @param web - The web
   * @param contentType - Content type
   */
  private async addContentType(
    web: IWeb,
    contentType: IContentType
  ): Promise<IContentTypeAddResult> {
    try {
      super.log_info(
        'addContentType',
        `Adding content type [${contentType.Name}] (${contentType.ID})`
      )
      return await web.contentTypes.add(
        contentType.ID,
        contentType.Name,
        contentType.Description,
        contentType.Group
      )
    } catch (error) {
      throw error
    }
  }

  private getExistingFieldLink(
    contentType: IContentType,
    fieldName: string
  ): IFieldReference {
    const ct = this.context.contentTypes[contentType.Name] ?? this.context.contentTypes[contentType.ID]
    const existingFieldLink = ct?.FieldRefs?.find((fr) => fr.Name === fieldName)
    return existingFieldLink
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
    try {
      const fieldRefs = contentType.FieldRefs
      for (const [index, fieldReference] of fieldRefs.entries()) {
        const existingFieldLink = this.getExistingFieldLink(
          contentType,
          fieldReference.Name
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
      super.log_info(
        'processContentTypeFieldRefs',
        `Successfully processed field refs for content type [${contentType.Name}] (${contentType.ID})`
      )
    } catch (error) {
      // eslint-disable-next-line no-console
      console.log(error)
      super.log_info(
        'processContentTypeFieldRefs',
        `Failed to process field refs for content type [${contentType.Name}] (${contentType.ID})`,
        { error: error.args && error.args.get_message() }
      )
    }
  }

  private async _initContext(web: IWeb): Promise<void> {
    this.context.contentTypes = (
      await web.contentTypes
        .select('Id', 'Name', 'FieldLinks')
        .expand('FieldLinks')()
    ).reduce((object, contentType) => {
      const ct = {
        ID: contentType.Id.StringValue,
        Name: contentType.Name,
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
