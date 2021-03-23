import initSpfxJsom, { ExecuteJsomQuery, JsomContext } from 'spfx-jsom'
import { HandlerBase } from './handlerbase'
import { Web, ContentTypeAddResult } from '@pnp/sp'
import { ProvisioningContext } from '../provisioningcontext'
import { IProvisioningConfig } from '../provisioningconfig'
import { IContentType } from '../schema'

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
    super(ContentTypes.name, config)
  }

  /**
   * Provisioning Content Types
   *
   * @param web - The web
   * @param contentTypes - The content types
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: Web,
    contentTypes: IContentType[],
    context: ProvisioningContext
  ): Promise<void> {
    this.jsomContext = (
      await initSpfxJsom(context.web.ServerRelativeUrl)
    ).jsomContext
    this.context = context
    super.scope_started()
    try {
      this.context.contentTypes = (
        await web.contentTypes
          .select('Id', 'Name', 'FieldLinks')
          .expand('FieldLinks')
          .get<any[]>()
      ).reduce((object, contentType) => {
        object[contentType.Name] = {
          ID: contentType.Id.StringValue,
          Name: contentType.Name,
          FieldRefs: contentType.FieldLinks.map((fieldLink: any) => ({
            ID: fieldLink.Id,
            Name: fieldLink.Name,
            Required: fieldLink.Required,
            Hidden: fieldLink.Hidden
          }))
        }
        return object
      }, {})
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
            chain.then(() => this.processContentType(web, contentType)),
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
   * @param web - The web
   * @param contentType - Content type
   */
  private async processContentType(
    web: Web,
    contentType: IContentType
  ): Promise<void> {
    try {
      let contentTypeId = this.context.contentTypes[contentType.Name].ID
      if (!contentTypeId) {
        const contentTypeAddResult = await this.addContentType(web, contentType)
        contentTypeId = contentTypeAddResult.data.Id
      }
      super.log_info(
        'processContentType',
        `Processing content type [${contentType.Name}] (${contentTypeId})`
      )
      const spContentType = this.jsomContext.web
        .get_contentTypes()
        .getById(contentTypeId)
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
    web: Web,
    contentType: IContentType
  ): Promise<ContentTypeAddResult> {
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
      for (let index = 0; index < contentType.FieldRefs.length; index++) {
        const fieldReference = contentType.FieldRefs[index]
        const existingFieldLink = this.context.contentTypes[
          contentType.Name
        ].FieldRefs.find((fr) => fr.Name === fieldReference.Name)
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
      super.log_info(
        'processContentTypeFieldRefs',
        `Failed to process field refs for content type [${contentType.Name}] (${contentType.ID})`,
        { error: error.args && error.args.get_message() }
      )
    }
  }
}
