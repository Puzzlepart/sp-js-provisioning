/* eslint-disable unicorn/prevent-abbreviations */
import { IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { IContentType, IFieldReference } from '../schema'
import {
  CsomRequest,
  describeCsomError,
  executeQuery,
  guid,
  guidOf,
  reference,
  valueObject
} from '../util/csom'
import { HandlerBase } from './handlerbase'

/** `SP.ContentTypeCreationInformation`: what a new content type is made from. */
const CONTENT_TYPE_CREATION_INFORMATION =
  '{168f3091-4554-4f14-8866-b20d48e45b54}'

/** `SP.FieldLinkCreationInformation`: what a new field link is made from. */
const FIELD_LINK_CREATION_INFORMATION = '{63fb2c92-8f65-4bbb-a658-b6cd294403f4}'

/** A site column the web can use: its own, and on a subweb its parent webs'. */
interface IAvailableField {
  Id: string
  InternalName: string
  Title: string
}

/** A field link of a content type, as SharePoint has it. */
interface IFieldLink extends IFieldReference {
  Name: string
  Required: boolean
  Hidden: boolean
}

/**
 * Describes the Content Types Object Handler.
 *
 * A content type is created with the id the template gives it, which the REST API cannot do
 * (it makes a child of Item with an id of its own), so the writes go through CSOM over HTTP
 * (`src/util/csom.ts`). What the web has is read through REST first, and only what differs is
 * written: a name, description or group, a missing field link, a changed `Required` or `Hidden`,
 * a changed order. A content type that needs nothing is not sent `Update(true)`, which pushes it to
 * every content type and list that inherits from it.
 */
export class ContentTypes extends HandlerBase {
  public context: ProvisioningContext
  private availableFields: Promise<IAvailableField[]> | undefined

  /**
   * Creates a new instance of the ContentTypes class
   */
  constructor(config: IProvisioningConfig) {
    super('ContentTypes', config)
  }

  /**
   * Provisioning Content Types, parents before their children (in id order)
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
    this.context = context
    this.availableFields = undefined
    super.scope_started()
    try {
      await this._initContext(web)
      const sorted = contentTypes.slice().sort((a, b) => {
        if (a.ID < b.ID) {
          return -1
        }
        if (a.ID > b.ID) {
          return 1
        }
        return 0
      })
      for (const contentType of sorted)
        await this.processContentType(web, contentType)
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Provision a content type: create it when the web does not have it, then write what differs.
   *
   * @param web - The web
   * @param contentType - Content type
   */
  private async processContentType(
    web: IWeb,
    contentType: IContentType
  ): Promise<void> {
    const contentTypeId =
      contentType.ID ?? this.context.contentTypes[contentType.Name]?.ID
    if (!contentTypeId)
      throw new Error(
        `Content type with name '${contentType.Name}' does not exist in the web.`
      )

    const existing = this.existingContentType(contentTypeId)
    const current =
      existing ??
      (await this.createContentType(web, contentType, contentTypeId))

    super.log_info(
      'processContentType',
      `Processing content type [${contentType.Name}] (${contentTypeId})`
    )
    const fieldLinks = await this.updateContentType(
      web,
      contentType,
      current.ID,
      current
    )

    if (current.Name !== contentType.Name)
      delete this.context.contentTypes[current.Name]
    this.context.contentTypes[current.ID] = {
      ID: current.ID,
      Name: contentType.Name,
      Description: contentType.Description || current.Description || '',
      Group: contentType.Group || current.Group || '',
      FieldRefs: fieldLinks
    }
    this.context.contentTypes[contentType.Name] = this.context.contentTypes[
      current.ID
    ]
  }

  /**
   * The content type the web has with an id, whatever the case of the id's hex digits.
   *
   * @param contentTypeId - The id
   */
  private existingContentType(contentTypeId: string): IContentType | undefined {
    const found = this.context.contentTypes[contentTypeId]
    if (found) return found
    const id = contentTypeId.toUpperCase()
    for (const key of Object.keys(this.context.contentTypes)) {
      const candidate = this.context.contentTypes[key]
      if (candidate.ID && candidate.ID.toUpperCase() === id) return candidate
    }
  }

  /**
   * Creates a content type with the template's id and returns it as SharePoint stored it, with
   * the field links it inherits.
   *
   * @param web - The web
   * @param contentType - Content type
   * @param contentTypeId - Its id
   */
  private async createContentType(
    web: IWeb,
    contentType: IContentType,
    contentTypeId: string
  ): Promise<IContentType> {
    super.log_info(
      'createContentType',
      `Creating content type [${contentType.Name}] (${contentTypeId})`
    )
    const request = new CsomRequest()
    // No parent: SharePoint refuses `Id` together with `ParentContentType`
    // ("parameters.Id, parameters.ParentContentType cannot be used together") and reads the
    // parent out of the id.
    const created = request.method(
      request.property(request.web(), 'ContentTypes'),
      'Add',
      valueObject(CONTENT_TYPE_CREATION_INFORMATION, {
        Description: contentType.Description || null,
        Group: contentType.Group || null,
        Id: contentTypeId,
        Name: contentType.Name,
        ParentContentType: null
      })
    )
    request.instantiate(created)
    const read = request.query(created, ['StringId'])
    let storedId: string
    try {
      const results = await executeQuery(web, request)
      storedId = results[read] && results[read].StringId
    } catch (error) {
      throw new Error(
        `Failed to create content type [${
          contentType.Name
        }] (${contentTypeId}): ${describeCsomError(error)}`
      )
    }
    if (
      typeof storedId !== 'string' ||
      storedId.toUpperCase() !== contentTypeId.toUpperCase()
    )
      throw new Error(
        `Content type [${contentType.Name}] was created with the id ${storedId}, not ${contentTypeId} as the template says.`
      )
    return this.toContentType(
      await web.contentTypes
        .getById(storedId)
        .select('Id', 'Name', 'Description', 'Group', 'FieldLinks')
        .expand('FieldLinks')()
    )
  }

  /**
   * The field link a field reference names: by internal name, or by the column's id.
   *
   * @param fieldLinks - The content type's field links
   * @param fieldReference - The field reference
   */
  private getExistingFieldLink(
    fieldLinks: IFieldLink[],
    fieldReference: IFieldReference
  ): IFieldLink | undefined {
    const fieldReferenceId = guidOf(fieldReference.ID)
    return fieldLinks.find(
      (fieldLink) =>
        (!!fieldReference.Name && fieldLink.Name === fieldReference.Name) ||
        (!!fieldReferenceId && guidOf(fieldLink.ID) === fieldReferenceId)
    )
  }

  /**
   * The site column a field reference names, among the columns the web can use: by internal
   * name, else by id, else by title. Read once per run, and only when a link has to be added.
   *
   * @param web - The web
   * @param fieldReference - The field reference
   */
  private async getAvailableField(
    web: IWeb,
    fieldReference: IFieldReference
  ): Promise<IAvailableField | undefined> {
    if (!this.availableFields)
      // Without `top()`: the collection is not paged, and `$top` truncates it.
      this.availableFields = web.availablefields.select(
        'Id',
        'InternalName',
        'Title'
      )<IAvailableField[]>()
    const fields = await this.availableFields
    const id = guidOf(fieldReference.ID)
    const name = fieldReference.Name
    return (
      (name && fields.find((field) => field.InternalName === name)) ||
      (id && fields.find((field) => guidOf(field.Id) === id)) ||
      (name && fields.find((field) => field.Title === name)) ||
      undefined
    )
  }

  /**
   * Writes what the template has differently from the content type: name, description and group,
   * the field links it lacks, `Required` and `Hidden` where the template gives them, and the order
   * of the links, in one request that ends with `Update(true)`. Nothing differs: no request.
   *
   * @param web - The web
   * @param contentType - Content type
   * @param contentTypeId - Its id, as SharePoint has it
   * @param current - The content type as the web has it
   * @returns The content type's field links after the update
   */
  private async updateContentType(
    web: IWeb,
    contentType: IContentType,
    contentTypeId: string,
    current: IContentType
  ): Promise<IFieldLink[]> {
    const label = `[${contentType.Name}] (${contentTypeId})`
    const request = new CsomRequest()
    const spContentType = request.method(
      request.property(request.web(), 'ContentTypes'),
      'GetById',
      contentTypeId
    )
    let fieldLinksPath: number | undefined
    const spFieldLinks = () =>
      fieldLinksPath === undefined
        ? (fieldLinksPath = request.property(spContentType, 'FieldLinks'))
        : fieldLinksPath
    let fieldsPath: number | undefined
    const spFields = () =>
      fieldsPath === undefined
        ? (fieldsPath = request.property(request.web(), 'AvailableFields'))
        : fieldsPath

    if (current.Name !== contentType.Name)
      request.set(spContentType, 'Name', contentType.Name)
    if (
      contentType.Description &&
      contentType.Description !== current.Description
    )
      request.set(spContentType, 'Description', contentType.Description)
    if (contentType.Group && contentType.Group !== current.Group)
      request.set(spContentType, 'Group', contentType.Group)

    let fieldLinks = (current.FieldRefs || []).map(
      (fieldLink): IFieldLink => ({
        ID: fieldLink.ID,
        Name: fieldLink.Name,
        Required: !!fieldLink.Required,
        Hidden: !!fieldLink.Hidden
      })
    )
    if (contentType.FieldRefs) {
      const named: string[] = []
      for (const fieldReference of contentType.FieldRefs) {
        const wantsRequired = typeof fieldReference.Required === 'boolean'
        const wantsHidden = typeof fieldReference.Hidden === 'boolean'
        const existingFieldLink = this.getExistingFieldLink(
          fieldLinks,
          fieldReference
        )
        if (existingFieldLink) {
          if (!named.includes(existingFieldLink.Name))
            named.push(existingFieldLink.Name)
          const required =
            wantsRequired &&
            fieldReference.Required !== existingFieldLink.Required
          const hidden =
            wantsHidden && fieldReference.Hidden !== existingFieldLink.Hidden
          if (!required && !hidden) continue
          const fieldLink = request.method(
            spFieldLinks(),
            'GetById',
            guid(existingFieldLink.ID)
          )
          if (required) {
            request.set(fieldLink, 'Required', fieldReference.Required)
            existingFieldLink.Required = fieldReference.Required
          }
          if (hidden) {
            request.set(fieldLink, 'Hidden', fieldReference.Hidden)
            existingFieldLink.Hidden = fieldReference.Hidden
          }
          continue
        }
        const siteField = await this.getAvailableField(web, fieldReference)
        if (!siteField) {
          // SiteFields logs a column SharePoint refuses and goes on: the content types that name
          // it get their other columns.
          super.log_error(
            'processContentTypeFieldRefs',
            `Field ref ${
              fieldReference.Name || fieldReference.ID
            } of content type ${label} is left out: the web has no site column with that ${[
              fieldReference.Name ? 'name' : '',
              fieldReference.ID ? 'id' : ''
            ]
              .filter(Boolean)
              .join(' or ')}.`
          )
          continue
        }
        super.log_info(
          'processContentTypeFieldRefs',
          `Adding field ref ${siteField.InternalName} to content type ${label}`
        )
        const fieldLink = request.method(
          spFieldLinks(),
          'Add',
          valueObject(FIELD_LINK_CREATION_INFORMATION, {
            Field: reference(
              request.method(spFields(), 'GetById', guid(siteField.Id))
            )
          })
        )
        request.instantiate(fieldLink)
        if (wantsRequired)
          request.set(fieldLink, 'Required', fieldReference.Required)
        if (wantsHidden) request.set(fieldLink, 'Hidden', fieldReference.Hidden)
        fieldLinks.push({
          ID: siteField.Id,
          Name: siteField.InternalName,
          Required: !!fieldReference.Required,
          Hidden: !!fieldReference.Hidden
        })
        named.push(siteField.InternalName)
      }

      // `Reorder` puts the links it names first, in that order, and the others after them: it is
      // sent only when that is not how they already are.
      if (
        named.length > 1 &&
        named.some((name, index) => fieldLinks[index].Name !== name)
      ) {
        super.log_info(
          'processContentTypeFieldRefs',
          `Reordering field refs for content type ${label}`
        )
        request.call(spFieldLinks(), 'Reorder', named)
        fieldLinks = [
          ...named.map((name) =>
            fieldLinks.find((fieldLink) => fieldLink.Name === name)
          ),
          ...fieldLinks.filter((fieldLink) => !named.includes(fieldLink.Name))
        ]
      }
    }

    if (request.actionCount === 0) {
      super.log_info(
        'processContentType',
        `Content type ${label} is unchanged.`
      )
      return fieldLinks
    }
    // `true`: the content types and lists that inherit from it follow.
    request.call(spContentType, 'Update', true)
    try {
      await executeQuery(web, request)
    } catch (error) {
      throw new Error(
        `Failed to update content type ${label}: ${describeCsomError(error)}`
      )
    }
    super.log_info(
      'processContentType',
      `Successfully updated content type ${label}`
    )
    return fieldLinks
  }

  /**
   * A content type as the REST API returns it, in the shape the context keeps.
   *
   * @param contentType - The content type, with its `FieldLinks`
   */
  private toContentType(contentType: any): IContentType {
    return {
      ID: contentType.Id.StringValue,
      Name: contentType.Name,
      Description: contentType.Description ?? '',
      Group: contentType.Group ?? '',
      FieldRefs: (contentType['FieldLinks'] || []).map((fieldLink: any) => ({
        ID: fieldLink.Id,
        Name: fieldLink.Name,
        Required: fieldLink.Required,
        Hidden: fieldLink.Hidden
      }))
    }
  }

  private async _initContext(web: IWeb): Promise<void> {
    this.context.contentTypes = (
      await web.contentTypes
        .select('Id', 'Name', 'Description', 'Group', 'FieldLinks')
        .expand('FieldLinks')()
    ).reduce((object, contentType) => {
      const ct = this.toContentType(contentType)
      object[ct.Name] = ct
      object[ct.ID] = ct
      return object
    }, {})
  }
}
