import { extractWebUrl, IWeb, SPInstance, spPost } from '@pnp/sp/presets/all'
import { readableErrorMessage } from '../provisioningerror'
import { CsomError } from './csomerror'

export { CsomError }

/**
 * SharePoint's client object model (CSOM) over HTTP: an XML request posted to
 * `_vti_bin/client.svc/ProcessQuery`, for what the REST API cannot do (a content type with the id
 * the template gives it, field links and their order, the term store, the property bag, a group
 * as a group's owner). It is what JSOM sends, without JSOM: the request goes through the web's own
 * PnPjs pipeline, so it works signed in however the web is (SPFx in the browser, a bearer token or
 * cookies in Node).
 *
 * A request names objects by path (`Current.Web.ContentTypes.GetById("0x01…")`) and lists actions
 * on them. SharePoint runs the actions in order and stops at the first that fails; what the ones
 * before it did stays. The answer is 200 either way, with the error in the body.
 */

/** `SP.RequestContext`, whose static `Current` property is the root of every object path. */
const REQUEST_CONTEXT = '{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}'

/** A GUID, as text with or without braces. */
const GUID = /^{?([\da-f]{8}(?:-[\da-f]{4}){3}-[\da-f]{12})}?$/i

/** Characters XML 1.0 cannot hold, not even as a character reference. */
// eslint-disable-next-line no-control-regex
const NOT_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g

/** A GUID parameter ({@link guid}). */
export interface ICsomGuid {
  guid: string
}

/** A parameter that is another object of the request, by its path ({@link reference}). */
export interface ICsomReference {
  path: number
}

/**
 * A parameter that is an object SharePoint builds from its properties, such as a
 * `ContentTypeCreationInformation` ({@link valueObject}).
 */
export interface ICsomValueObject {
  typeId: string
  properties: { [name: string]: CsomValue }
}

/**
 * What a parameter or a property can be: a text, a whole number, a boolean, nothing (`null`), a
 * list of those, a GUID, another object of the request, or a value object.
 */
export type CsomValue =
  | string
  | number
  | boolean
  | null
  | undefined
  | CsomValue[]
  | ICsomGuid
  | ICsomReference
  | ICsomValueObject

/** The results of a request, by the id of the action that produced them. */
export interface ICsomResults {
  [actionId: number]: any
}

/**
 * A GUID parameter. Throws when the text is not a GUID, since it is written into the request.
 *
 * @param value - The GUID, with or without braces
 */
export function guid(value: string): ICsomGuid {
  return { guid: bareGuid(value, 'A GUID parameter') }
}

/**
 * A parameter that is another object of the request.
 *
 * @param path - The object's path, as {@link CsomRequest} returned it
 */
export function reference(path: number): ICsomReference {
  return { path }
}

/**
 * A parameter that is an object SharePoint builds from its properties.
 *
 * @param typeId - The id of its type (`{168f3091-…}` is `ContentTypeCreationInformation`)
 * @param properties - Its properties; they are sent in name order
 */
export function valueObject(
  typeId: string,
  properties: { [name: string]: CsomValue }
): ICsomValueObject {
  return { typeId: `{${bareGuid(typeId, 'A type id')}}`, properties }
}

/**
 * A GUID from a CSOM answer (`/Guid(1b2c…)/`), a REST answer or a template (`{1B2C…}`), bare and
 * in lower case; `undefined` when the text holds none.
 *
 * @param value - The text
 */
export function guidOf(value: any): string | undefined {
  const found = /[\da-f]{8}(?:-[\da-f]{4}){3}-[\da-f]{12}/i.exec(
    typeof value === 'string' ? value : ''
  )
  return found ? found[0].toLowerCase() : undefined
}

/**
 * Whether the answer to an `instantiate` action says the object is not there (a content type, a
 * term set or a term with an id SharePoint does not have).
 *
 * @param result - The action's result
 */
export function isNull(result: any): boolean {
  return !result || result.IsNull === true
}

/**
 * A GUID without braces, in lower case.
 *
 * @param value - The text
 * @param what - What it is, for the error
 */
function bareGuid(value: string, what: string): string {
  const found = GUID.exec(typeof value === 'string' ? value.trim() : '')
  if (!found) throw new Error(`${what} must be a GUID, got ${String(value)}`)
  return found[1].toLowerCase()
}

/**
 * A text as XML content or as an attribute's value.
 *
 * @param text - The text
 */
function escapeXml(text: string): string {
  return String(text)
    .replace(NOT_XML, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\r/g, '&#xD;')
}

/**
 * A value as an XML element: a method's `Parameter`, a value object's `Property` or a list's
 * `Object`.
 *
 * @param tag - The element
 * @param value - The value
 * @param name - The `Name` attribute (a value object's property)
 */
function valueXml(tag: string, value: CsomValue, name?: string): string {
  const open =
    name === undefined ? `<${tag}` : `<${tag} Name="${escapeXml(name)}"`
  if (value === null || value === undefined) return `${open} Type="Null" />`
  if (typeof value === 'string')
    return `${open} Type="String">${escapeXml(value)}</${tag}>`
  if (typeof value === 'boolean')
    return `${open} Type="Boolean">${value ? 'true' : 'false'}</${tag}>`
  if (typeof value === 'number') {
    if (!Number.isInteger(value) || Math.abs(value) > 2147483647)
      throw new Error(`A CSOM number must be a 32-bit integer, got ${value}`)
    return `${open} Type="Int32">${value}</${tag}>`
  }
  if (Array.isArray(value))
    return `${open} Type="Array">${value
      .map((item) => valueXml('Object', item))
      .join('')}</${tag}>`
  if (typeof (value as ICsomGuid).guid === 'string')
    return `${open} Type="Guid">{${bareGuid(
      (value as ICsomGuid).guid,
      'A GUID parameter'
    )}}</${tag}>`
  if (typeof (value as ICsomReference).path === 'number')
    return `${open} ObjectPathId="${pathId((value as ICsomReference).path)}" />`
  const object = value as ICsomValueObject
  if (typeof object.typeId === 'string' && object.properties)
    return `${open} TypeId="${escapeXml(object.typeId)}">${Object.keys(
      object.properties
    )
      .sort()
      .map((key) => valueXml('Property', object.properties[key], key))
      .join('')}</${tag}>`
  throw new Error(`Not a CSOM value: ${JSON.stringify(value)}`)
}

/**
 * A path or action id as it is written into the request.
 *
 * @param id - The id
 */
function pathId(id: number): number {
  if (!Number.isInteger(id) || id < 0)
    throw new Error(`Not an object path of this request: ${id}`)
  return id
}

/**
 * A method's parameters, or nothing when it has none.
 *
 * @param parameters - The parameters
 */
function parametersXml(parameters: CsomValue[]): string {
  return parameters.length > 0
    ? `<Parameters>${parameters
        .map((parameter) => valueXml('Parameter', parameter))
        .join('')}</Parameters>`
    : ''
}

/**
 * The properties a query asks for.
 *
 * @param tag - `Query` (the object's own) or `ChildItemQuery` (each item of a collection)
 * @param properties - The properties
 */
function selectXml(tag: string, properties: string[]): string {
  return `<${tag} SelectAllProperties="false"><Properties>${properties
    .map(
      (property) =>
        `<Property Name="${escapeXml(property)}" ScalarProperty="true" />`
    )
    .join('')}</Properties></${tag}>`
}

/**
 * One CSOM request: object paths (where an object is) and actions (what to do with it). The
 * methods that make a path return the path's id, to build on or to act on; the methods that add an
 * action return the action's id, the key of its result in what {@link executeQuery} returns.
 */
export class CsomRequest {
  private nextId = 0
  private paths: string[] = []
  private actions: string[] = []
  private currentPath: number | undefined
  private webPath: number | undefined

  /**
   * `SP.RequestContext.Current`: the context of the web the request is posted to.
   */
  public current(): number {
    if (this.currentPath === undefined) {
      this.currentPath = this.nextId++
      this.paths.push(
        `<StaticProperty Id="${this.currentPath}" TypeId="${REQUEST_CONTEXT}" Name="Current" />`
      )
    }
    return this.currentPath
  }

  /**
   * The web the request is posted to (`Current.Web`).
   */
  public web(): number {
    if (this.webPath === undefined)
      this.webPath = this.property(this.current(), 'Web')
    return this.webPath
  }

  /**
   * An object that is a property of another (`Web.ContentTypes`).
   *
   * @param parent - The path of the object that has it
   * @param name - The property
   */
  public property(parent: number, name: string): number {
    const id = this.nextId++
    this.paths.push(
      `<Property Id="${id}" ParentId="${pathId(parent)}" Name="${escapeXml(
        name
      )}" />`
    )
    return id
  }

  /**
   * An object a method returns (`ContentTypes.GetById("0x01…")`, `ContentTypes.Add(…)`). The
   * method runs when an action first needs the object.
   *
   * @param parent - The path of the object that has the method
   * @param name - The method
   * @param parameters - Its parameters
   */
  public method(
    parent: number,
    name: string,
    ...parameters: CsomValue[]
  ): number {
    const id = this.nextId++
    const open = `<Method Id="${id}" ParentId="${pathId(
      parent
    )}" Name="${escapeXml(name)}"`
    const inner = parametersXml(parameters)
    this.paths.push(inner ? `${open}>${inner}</Method>` : `${open} />`)
    return id
  }

  /**
   * An object a static method returns (`TaxonomySession.GetTaxonomySession()`).
   *
   * @param typeId - The id of the type that has the method
   * @param name - The method
   * @param parameters - Its parameters
   */
  public staticMethod(
    typeId: string,
    name: string,
    ...parameters: CsomValue[]
  ): number {
    const id = this.nextId++
    const open = `<StaticMethod Id="${id}" Name="${escapeXml(
      name
    )}" TypeId="{${bareGuid(typeId, 'A type id')}}"`
    const inner = parametersXml(parameters)
    this.paths.push(inner ? `${open}>${inner}</StaticMethod>` : `${open} />`)
    return id
  }

  /**
   * Makes SharePoint find or make the object now. Its result is `{ IsNull }`: whether there is no
   * such object ({@link isNull}).
   *
   * @param path - The object
   */
  public instantiate(path: number): number {
    const id = this.nextId++
    this.actions.push(
      `<ObjectPath Id="${id}" ObjectPathId="${pathId(path)}" />`
    )
    return id
  }

  /**
   * Reads properties of an object, and of each item when it is a collection (the items are the
   * result's `_Child_Items_`). The result is `null` for an object that is not there.
   *
   * @param path - The object
   * @param properties - The object's properties to read
   * @param childProperties - The properties to read of each item of a collection
   */
  public query(
    path: number,
    properties: string[],
    childProperties?: string[]
  ): number {
    const id = this.nextId++
    this.actions.push(
      `<Query Id="${id}" ObjectPathId="${pathId(path)}">${selectXml(
        'Query',
        properties
      )}${
        childProperties ? selectXml('ChildItemQuery', childProperties) : ''
      }</Query>`
    )
    return id
  }

  /**
   * Sets a property of an object. Most objects keep it only after their `Update` is called.
   *
   * @param path - The object
   * @param name - The property
   * @param value - The value
   */
  public set(path: number, name: string, value: CsomValue): number {
    const id = this.nextId++
    this.actions.push(
      `<SetProperty Id="${id}" ObjectPathId="${pathId(path)}" Name="${escapeXml(
        name
      )}">${valueXml('Parameter', value)}</SetProperty>`
    )
    return id
  }

  /**
   * Calls a method of an object (`Update(true)`, `Reorder([…])`, `CommitAll()`).
   *
   * @param path - The object
   * @param name - The method
   * @param parameters - Its parameters
   */
  public call(path: number, name: string, ...parameters: CsomValue[]): number {
    const id = this.nextId++
    const open = `<Method Name="${escapeXml(
      name
    )}" Id="${id}" ObjectPathId="${pathId(path)}"`
    const inner = parametersXml(parameters)
    this.actions.push(inner ? `${open}>${inner}</Method>` : `${open} />`)
    return id
  }

  /**
   * How many actions the request has: one without any need not be sent.
   */
  public get actionCount(): number {
    return this.actions.length
  }

  /**
   * The request as XML.
   */
  public toString(): string {
    return [
      '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0"',
      ' ApplicationName="sp-js-provisioning" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">',
      `<Actions>${this.actions.join('')}</Actions>`,
      `<ObjectPaths>${this.paths.join('')}</ObjectPaths>`,
      '</Request>'
    ].join('')
  }
}

/**
 * What a failed request says, for a message a person reads: SharePoint's own words and the
 * error's type for a {@link CsomError}, else whatever the error holds.
 *
 * @param error - What the request threw or rejected with
 */
export function describeCsomError(error: unknown): string {
  if (error instanceof CsomError)
    return error.typeName
      ? `${error.serverMessage} (${error.typeName})`
      : String(error.serverMessage)
  return readableErrorMessage(error) ?? 'unknown error'
}

/**
 * Checks that a principal id is a positive integer, since it is written into the request XML.
 *
 * @param id - The id
 * @param name - What it is, for the error
 */
function principalId(id: number, name: string): number {
  if (!Number.isInteger(id) || id <= 0)
    throw new Error(`${name} must be a positive integer, got ${id}`)
  return id
}

/**
 * The request that makes a site group the owner of another site group: `Group.Owner = group`, then
 * `Group.Update()`. REST's `SetUserAsOwner` answers success for a group id and changes nothing.
 *
 * @param groupId - The group to change
 * @param ownerGroupId - The group that becomes its owner (may be the group itself)
 */
export function setGroupOwnerToGroupRequest(
  groupId: number,
  ownerGroupId: number
): string {
  const request = new CsomRequest()
  const groups = request.property(request.web(), 'SiteGroups')
  const group = request.method(
    groups,
    'GetById',
    principalId(groupId, 'groupId')
  )
  const owner = request.method(
    groups,
    'GetById',
    principalId(ownerGroupId, 'ownerGroupId')
  )
  request.set(group, 'Owner', reference(owner))
  request.call(group, 'Update')
  return request.toString()
}

/**
 * Posts a CSOM request to the web's `_vti_bin/client.svc/ProcessQuery` and throws a
 * {@link CsomError} when SharePoint reports an error in the response (CSOM answers 200 with the
 * error in the body).
 *
 * @param web - The web
 * @param request - The request XML
 */
export async function processQuery(web: IWeb, request: string): Promise<any[]> {
  const webUrl = extractWebUrl(web.toUrl())
  const response = await spPost<any[]>(
    SPInstance([web, webUrl], '_vti_bin/client.svc/ProcessQuery'),
    { body: request, headers: { 'Content-Type': 'text/xml' } }
  )
  const errorInfo =
    Array.isArray(response) && response[0] && response[0].ErrorInfo
  if (errorInfo)
    throw new CsomError(
      errorInfo.ErrorMessage,
      errorInfo.ErrorTypeName,
      errorInfo.ErrorCode
    )
  return response
}

/**
 * Sends a request and returns the results of its actions by action id. Throws a
 * {@link CsomError} when SharePoint reports an error.
 *
 * @param web - The web
 * @param request - The request
 */
export async function executeQuery(
  web: IWeb,
  request: CsomRequest
): Promise<ICsomResults> {
  const response = await processQuery(web, request.toString())
  const results: ICsomResults = {}
  // [header, action id, result, action id, result, ...]
  if (Array.isArray(response))
    for (let index = 1; index + 1 < response.length; index += 2)
      if (typeof response[index] === 'number')
        results[response[index]] = response[index + 1]
  return results
}
