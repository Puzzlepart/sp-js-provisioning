/* eslint-disable unicorn/prevent-abbreviations */
import { combine } from '@pnp/core'
import { BufferParse } from '@pnp/queryable'
import { IWeb, SPInstance, spGet, Web } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { readableErrorMessage } from '../provisioningerror'
import { IFileObject } from '../schema'
import { replaceUrlTokens } from '../util'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'

/** The web, as the handler needs it. */
interface IWebInfo {
  ServerRelativeUrl: string
  Url: string
}

/**
 * Whether two contents hold the same bytes.
 *
 * @param a - One
 * @param b - The other
 */
function sameBytes(a: ArrayBuffer, b: ArrayBuffer): boolean {
  if (a.byteLength !== b.byteLength) return false
  const left = new Uint8Array(a)
  const right = new Uint8Array(b)
  for (let index = 0; index < left.length; index++)
    if (left[index] !== right[index]) return false
  return true
}

/**
 * Whether a request failed because what it asked for is not there.
 *
 * @param error - What it rejected with
 */
function notFound(error: any): boolean {
  return !!error && error.status === 404
}

/**
 * Describes the Files Object Handler.
 *
 * Uploads each file from its source (`Src`, with tokens replaced) into a folder of the web, through
 * REST, and sets its list item's `Properties`. The source is read as bytes: on the web's own
 * origin through the web's pipeline (signed in as the web is), anywhere else with a plain `fetch`
 * without credentials (a public URL). A file whose contents are what the source has is not
 * uploaded again, and a property that has the template's value is not written: an unchanged
 * re-apply writes nothing. A file that fails is logged as an error and the next one is tried.
 * Classic web parts (`WebParts`, `RemoveExistingWebParts`) are no longer supported.
 */
export class Files extends HandlerBase {
  private tokenHelper: TokenHelper

  /**
   * Creates a new instance of the Files class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super('Files', config)
  }

  /**
   * Provisioning Files
   *
   * @param web - The web
   * @param files - The files to provision
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: IWeb,
    files: IFileObject[],
    context?: ProvisioningContext
  ): Promise<void> {
    this.tokenHelper = new TokenHelper(context, this.config)
    super.scope_started()
    let webInfo: IWebInfo
    try {
      // A fresh queryable: `select()` sets the query on the instance itself.
      webInfo = await Web(web).select('ServerRelativeUrl', 'Url')<IWebInfo>()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
    for (const file of files) await this.processFile(web, file, webInfo)
    super.scope_ended()
  }

  /**
   * Provisions a file; a failure is logged, naming the file, and not thrown.
   *
   * @param web - The web
   * @param file - The file
   * @param webInfo - The web's URLs
   */
  private async processFile(
    web: IWeb,
    file: IFileObject,
    webInfo: IWebInfo
  ): Promise<void> {
    const label = [file.Folder, file.Url]
      .map((part) => String(part || '').replace(/^\/+|\/+$/g, ''))
      .filter(Boolean)
      .join('/')
    super.log_info('processFile', `Processing file ${label}`)
    if (
      (file.WebParts && file.WebParts.length > 0) ||
      file.RemoveExistingWebParts
    )
      super.log_error(
        'processFile',
        `File ${label}: classic web parts (WebParts, RemoveExistingWebParts) are no longer supported and are left out; the file itself is provisioned.`
      )
    try {
      const folder = combine('/', webInfo.ServerRelativeUrl, file.Folder)
      const target = combine('/', folder, file.Url)
      const existing = await this.readExisting(web, target)
      if (existing && file.Overwrite === false) {
        super.log_info(
          'processFile',
          `File ${label} exists and is left as it is (Overwrite is false).`
        )
        return
      }
      const source = replaceUrlTokens(
        this.tokenHelper.replaceTokens(file.Src),
        this.config
      )
      const content = await this.readSource(web, webInfo, source)
      if (existing && sameBytes(existing, content))
        super.log_info('processFile', `File ${label} is unchanged.`)
      else {
        await Web(web)
          .getFolderByServerRelativePath(folder)
          .files.addUsingPath(file.Url, content, { Overwrite: true })
        super.log_info(
          'processFile',
          `Uploaded file ${label} (${content.byteLength} bytes)`
        )
      }
      await this.processProperties(web, target, file, label)
    } catch (error) {
      super.log_error(
        'processFile',
        `Failed to provision file ${label}: ${
          readableErrorMessage(error) ?? 'unknown error'
        }`
      )
    }
  }

  /**
   * The contents of the file the web has at a URL, or `undefined` when it has none.
   *
   * @param web - The web
   * @param serverRelativeUrl - The file's URL
   */
  private async readExisting(
    web: IWeb,
    serverRelativeUrl: string
  ): Promise<ArrayBuffer | undefined> {
    try {
      return await Web(web)
        .getFileByServerRelativePath(serverRelativeUrl)
        .getBuffer()
    } catch (error) {
      if (notFound(error)) return
      throw error
    }
  }

  /**
   * A file's source, as bytes. On the web's origin (or server-relative) it is read through the
   * web's pipeline, which signs the request in as the web is; elsewhere with a plain `fetch` and no
   * credentials, which a public URL (one that allows any origin) answers.
   *
   * @param web - The web
   * @param webInfo - The web's URLs
   * @param source - The source URL, tokens replaced
   */
  private async readSource(
    web: IWeb,
    webInfo: IWebInfo,
    source: string
  ): Promise<ArrayBuffer> {
    const origin = new URL(webInfo.Url).origin
    const url = new URL(source, `${origin}/`)
    if (url.origin === origin)
      return spGet<ArrayBuffer>(
        SPInstance([web, url.toString()]).using(BufferParse())
      )
    const response = await fetch(url.toString(), { method: 'GET' })
    if (!response.ok)
      throw new Error(
        `Failed to read ${url.toString()}: ${response.status} ${
          response.statusText
        }`.trim()
      )
    return response.arrayBuffer()
  }

  /**
   * Sets the template's `Properties` on the file's list item, only those whose value differs.
   *
   * @param web - The web
   * @param serverRelativeUrl - The file's URL
   * @param file - The file
   * @param label - The file, for the log
   */
  private async processProperties(
    web: IWeb,
    serverRelativeUrl: string,
    file: IFileObject,
    label: string
  ): Promise<void> {
    const properties = file.Properties || {}
    const keys = Object.keys(properties)
    if (keys.length === 0) return
    const item = await Web(web)
      .getFileByServerRelativePath(serverRelativeUrl)
      .listItemAllFields.select('ID', 'ParentList/Id', ...keys)
      .expand('ParentList')<Record<string, any>>()
    const changed: Record<string, string | number> = {}
    for (const key of keys)
      if (
        item[key] === undefined ||
        item[key] === null ||
        String(item[key]) !== String(properties[key])
      )
        changed[key] = properties[key]
    if (Object.keys(changed).length === 0) {
      super.log_info(
        'processProperties',
        `The properties of file ${label} are unchanged.`
      )
      return
    }
    await Web(web)
      .lists.getById(item.ParentList.Id)
      .items.getById(item.ID)
      .update(changed)
    super.log_info(
      'processProperties',
      `Set properties of file ${label}: ${Object.keys(changed).join(', ')}`
    )
  }
}
