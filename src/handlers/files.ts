/* eslint-disable unicorn/prevent-abbreviations */
import * as xmljs from 'xml-js'
import { HandlerBase } from './handlerbase'
import { IFile, IWebPart } from '../schema'
import { Web, File, FileAddResult } from '@pnp/sp'
import { combine, isArray } from '@pnp/common'
import { Logger, LogLevel } from '@pnp/logging'
import { replaceUrlTokens } from '../util'
import { ProvisioningContext } from '../provisioningcontext'
import { IProvisioningConfig } from '../provisioningconfig'
import { TokenHelper } from '../util/tokenhelper'

/**
 * Describes the Features Object Handler
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
   * @param files - The files  to provision
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: Web,
    files: IFile[],
    context?: ProvisioningContext
  ): Promise<void> {
    this.tokenHelper = new TokenHelper(context, this.config)
    super.scope_started()
    if (this.config.spfxContext) {
      throw 'Files Handler not supported in SPFx.'
    }
    const { ServerRelativeUrl } = await web.get()
    try {
      await files.reduce(
        (chain: any, file) =>
          chain.then(() => this.processFile(web, file, ServerRelativeUrl)),
        Promise.resolve()
      )
      super.scope_ended()
    } catch {
      super.scope_ended()
    }
  }

  /**
   * Get blob for a file
   *
   * @param file - The file
   */
  private async getFileBlob(file: IFile): Promise<Blob> {
    const fileSourceWithoutTokens = replaceUrlTokens(
      this.tokenHelper.replaceTokens(file.Src),
      this.config
    )
    const response = await fetch(fileSourceWithoutTokens, {
      credentials: 'include',
      method: 'GET'
    })
    const fileContents = await response.text()
    const blob = new Blob([fileContents], { type: 'text/plain' })
    return blob
  }

  /**
   * Procceses a file
   *
   * @param web - The web
   * @param file - The file
   * @param webServerRelativeUrl - ServerRelativeUrl for the web
   */
  private async processFile(
    web: Web,
    file: IFile,
    webServerRelativeUrl: string
  ): Promise<void> {
    Logger.log({
      level: LogLevel.Info,
      message: `Processing file ${file.Folder}/${file.Url}`
    })
    try {
      const blob = await this.getFileBlob(file)
      const folderServerRelativeUrl = combine(
        '/',
        webServerRelativeUrl,
        file.Folder
      )
      const pnpFolder = web.getFolderByServerRelativeUrl(
        folderServerRelativeUrl
      )
      let fileServerRelativeUrl = combine(
        '/',
        folderServerRelativeUrl,
        file.Url
      )
      let fileAddResult: FileAddResult
      let pnpFile: File
      try {
        fileAddResult = await pnpFolder.files.add(
          file.Url,
          blob,
          file.Overwrite
        )
        pnpFile = fileAddResult.file
        fileServerRelativeUrl = fileAddResult.data.ServerRelativeUrl
      } catch {
        pnpFile = web.getFileByServerRelativePath(fileServerRelativeUrl)
      }
      await this.processProperties(web, pnpFile, file)
      await this.processWebParts(
        file,
        webServerRelativeUrl,
        fileServerRelativeUrl
      )
      await this.processPageListViews(web, file.WebParts, fileServerRelativeUrl)
    } catch (error) {
      throw error
    }
  }

  /**
   * Remove exisiting webparts if specified
   *
   * @param webServerRelativeUrl - ServerRelativeUrl for the web
   * @param fileServerRelativeUrl - ServerRelativeUrl for the file
   * @param shouldRemove - Should web parts be removed
   */
  private removeExistingWebParts(
    webServerRelativeUrl: string,
    fileServerRelativeUrl: string,
    shouldRemove: boolean
  ) {
    return new Promise((resolve, reject) => {
      if (shouldRemove) {
        Logger.log({
          level: LogLevel.Info,
          message: `Deleting existing webpart from file ${fileServerRelativeUrl}`
        })
        const clientContext = new SP.ClientContext(webServerRelativeUrl)
        const spFile = clientContext
          .get_web()
          .getFileByServerRelativeUrl(fileServerRelativeUrl)
        const webPartManager = spFile.getLimitedWebPartManager(
          SP.WebParts.PersonalizationScope.shared
        )
        const webParts = webPartManager.get_webParts()
        clientContext.load(webParts)
        clientContext.executeQueryAsync(() => {
          for (const wp of webParts.get_data()) wp.deleteWebPart()
          clientContext.executeQueryAsync(resolve, reject)
        }, reject)
      } else {
        Logger.log({
          level: LogLevel.Info,
          message: `Web parts should not be removed from file ${fileServerRelativeUrl}.`
        })
        resolve(true)
      }
    })
  }

  /**
   * Processes web parts
   *
   * @param file - The file
   * @param webServerRelativeUrl - ServerRelativeUrl for the web
   * @param fileServerRelativeUrl - ServerRelativeUrl for the file
   */
  private processWebParts(
    file: IFile,
    webServerRelativeUrl: string,
    fileServerRelativeUrl: string
  ) {
    return new Promise(async (resolve, reject) => {
      Logger.log({
        level: LogLevel.Info,
        message: `Processing webparts for file ${file.Folder}/${file.Url}`
      })
      await this.removeExistingWebParts(
        webServerRelativeUrl,
        fileServerRelativeUrl,
        file.RemoveExistingWebParts
      )
      if (file.WebParts && file.WebParts.length > 0) {
        const clientContext = new SP.ClientContext(webServerRelativeUrl),
          spFile = clientContext
            .get_web()
            .getFileByServerRelativeUrl(fileServerRelativeUrl),
          webPartManager = spFile.getLimitedWebPartManager(
            SP.WebParts.PersonalizationScope.shared
          )
        await this.fetchWebPartContents(file.WebParts, (index, xml) => {
          file.WebParts[index].Contents.Xml = xml
        })
        for (const wp of file.WebParts) {
          const webPartXml = this.tokenHelper.replaceTokens(
            this.replaceWebPartXmlTokens(wp.Contents.Xml, clientContext)
          )
          const webPartDef = webPartManager.importWebPart(webPartXml)
          const webPartInstance = webPartDef.get_webPart()
          Logger.log({
            data: { webPartXml },
            level: LogLevel.Info,
            message: `Processing webpart ${wp.Title} for file ${file.Folder}/${file.Url}`
          })
          webPartManager.addWebPart(webPartInstance, wp.Zone, wp.Order)
          clientContext.load(webPartInstance)
        }
        clientContext.executeQueryAsync(resolve, (sender, args) => {
          Logger.log({
            data: { error: args.get_message() },
            level: LogLevel.Error,
            message: `Failed to process webparts for file ${file.Folder}/${file.Url}`
          })
          reject({ sender, args })
        })
      } else {
        resolve(true)
      }
    })
  }

  /**
   * Fetches web part contents
   *
   * @param webParts - Web parts
   * @param callbackFunc - Callback function that takes index of the the webpart and the retrieved XML
   */
  private fetchWebPartContents = (
    webParts: Array<IWebPart>,
    callbackFunction: (index, xml) => void
  ) => {
    return new Promise<any>((resolve, reject) => {
      const fileFetchPromises = webParts.map((wp, index) => {
        return (() => {
          return new Promise<any>(async (_res) => {
            if (wp.Contents.FileSrc) {
              const fileSource = replaceUrlTokens(
                this.tokenHelper.replaceTokens(wp.Contents.FileSrc),
                this.config
              )
              Logger.log({
                data: null,
                level: LogLevel.Info,
                message: `Retrieving contents from file '${fileSource}'.`
              })
              const response = await fetch(fileSource, {
                credentials: 'include',
                method: 'GET'
              })
              const xml = await response.text()
              if (isArray(wp.PropertyOverrides)) {
                const object: any = xmljs.xml2js(xml)
                if (object.elements[0].name === 'webParts') {
                  const existingProperties =
                    object.elements[0].elements[0].elements[1].elements[0]
                      .elements
                  const updatedProperties = []
                  for (const property of existingProperties) {
                    const hasOverride =
                      wp.PropertyOverrides.filter(
                        (po) => po.name === property.attributes.name
                      ).length > 0
                    if (!hasOverride) {
                      updatedProperties.push(property)
                    }
                  }
                  // eslint-disable-next-line unicorn/no-array-for-each
                  wp.PropertyOverrides.forEach(({ name, type, value }) => {
                    updatedProperties.push({
                      attributes: { name, type },
                      elements: [{ text: value, type: 'text' }],
                      name: 'property',
                      type: 'element'
                    })
                  })
                  object.elements[0].elements[0].elements[1].elements[0].elements = updatedProperties
                  callbackFunction(index, xmljs.js2xml(object))
                  _res(true)
                } else {
                  callbackFunction(index, xml)
                  _res(true)
                }
              } else {
                callbackFunction(index, xml)
                _res(true)
              }
            } else {
              _res(true)
            }
          })
        })()
      })
      Promise.all(fileFetchPromises).then(resolve).catch(reject)
    })
  }

  /**
   * Processes page list views
   *
   * @param web - The web
   * @param webParts - Web parts
   * @param fileServerRelativeUrl - ServerRelativeUrl for the file
   */
  private processPageListViews(
    web: Web,
    webParts: Array<IWebPart>,
    fileServerRelativeUrl: string
  ): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      if (webParts) {
        Logger.log({
          data: { webParts, fileServerRelativeUrl },
          level: LogLevel.Info,
          message: `Processing page list views for file ${fileServerRelativeUrl}`
        })
        const listViewWebParts = webParts.filter((wp) => wp.ListView)
        if (listViewWebParts.length > 0) {
          listViewWebParts
            .reduce(
              (chain: any, wp) =>
                chain.then(() =>
                  this.processPageListView(
                    web,
                    wp.ListView,
                    fileServerRelativeUrl
                  )
                ),
              Promise.resolve()
            )
            .then(() => {
              Logger.log({
                data: {},
                level: LogLevel.Info,
                message: `Successfully processed page list views for file ${fileServerRelativeUrl}`
              })
              resolve()
            })
            .catch((error) => {
              Logger.log({
                data: { err: error, fileServerRelativeUrl },
                level: LogLevel.Error,
                message: `Failed to process page list views for file ${fileServerRelativeUrl}`
              })
              reject(error)
            })
        } else {
          resolve()
        }
      } else {
        resolve()
      }
    })
  }

  /**
   * Processes page list view
   *
   * @param web - The web
   * @param listView - List view
   * @param fileServerRelativeUrl - ServerRelativeUrl for the file
   */
  private processPageListView(
    web: Web,
    listView,
    fileServerRelativeUrl: string
  ) {
    return new Promise<void>((resolve, reject) => {
      const views = web.lists.getByTitle(listView.List).views
      views
        .get()
        .then((listViews) => {
          const wpView = listViews.filter(
            (v) => v.ServerRelativeUrl === fileServerRelativeUrl
          )
          if (wpView.length === 1) {
            const view = views.getById(wpView[0].Id)
            const settings = listView.View.AdditionalSettings || {}
            view
              .update(settings)
              .then(() => {
                view.fields
                  .removeAll()
                  .then(() => {
                    listView.View.ViewFields.reduce(
                      (chain, viewField) =>
                        chain.then(() => view.fields.add(viewField)),
                      Promise.resolve()
                    )
                      .then(resolve)
                      .catch((error) => {
                        Logger.log({
                          data: { fileServerRelativeUrl, listView, err: error },
                          level: LogLevel.Error,
                          message: `Failed to process page list view for file ${fileServerRelativeUrl}`
                        })
                        reject(error)
                      })
                  })
                  .catch((error) => {
                    Logger.log({
                      data: { fileServerRelativeUrl, listView, err: error },
                      level: LogLevel.Error,
                      message: `Failed to process page list view for file ${fileServerRelativeUrl}`
                    })
                    reject(error)
                  })
              })
              .catch((error) => {
                Logger.log({
                  data: { fileServerRelativeUrl, listView, err: error },
                  level: LogLevel.Error,
                  message: `Failed to process page list view for file ${fileServerRelativeUrl}`
                })
                reject(error)
              })
          } else {
            resolve()
          }
        })
        .catch((error) => {
          Logger.log({
            data: { fileServerRelativeUrl, listView, err: error },
            level: LogLevel.Error,
            message: `Failed to process page list view for file ${fileServerRelativeUrl}`
          })
          reject(error)
        })
    })
  }

  /**
   * Process list item properties for the file
   *
   * @param web - The web
   * @param pnpFile - The PnP file
   * @param properties - The properties to set
   */
  private async processProperties(web: Web, pnpFile: File, file: IFile) {
    const hasProperties =
      file.Properties && Object.keys(file.Properties).length > 0
    if (hasProperties) {
      Logger.log({
        level: LogLevel.Info,
        message: `Processing properties for ${file.Folder}/${file.Url}`
      })
      const listItemAllFields = await pnpFile.listItemAllFields
        .select('ID', 'ParentList/ID', 'ParentList/Title')
        .expand('ParentList')
        .get()
      await web.lists
        .getById(listItemAllFields.ParentList.Id)
        .items.getById(listItemAllFields.ID)
        .update(file.Properties)
      Logger.log({
        level: LogLevel.Info,
        message: `Successfully processed properties for ${file.Folder}/${file.Url}`
      })
    }
  }

  /**
   * Replaces tokens in a string, e.g. `{site}`
   *
   * @param str - The string
   * @param ctx - Client context
   */
  private replaceWebPartXmlTokens(
    string: string,
    context: SP.ClientContext
  ): string {
    const site = combine(
      document.location.protocol,
      '//',
      document.location.host,
      context.get_url()
    )
    return string.replace(/{site}/g, site)
  }
}
