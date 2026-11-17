import {
  ClientsideText,
  ClientsideWebpart,
  CreateClientsidePage,
  IClientsidePage,
  IClientsidePageComponent,
  IWeb,
  Web
} from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { mapStrings, replaceUrlTokens } from '../util'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'
import { IClientSideControl, IClientSidePage } from '../schema'
import '@pnp/sp/presets/all'
import '@pnp/sp/comments/clientside-page'

/**
 * Describes the ClientSidePages Object Handler.
 *
 * Creates each modern page in `SitePages` with its sections, columns and controls, saves and
 * publishes it, and turns its comments on or off. A page the web has is deleted and created anew
 * with `Overwrite: true`, and left as it is otherwise, unless it has never been published: a setup
 * that stopped part way leaves its page so (a draft, often without its content), and Navigation is
 * refused a node to it. Such a page gets the template's title and content, in place, and is
 * published.
 */
export class ClientSidePages extends HandlerBase {
  private tokenHelper: TokenHelper

  /**
   * Creates a new instance of the ObjectClientSidePages class
   */
  constructor(config: IProvisioningConfig) {
    super('ClientSidePages', config)
  }

  /**
   * Provisioning Client Side Pages
   *
   * @param web - The web
   * @param clientSidePages - The client side pages to provision
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: IWeb,
    clientSidePages: IClientSidePage[],
    context?: ProvisioningContext
  ): Promise<void> {
    this.tokenHelper = new TokenHelper(context, this.config)
    super.scope_started()
    try {
      const partDefinitions = await web.getClientsideWebParts()
      await clientSidePages.reduce(
        (chain: Promise<any>, clientSidePage) =>
          chain.then(() =>
            this._processClientSidePage(web, clientSidePage, partDefinitions)
          ),
        Promise.resolve()
      )
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
    super.scope_ended()
  }

  private _getClientSideWebPart(
    partDefinitions: IClientsidePageComponent[],
    control: IClientSideControl
  ) {
    if (control.Id === 'Text' || control.Id === 'PageText') {
      super.log_info(
        'getClientSideWebPart',
        `Control ID is ${control.Id}. Creating a ClientsideText part with text ${control.Text}`
      )
      const textPart = new ClientsideText(control.Text)
      return {
        part: textPart,
        partDef: {
          Name: 'PageText'
        }
      }
    }
    const partDef = partDefinitions.find((c) =>
      c.Id.toLowerCase().includes(control.Id.toLowerCase())
    )
    if (!partDef) {
      super.log_warn(
        'getClientSideWebPart',
        `Client side web part with definition id ${control.Id} not found.`
      )
      return {
        part: null,
        partDef: null
      }
    }
    // Tokens are replaced in each string: a value with a quote or a backslash
    // broke the JSON text they used to be replaced in, and the part was lost.
    const part = ClientsideWebpart.fromComponentDef(partDef).setProperties<any>(
      mapStrings(control.Properties, (text) =>
        replaceUrlTokens(this.tokenHelper.replaceTokens(text), this.config)
      )
    )
    if (control.ServerProcessedContent) {
      part.setServerProcessedContent(
        mapStrings(control.ServerProcessedContent, (text) =>
          this.tokenHelper.replaceTokens(text)
        )
      )
    }
    return { part, partDef } as const
  }

  /**
   * Adds a part to a section or column of a page. A part that cannot be added is logged as an
   * error and the page is saved without it.
   *
   * @param target - The vertical section or column
   * @param part - The part
   * @param partDef - The part's definition (for its name)
   * @param pageName - The page (for logging)
   */
  private _addControl(
    target: { addControl(part: any): any },
    part: any,
    partDef: { Name: string },
    pageName: string
  ): void {
    try {
      target.addControl(part)
    } catch (error) {
      super.log_error(
        'processClientSidePage',
        `Failed adding part ${partDef.Name} to client side page ${pageName}: ${
          (error && error.message) || error
        }`
      )
    }
  }

  /**
   * Whether a page's file has a published (major) version. A page a setup created and did not get
   * to publish (it stopped part way) has none: it is a draft (0.x) that only its author sees.
   *
   * @param web - The web
   * @param serverRelativeUrl - The page's file
   */
  private async _hasPublishedVersion(
    web: IWeb,
    serverRelativeUrl: string
  ): Promise<boolean> {
    // A fresh queryable: `select()` sets the query on the instance itself.
    const { MajorVersion } = await Web(web)
      .getFileByServerRelativePath(serverRelativeUrl)
      .select('MajorVersion')<{ MajorVersion: number }>()
    return MajorVersion > 0
  }

  /**
   * Provision a client side page
   *
   * @param web - The web
   * @param clientSidePage - Cient side page
   * @param partDefinitions - Cient side web parts
   */
  private async _processClientSidePage(
    web: IWeb,
    clientSidePage: IClientSidePage,
    partDefinitions: IClientsidePageComponent[]
  ) {
    super.log_info(
      'processClientSidePage',
      `Processing client side page ${clientSidePage.Name}`
    )
    const { ServerRelativeUrl } = await web.select('ServerRelativeUrl')()
    const serverRelativeUrl = `${ServerRelativeUrl}/SitePages/${clientSidePage.Name}`
    let page: IClientsidePage = null
    try {
      page = await web.loadClientsidePage(serverRelativeUrl)

      if (clientSidePage.Overwrite && page) {
        super.log_info(
          'processClientSidePage',
          `Overwrite option is enabled. Deleting client side page ${clientSidePage.Name}`
        )
        await page.delete()
      }
    } catch {}

    if (!clientSidePage.Overwrite && page) {
      if (await this._hasPublishedVersion(web, serverRelativeUrl)) {
        super.log_info(
          'processClientSidePage',
          `Client side page ${clientSidePage.Name} already exists and overwrite option is disabled. Skipping`
        )
        return
      }
      // Written in place: the page keeps its file, its item and its layout (set when a page is
      // created).
      super.log_info(
        'processClientSidePage',
        `Client side page ${clientSidePage.Name} exists but has never been published, as a setup that stopped part way leaves it: writing the template's title and content to it and publishing it`
      )
      page.title = clientSidePage.Title
      page.sections = []
    } else
      page = await CreateClientsidePage(
        web,
        clientSidePage.Name,
        clientSidePage.Title,
        clientSidePage.PageLayoutType
      )
    if (clientSidePage.VerticalSection) {
      const verticalSection = page.addVerticalSection()
      for (const control of clientSidePage.VerticalSection) {
        const { part, partDef } = this._getClientSideWebPart(
          partDefinitions,
          control
        )
        if (!part) continue
        this._addControl(verticalSection, part, partDef, clientSidePage.Name)
      }
    }
    const sections = clientSidePage.Sections || []
    for (const s of sections) {
      const section = page.addSection()
      for (const col of s.Columns) {
        const column = section.addColumn(col.Factor)
        for (const control of col.Controls) {
          const { part, partDef } = this._getClientSideWebPart(
            partDefinitions,
            control
          )
          if (!part) continue
          this._addControl(column, part, partDef, clientSidePage.Name)
        }
      }
    }
    super.log_info(
      'processClientSidePage',
      `Saving client side page ${clientSidePage.Name}`
    )
    page.commentsDisabled = clientSidePage.CommentsDisabled
    await page.save()
    await (clientSidePage.CommentsDisabled
      ? page.disableComments()
      : page.enableComments())
  }
}
