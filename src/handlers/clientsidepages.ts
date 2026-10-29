import {
  ClientsideText,
  ClientsideWebpart,
  CreateClientsidePage,
  IClientsidePage,
  IClientsidePageComponent,
  IWeb
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
 * Describes the Composed Look Object Handler
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
    let page: IClientsidePage = null
    try {
      page = await web.loadClientsidePage(
        `${ServerRelativeUrl}/SitePages/${clientSidePage.Name}`
      )

      if (clientSidePage.Overwrite && page) {
        super.log_info(
          'processClientSidePage',
          `Overwrite option is enabled. Deleting client side page ${clientSidePage.Name}`
        )
        await page.delete()
      }
    } catch {}

    if (!clientSidePage.Overwrite && page) {
      super.log_info(
        'processClientSidePage',
        `Client side page ${clientSidePage.Name} already exists and overwrite option is disabled. Skipping`
      )
      return
    }

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
