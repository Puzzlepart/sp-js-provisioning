import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { IClientSideControl, IClientSidePage } from '../schema'
import { replaceUrlTokens } from '../util'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'
import { CreateClientsidePage, IClientsidePageComponent, ClientsideWebpart } from '@pnp/sp/clientside-pages'
import { IWeb } from '@pnp/sp/webs/types'

/**
 * Describes the Composed Look Object Handler
 */
export class ClientSidePages extends HandlerBase {
  private tokenHelper: TokenHelper

  /**
   * Creates a new instance of the ObjectClientSidePages class
   */
  constructor(config: IProvisioningConfig) {
    super(ClientSidePages.name, config)
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
      const clientsideWebParts = await web.getClientsideWebParts()
      await clientSidePages.reduce(
        (chain: Promise<any>, clientSidePage) =>
          chain.then(() =>
            this.processClientSidePage(web, clientSidePage, clientsideWebParts)
          ),
        Promise.resolve()
      )
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Provision a client side page
   *
   * @param web - The web
   * @param clientSidePage - Cient side page
   * @param clientsideWebParts - Cient side web parts
   */
  private async processClientSidePage(
    web: IWeb,
    clientSidePage: IClientSidePage,
    clientsideWebParts: IClientsidePageComponent[]
  ) {
    super.log_info(
      'processClientSidePage',
      `Processing client side page ${clientSidePage.Name}`
    )
    const page = await CreateClientsidePage(
      web,
      clientSidePage.Name,
      clientSidePage.Title,
      clientSidePage.PageLayoutType
    )
    const sections = clientSidePage.Sections || []
    for (const s of sections) {
      const section = page.addSection()
      for (const col of s.Columns) {
        const column = section.addColumn(col.Factor)
        for (const control of col.Controls) {
          try {
            const columnControl = this.getColumnControl(control, clientsideWebParts)
            if (columnControl !== null) column.addControl(columnControl)
          } catch {}
        }
      }
    }
    if (clientSidePage.VerticalSection) {
      const a = page.addVerticalSection()
      for (const control of clientSidePage.VerticalSection) {
        const columnControl = this.getColumnControl(control, clientsideWebParts)
        a.addControl(columnControl)
      }
    }
    super.log_info(
      'processClientSidePage',
      `Saving client side page ${clientSidePage.Name}`
    )
    page.commentsDisabled = clientSidePage.CommentsDisabled
    await page.save()
  }

  /**
   * Provision a client side page
   *
   * @param control - Control
   * @param clientsideWebParts - Cient side web parts
   */
  private getColumnControl(control: IClientSideControl, clientsideWebParts: IClientsidePageComponent[]) {
    const partDefinition = clientsideWebParts.find((c) =>
      c.Id.toLowerCase().includes(control.Id.toLowerCase())
    )
    if (!partDefinition) {
      super.log_warn(
        'processClientSidePage',
        `Client side web part with definition id ${control.Id} not found.`
      )
      return null
    }
    let properties = this.tokenHelper.replaceTokens(
      JSON.stringify(control.Properties)
    )
    properties = replaceUrlTokens(properties, this.config)
    const part = ClientsideWebpart.fromComponentDef(
      partDefinition
    ).setProperties<any>(JSON.parse(properties))
    if (control.ServerProcessedContent) {
      const serverProcessedContent = this.tokenHelper.replaceTokens(
        JSON.stringify(control.ServerProcessedContent)
      )
      part.data.webPartData.serverProcessedContent = JSON.parse(
        serverProcessedContent
      )
    }
    return part
  }
}
