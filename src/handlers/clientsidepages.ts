import {
  ClientSidePage, ClientSidePageComponent, ClientSideWebpart, Web
} from '@pnp/sp'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { IClientSidePage } from '../schema'
import { replaceUrlTokens } from '../util'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'

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
    web: Web,
    clientSidePages: IClientSidePage[],
    context?: ProvisioningContext
  ): Promise<void> {
    this.tokenHelper = new TokenHelper(context, this.config)
    super.scope_started()
    try {
      const partDefinitions = await web.getClientSideWebParts()
      await clientSidePages.reduce(
        (chain: Promise<any>, clientSidePage) =>
          chain.then(() =>
            this.processClientSidePage(web, clientSidePage, partDefinitions)
          ),
        Promise.resolve()
      )
    } catch (error) {
      // eslint-disable-next-line no-console
      console.log(error)
      super.scope_ended()
      throw error
    }
  }

  /**
   * Provision a client side page
   *
   * @param web - The web
   * @param clientSidePage - Cient side page
   * @param partDefinitions - Cient side web parts
   */
  private async processClientSidePage(
    web: Web,
    clientSidePage: IClientSidePage,
    partDefinitions: ClientSidePageComponent[]
  ) {
    super.log_info(
      'processClientSidePage',
      `Processing client side page ${clientSidePage.Name}`
    )
    const page = await ClientSidePage.create(
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
          const partDef = partDefinitions.find((c) =>
            c.Id.toLowerCase().includes(control.Id.toLowerCase())
          )
          if (!partDef) {
            super.log_warn(
              'processClientSidePage',
              `Client side web part with definition id ${control.Id} not found.`
            )
            continue
          }
          try {
            let properties = this.tokenHelper.replaceTokens(
              JSON.stringify(control.Properties)
            )
            properties = replaceUrlTokens(properties, this.config)
            const part = ClientSideWebpart.fromComponentDef(
              partDef
            ).setProperties<any>(JSON.parse(properties))
            if (control.ServerProcessedContent) {
              const serverProcessedContent = this.tokenHelper.replaceTokens(
                JSON.stringify(control.ServerProcessedContent)
              )
              part.data.webPartData.serverProcessedContent = JSON.parse(
                serverProcessedContent
              )
            }
            super.log_info(
              'processClientSidePage',
              `Adding ${partDef.Name} to client side page ${clientSidePage.Name}`
            )
            column.addControl(part)
          } catch (error) {
            // eslint-disable-next-line no-console
            console.log(error)
            super.log_info(
              'processClientSidePage',
              `Failed adding part ${partDef.Name} to client side page ${clientSidePage.Name}`
            )
          }
        }
      }
    }
    super.log_info(
      'processClientSidePage',
      `Saving client side page ${clientSidePage.Name}`
    )
    await page.save()
    if (clientSidePage.CommentsDisabled) {
      super.log_info(
        'processClientSidePage',
        `Disabling comments for client side page ${clientSidePage.Name}`
      )
      await page.disableComments()
    }
  }
}
