import { IFieldAddResult, IWeb } from '@pnp/sp/presets/all'
import * as xmljs from 'xml-js'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'

/**
 * Describes the Site Fields Object Handler
 */
export class SiteFields extends HandlerBase {
  private context: ProvisioningContext
  private tokenHelper: TokenHelper

  /**
   * Creates a new instance of the ObjectSiteFields class
   */
  constructor(config: IProvisioningConfig) {
    super('SiteFields', config)
  }

  /**
   * Provisioning Client Side Pages
   *
   * @param web - The web
   * @param siteFields - The site fields
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: IWeb,
    siteFields: string[],
    context?: ProvisioningContext
  ): Promise<void> {
    this.context = context
    this.tokenHelper = new TokenHelper(this.context, this.config)
    super.scope_started()
    try {
      this.context.siteFields = (
        await web.fields
          .select('Id', 'InternalName')
          <Array<{ Id: string; InternalName: string }>>()
      ).reduce((object, l) => {
        object[l.InternalName] = l.Id
        return object
      }, {})
      await siteFields.reduce(
        (chain: any, schemaXml) =>
          chain.then(() => this.processSiteField(web, schemaXml)),
        Promise.resolve()
      )
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Provision a site field
   *
   * @param web - The web
   * @param clientSidePage - Cient side page
   */
  private async processSiteField(
    web: IWeb,
    schemaXml: string
  ): Promise<IFieldAddResult> {
    try {
      schemaXml = this.tokenHelper.replaceTokens(schemaXml)
      const schemaXmlJson = JSON.parse(xmljs.xml2json(schemaXml))
      const { DisplayName, Name } = schemaXmlJson.elements[0].attributes
      if (this.context.siteFields[Name]) {
        super.log_info('processSiteField', `Updating site field ${DisplayName}`)
        return await web.fields
          .getByInternalNameOrTitle(Name)
          .update({ SchemaXml: schemaXml })
      } else {
        super.log_info('processSiteField', `Adding site field ${DisplayName}`)
        return await web.fields.createFieldAsXml(schemaXml)
      }
    } catch (error) {
      throw error
    }
  }
}
