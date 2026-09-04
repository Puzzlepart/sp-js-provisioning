import { IContentType } from './schema'

/**
 * Describes the Provisioning Context
 */
export class ProvisioningContext {
  public web = null
  /**
   * Absolute URL of the web being provisioned (derived from the pnp web instance).
   */
  public webUrl?: string
  /**
   * Id of the site collection the web belongs to (`{sitecollectionid}` token).
   */
  public siteId?: string
  /**
   * Id of the default site collection term store (`{sitecollectiontermstoreid}` token).
   */
  public termStoreId?: string
  public lists: { [key: string]: string } = {}
  public listViews: { [key: string]: string } = {}
  public siteFields: { [key: string]: string } = {}
  public contentTypes: { [key: string]: IContentType } = {}
}
