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
  /**
   * Titles of the lists made in this setup: created by this provisioner, across its
   * `applyTemplate` calls (PP365 applies a template and then each extension through one). Only
   * these get `RemoveExistingViews` and `RemoveExistingContentTypes`.
   */
  public createdLists: { [title: string]: boolean } = {}
  /**
   * Titles of the lists `WebProvisioner` created before `SiteFields` in the current `applyTemplate`,
   * because a site column's `{listid:}` names them: `Lists` finds them existing and does not send
   * their settings again.
   */
  public preCreatedLists: { [title: string]: boolean } = {}
  public listViews: { [key: string]: string } = {}
  public siteFields: { [key: string]: string } = {}
  /**
   * Internal names of the site columns made in this setup (created by this provisioner, across its
   * `applyTemplate` calls): a target the template changes on one of them follows the template.
   */
  public createdSiteFields: { [internalName: string]: boolean } = {}
  public contentTypes: { [key: string]: IContentType } = {}
}
