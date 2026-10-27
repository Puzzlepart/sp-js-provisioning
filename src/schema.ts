export interface Schema {
  Parameters?: Record<string, string>
  Version?: string
  Hooks?: IHooks[]
  Navigation?: INavigation
  CustomActions?: ICustomAction[]
  ComposedLook?: IComposedLook
  WebSettings?: IWebSettings
  Features?: IFeature[]
  Lists?: IListInstance[]
  Files?: IFileObject[]
  PropertyBagEntries?: IPropertyBagEntry[]
  ClientSidePages?: IClientSidePage[]
  SiteFields?: string[]
  ContentTypes?: IContentType[]
  Taxonomy?: ITaxonomy
  Security?: IWebSecurity
  [key: string]: any
}

export default Schema

export interface ITermGroup {
  Id: string
  Name: string
}

export interface ITerm {
  Id: string
  Name: string
  SortOrder?: number
  CustomProperties?: Record<string, string>
}

export interface ITermSet {
  Id: string
  Name: string
  Description?: string
  IsOpenForTermCreation?: boolean
  Terms?: ITerm[]
  /**
   * When true, terms that already exist (matched by Id) get any **missing**
   * entries from {@link ITerm.CustomProperties} added. A property that already
   * exists on the term is left as-is even if its value differs from the
   * template (existing values are never overwritten); the term's name/labels
   * are also left untouched. Without this flag existing terms are not modified
   * at all. Use it to backfill custom properties (e.g. PhaseSubText /
   * PhaseDescription) onto a term set provisioned before those properties
   * existed.
   */
  UpdateExistingTerms?: boolean
}

export interface ITaxonomy {
  TermGroup: ITermGroup
  TermSets: ITermSet[]
}

export interface IFieldReference {
  ID: string
  Name?: string
  Required?: boolean
  Hidden?: boolean
}

export interface IContentType {
  ID: string
  Name: string
  Description: string
  Group: string
  FieldRefs: IFieldReference[]
}

export interface IClientSideControl {
  Id: string
  Properties: { [key: string]: any }
  ServerProcessedContent?: {
    htmlStrings: Record<string, string>
    searchablePlainTexts: Record<string, string>
    imageSources: Record<string, string>
    links: Record<string, string>
  }
  Text?: string
}

export interface IClientSidePageColumn {
  Factor: any
  Controls: IClientSideControl[]
}

export interface IClientSidePageSection {
  Columns: IClientSidePageColumn[]
}

export interface IClientSidePage {
  Name: string
  Title: string
  PageLayoutType: any
  CommentsDisabled?: boolean
  Sections?: IClientSidePageSection[]
  VerticalSection?: IClientSideControl[]
  Overwrite?: boolean
}

export interface IFeature {
  id: string
  deactivate: boolean
  force: boolean
}

export interface IFileObject {
  Folder: string
  Src: string
  Url: string
  Overwrite?: boolean
  RemoveExistingWebParts?: boolean
  WebParts?: IWebPart[]
  Properties?: { [key: string]: string | number }
}

export interface IWebPartPropertyOverride {
  name: string
  type: string
  value: string
}

export interface IWebPart {
  Title: string
  Zone: string
  Order: number
  Contents: IWebPartContents
  PropertyOverrides?: IWebPartPropertyOverride[]
  ListView?: {
    List: string
    View: IListView
  }
}

export interface IWebPartContents {
  Xml?: string
  FileSrc?: string
}

export interface IComposedLook {
  ColorPaletteUrl: string
  /** Without it, the theme has no font scheme of its own. */
  FontSchemeUrl?: string
  /** Without it, the theme has no background image (one set before is removed). */
  BackgroundImageUrl?: string
}

export interface ICustomAction {
  Name: string
  Description?: string
  Title: string
  Location: string
  Url?: string
  ClientSideComponentId?: string
  ClientSideComponentProperties?: string
  RegistrationId?: string
  RegistrationType?: number
  Sequence?: number

  [key: string]: any
}

export interface IWebSettings {
  WelcomePage?: string
  AlternateCssUrl?: string
  SaveSiteAsTemplateEnabled?: boolean
  MasterUrl?: string
  CustomMasterUrl?: string
  RecycleBinEnabled?: boolean
  TreeViewEnabled?: boolean
  QuickLaunchEnabled?: boolean
  SiteLogoUrl?: string

  [key: string]: string | boolean
}

export interface IHooks {
  Title?: string
  Url: string
  Method: string
  Headers?: Headers
  Body?: Body
}

export interface INavigation {
  QuickLaunch?: INavigationNode[]
  TopNavigationBar?: INavigationNode[]
}

export interface INavigationNode {
  Title: string
  Url: string
  IgnoreExisting?: boolean
  Children?: INavigationNode[]
}

/**
 * A role assignment on a list.
 */
export interface IRoleAssignment {
  /**
   * The principal: a PnP-style token for one of the web's associated groups
   * (`{associatedownergroupid}`, `{associatedmembergroupid}`,
   * `{associatedvisitorgroupid}`), a principal id, a site group name or a user
   * login name.
   */
  Principal: string | number
  /**
   * Role definition name as it exists on the web (localized, e.g. `Full kontroll`
   * or `Lese`). The English well-known names (`Full Control`, `Design`, `Edit`,
   * `Contribute`, `Read`, `View Only`) are resolved by role type as a fallback.
   */
  RoleDefinition: string
}

/**
 * List security, applied after the list's data rows. Mirrors the PnP
 * provisioning schema `Security` element.
 */
export interface IListSecurity {
  /**
   * Break role inheritance before adding role assignments.
   */
  BreakRoleInheritance?: boolean
  /**
   * Copy the inherited role assignments when breaking inheritance.
   */
  CopyRoleAssignments?: boolean
  /**
   * Clear unique permissions on child objects when breaking inheritance.
   */
  ClearSubscopes?: boolean
  /**
   * Role assignments to add. Requires `BreakRoleInheritance`, since SharePoint
   * rejects role assignments on a list that still inherits permissions;
   * inherited assignments are kept only with `CopyRoleAssignments`.
   */
  RoleAssignments?: IRoleAssignment[]
}

/**
 * A SharePoint site group. Mirrors the PnP provisioning schema `SiteGroup`
 * element. Groups are matched by title, so re-applying a template updates the
 * group instead of creating a duplicate.
 */
export interface ISiteGroup {
  /**
   * Group title, unique within the site collection. Supports tokens, e.g.
   * `{sitename} Styringsgruppe`.
   */
  Title: string
  Description?: string
  /**
   * Group owner, resolved like a role assignment principal: an associated
   * group token, a site group title (including this group's own title) or a
   * user login name. Without it SharePoint keeps the provisioning user as owner.
   */
  Owner?: string
  AllowMembersEditMembership?: boolean
  AllowRequestToJoinLeave?: boolean
  AutoAcceptRequestToJoinLeave?: boolean
  OnlyAllowMembersViewMembership?: boolean
  RequestToJoinLeaveEmailSetting?: string
  /**
   * Members to add, as login names (UPN, email or claims, e.g. an Entra ID
   * group claim). Each one is ensured on the web before it is added.
   */
  Members?: string[]
  /**
   * Remove existing members that are not in `Members`.
   */
  ClearExistingMembers?: boolean
}

/**
 * Web security: site groups, and role inheritance and role assignments on the
 * web itself. Mirrors the PnP provisioning schema `Security` element. The
 * members inherited from {@link IListSecurity} behave as they do for lists, but
 * apply to the web; the root web of a site collection always has unique
 * permissions, so it takes role assignments without `BreakRoleInheritance`.
 */
export interface IWebSecurity extends IListSecurity {
  /**
   * Site groups to create or update. They are provisioned before `Lists`, so
   * list `Security` can grant permissions to them.
   */
  SiteGroups?: ISiteGroup[]
}

export interface IListInstance {
  Title: string
  Description: string
  Template: number
  ContentTypesEnabled: boolean
  /**
   * Remove the content types not bound, on a list made in this setup (created by this provisioner,
   * across its `applyTemplate` calls). On a list that existed it does nothing.
   */
  RemoveExistingContentTypes?: boolean
  ContentTypeBindings?: IContentTypeBinding[]
  /**
   * List fields as schema XML. A missing one is created; one that exists is updated in place, never
   * deleted: a widening type change (Text to Note, Lookup or User single to multi) is applied, another type change
   * is a warning and leaves it as it is. A new target (a lookup's list, a taxonomy field's term
   * set) is applied on a list made in this setup, otherwise a warning that keeps the stored one.
   */
  Fields?: string[]
  FieldRefs?: IListInstanceFieldReference[]
  Views?: IListView[]
  /**
   * Remove the views the template does not name, on a list made in this setup (created by this
   * provisioner, across its `applyTemplate` calls), after the template's own exist. Never a
   * personal view, the default view or the view behind a page's list web part. On a list that
   * existed it does nothing (PnP Framework removes on every run).
   */
  RemoveExistingViews?: boolean
  DataRows?: IDataRows
  Folders?: IFolder[]
  Security?: IListSecurity

  AdditionalSettings?: {
    DefaultContentApprovalWorkflowId?: string
    DefaultDisplayFormUrl?: string
    DefaultEditFormUrl?: string
    DefaultNewFormUrl?: string
    Description?: string
    Direction?: string
    DocumentTemplateUrl?: string
    /**
     * Reader = 0; Author = 1; Approver = 2.
     */
    DraftVersionVisibility?: number
    EnableAttachments?: boolean
    EnableFolderCreation?: boolean
    EnableMinorVersions?: boolean
    EnableModeration?: boolean
    EnableVersioning?: boolean
    ForceCheckout?: boolean
    Hidden?: boolean
    IrmEnabled?: boolean
    IrmExpire?: boolean
    IrmReject?: boolean
    IsApplicationList?: boolean
    NoCrawl?: boolean
    OnQuickLaunch?: boolean
    Title?: string
    ValidationFormula?: string
    ValidationMessage?: string

    [key: string]: string | boolean | number
  }
}

/**
 * Seed data for a list (provisioned after fields/content types/views exist).
 *
 * Rows are upserted: when {@link KeyColumn} is set, an existing item whose
 * KeyColumn matches is updated (or skipped, per {@link UpdateBehavior}) instead
 * of creating a duplicate — so re-running the template is idempotent. Without a
 * KeyColumn every row is added.
 */
export interface IDataRows {
  /** Field internal name used to match existing rows for idempotent upsert. */
  KeyColumn?: string
  /** When a row matches {@link KeyColumn}: `Overwrite` (default) or `Skip`. */
  UpdateBehavior?: 'Overwrite' | 'Skip'
  Rows: IDataRow[]
}

/**
 * One row: a map of field **internal name** → value. Value shapes by field type:
 * - Text / Note / Number / Currency / Choice / Boolean / DateTime: the raw value.
 * - MultiChoice: `string[]`.
 * - URL: a string, or `{ Url, Description? }`.
 * - Lookup / LookupMulti: an item id (number), `{ lookupId }`, or `{ lookupValue }`
 *   (resolved against the target list's show field). Multi accepts an array.
 * - User / UserMulti: a login/email string, `{ login }`, or a user id. Multi accepts an array.
 * - Taxonomy (`TaxonomyFieldType` / `…Multi`): `{ termId, label }` (label required to
 *   write the hidden note field). Multi accepts an array of `{ termId, label }`.
 */
export interface IDataRow {
  [fieldInternalName: string]: any
}

/**
 * A folder to provision inside a list / document library. Provisioned after
 * fields, content types and views exist. Creation is idempotent: an existing
 * folder with the same name is left in place and recursion continues into its
 * children, so re-running the template is safe.
 */
export interface IFolder {
  /**
   * Leaf name of the folder (not a path). May contain spaces and Norwegian
   * characters (å, æ, ø) — written as-is; PnPjs handles URL encoding.
   */
  Name: string
  /** Optional nested subfolders, created beneath this folder. */
  Folders?: IFolder[]
}

export interface IListInstanceFieldReference extends IFieldReference {
  DisplayName?: string
  AdditionalProperties?: Record<string, string>
}

export interface IContentTypeBinding {
  ContentTypeID: string
  Name?: string
}

export interface IListView {
  Title: string
  PersonalView?: boolean
  ViewFields?: string[]
  AdditionalSettings?: {
    ViewQuery?: string
    RowLimit?: number
    Paged?: boolean
    Hidden?: boolean
    Scope?: 0 | 1
    DefaultView?: boolean
    JSLink?: string
  }
}

export interface IPropertyBagEntry {
  Key: string
  Value: string
  Indexed?: boolean
  Overwrite?: boolean
}
