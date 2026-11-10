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
  /**
   * Site columns as schema XML, matched by internal name (`Name`). A missing one is created; one the
   * web has gets the template merged onto its stored schema, written only when that changes
   * something, so what the template leaves out (an index, a formatter, a description) survives. A
   * widening type change is applied, another type change is a warning and leaves it as it is. A new
   * target (a lookup's list, a taxonomy column's term set or term store) is a warning that keeps the
   * stored one, unless the column was made in this setup. A column SharePoint refuses is logged as
   * an error and the others still run. Changes do not reach the column's existing list copies.
   */
  SiteFields?: string[]
  /**
   * Site content types, created with the id the template gives them, parents before their children
   * (in id order). One the web has is matched by id (or by `Name` when it has no `ID`) and gets
   * only what differs: a name, description or group, a field link it lacks, a changed `Required`
   * or `Hidden`, a changed order of the links. A change is pushed to the content types and lists
   * that inherit from it; an unchanged content type is not written. A content type SharePoint
   * refuses stops the template.
   */
  ContentTypes?: IContentType[]
  /**
   * A term group, its term sets and terms in the default site collection term store, with the ids
   * the template gives them. Missing ones are created; what the store has (by id, anywhere in it)
   * is left as it is, except the custom sort order, set when it differs, and the custom properties
   * a term lacks with `UpdateExistingTerms`. Runs first, before the site columns that use the term
   * sets. A term store that cannot be read or written stops the template.
   */
  Taxonomy?: ITaxonomy
  Security?: IWebSecurity
  [key: string]: any
}

export default Schema

export interface ITermGroup {
  /** The group's id (a GUID). A group the store has with this id is used as it is. */
  Id: string
  /** The name a new group gets. */
  Name: string
}

export interface ITerm {
  /** The term's id (a GUID). A term the store has with this id, in any term set, is kept. */
  Id: string
  /** The name a new term gets, in the term store's default language. */
  Name: string
  /**
   * The term's place in its term set's custom sort order: with two or more terms that have one,
   * the order is set when it differs from the term set's.
   */
  SortOrder?: number
  /** Set on a new term; on an existing one only those it lacks, with `UpdateExistingTerms`. */
  CustomProperties?: Record<string, string>
}

export interface ITermSet {
  /** The term set's id (a GUID). A term set the store has with this id is used as it is. */
  Id: string
  /** The name a new term set gets, in the term store's default language. */
  Name: string
  /** Set on a new term set. */
  Description?: string
  /** Set on a new term set. */
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
  /**
   * The site column's id. In a content type it finds the column when none has the internal name
   * in `Name`; a list's field ref needs it.
   */
  ID?: string
  /**
   * The site column's internal name. In a content type a column the web has under this name is
   * linked whatever `ID` says (its title also matches, last).
   */
  Name?: string
  /** Whether the column is required here. Left as it is when the template does not say. */
  Required?: boolean
  /** Whether the column is hidden here. Left as it is when the template does not say. */
  Hidden?: boolean
}

export interface IContentType {
  /**
   * The content type id, which names its parent: `0x0100…` is a child of Item, `<id>01` a child
   * of `<id>`. Without it, the content type the web has with this `Name` is updated, and a
   * missing one stops the template.
   */
  ID?: string
  Name: string
  /** Set when the template has one and it differs. */
  Description?: string
  /** Set when the template has one and it differs. */
  Group?: string
  /**
   * The site columns of the content type, as the web can use them (its own, and on a subweb its
   * parent webs'). A column the content type lacks is linked; one the web does not have is logged
   * as an error and left out, and the others are still linked. With two or more, the links are put
   * first in this order, before the ones the template does not name. A link the template does not
   * name is never removed.
   */
  FieldRefs?: IFieldReference[]
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

/**
 * The web's navigation. The template owns each collection it names: on every
 * run the collection ends with exactly the template's nodes, in its order. A
 * node whose title is already there at the same level is kept as it is (id,
 * URL, audiences); a missing one is added; a node the template does not
 * declare is deleted, children and nodes a site owner added included. A
 * collection the template leaves out is not touched. A node SharePoint refuses
 * (its URL names no existing file or folder) fails the run and leaves both
 * collections as they were.
 */
export interface INavigation {
  QuickLaunch?: INavigationNode[]
  TopNavigationBar?: INavigationNode[]
}

/**
 * A navigation node, matched to a stored node at the same level by its exact
 * title (case included).
 */
export interface INavigationNode {
  /**
   * Title, no tokens. SharePoint returns its own nodes' titles in the signed-in
   * account's language (`Hjem` or `Home`), so such a node matches only for an
   * account with the template's language; otherwise it is added and the
   * stored one deleted.
   */
  Title: string
  /**
   * Server-relative, web-relative (resolved against the web) or absolute.
   * Takes the URL tokens only (`{site}` is the site collection). Used when the
   * node is added, or with `IgnoreExisting`.
   */
  Url: string
  /**
   * Set the template's `Url` on a node that exists (where it differs). Without
   * it an existing node keeps its stored URL, at every level: SharePoint
   * updates a node's URL when its target is renamed, and an owner may have
   * changed it. A title stored more than once gets the template's URL either
   * way.
   */
  IgnoreExisting?: boolean
  /**
   * The node's children, owned like the collection: without `Children`, a
   * kept node loses the children it has.
   */
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
   * Break role inheritance before adding role assignments. Only done while it still inherits, so on
   * a re-run `CopyRoleAssignments` and `ClearSubscopes` do nothing.
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
  /** The site column's id: a list's field ref is found by it. */
  ID: string
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
    /**
     * The view's formatting, as the JSON text SharePoint stores (column or row formatting, a
     * card or tile layout). SharePoint ignores it when a view is created, so `Lists` sends the
     * settings again right after creating the view.
     */
    CustomFormatter?: string
    /**
     * The view's modern layout: `TILES` for a cards view, `COMPACTLIST`, `KANBAN` (board) or
     * `MODERNCALENDAR`. Like `CustomFormatter`, set again right after the view is created.
     */
    ViewType2?: 'KANBAN' | 'TILES' | 'COMPACTLIST' | 'MODERNCALENDAR'
  }
}

export interface IPropertyBagEntry {
  Key: string
  Value: string
  Indexed?: boolean
  Overwrite?: boolean
}
