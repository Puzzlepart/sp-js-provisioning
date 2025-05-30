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
  [key: string]: any
}

export default Schema

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
  FontSchemeUrl: string
  BackgroundImageUrl: string
}

export interface ICustomAction {
  Name: string
  Description?: string
  Title: string
  Location: string
  Url: string

  [key: string]: string
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

export interface IRoleAssignment {
  Principal: string
  RoleDefinition: string
}

export interface IListSecurity {
  BreakRoleInheritance?: boolean
  CopyRoleAssignments?: boolean
  ClearSubscopes?: boolean
  RoleAssignments?: IRoleAssignment[]
}

export interface IListInstance {
  Title: string
  Description: string
  Template: number
  ContentTypesEnabled: boolean
  RemoveExistingContentTypes?: boolean
  ContentTypeBindings?: IContentTypeBinding[]
  Fields?: string[]
  FieldRefs?: IListInstanceFieldReference[]
  Views?: IListView[]
  RemoveExistingViews?: boolean
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
