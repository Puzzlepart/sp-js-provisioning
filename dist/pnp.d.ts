declare module "schema" {
    export interface Schema {
        Navigation?: INavigation;
        CustomActions?: ICustomAction[];
        ComposedLook?: IComposedLook;
        WebSettings?: IWebSettings;
        Features?: IFeature[];
        Lists?: IList[];
        Files?: IFile[];
        PropertyBagEntries?: IPropertyBagEntry[];
        [key: string]: any;
    }
    export default Schema;
    export interface IFeature {
        id: string;
        deactivate: boolean;
        force: boolean;
    }
    export interface IFile {
        Folder: string;
        Src: string;
        Url: string;
        Overwrite?: boolean;
        RemoveExistingWebParts?: boolean;
        WebParts?: IWebPart[];
        Properties?: {
            [key: string]: string | number;
        };
    }
    export interface IWebPartPropertyOverride {
        name: string;
        type: string;
        value: string;
    }
    export interface IWebPart {
        Title: string;
        Zone: string;
        Order: number;
        Contents: IWebPartContents;
        PropertyOverrides?: IWebPartPropertyOverride[];
        ListView?: {
            List: string;
            View: IListView;
        };
    }
    export interface IWebPartContents {
        Xml?: string;
        FileSrc?: string;
    }
    export interface IComposedLook {
        ColorPaletteUrl: string;
        FontSchemeUrl: string;
        BackgroundImageUrl: string;
    }
    export interface ICustomAction {
        Name: string;
        Description?: string;
        Title: string;
        Location: string;
        Url: string;
        [key: string]: string;
    }
    export interface IWebSettings {
        WelcomePage?: string;
        AlternateCssUrl?: string;
        SaveSiteAsTemplateEnabled?: boolean;
        MasterUrl?: string;
        CustomMasterUrl?: string;
        RecycleBinEnabled?: boolean;
        TreeViewEnabled?: boolean;
        QuickLaunchEnabled?: boolean;
        SiteLogoUrl?: string;
        [key: string]: string | boolean;
    }
    export interface INavigation {
        QuickLaunch?: INavigationNode[];
        TopNavigationBar?: INavigationNode[];
    }
    export interface INavigationNode {
        Title: string;
        Url: string;
        IgnoreExisting?: boolean;
        Children?: INavigationNode[];
    }
    export interface IRoleAssignment {
        Principal: string;
        RoleDefinition: string;
    }
    export interface IListSecurity {
        BreakRoleInheritance?: boolean;
        CopyRoleAssignments?: boolean;
        ClearSubscopes?: boolean;
        RoleAssignments?: IRoleAssignment[];
    }
    export interface IList {
        Title: string;
        Description: string;
        Template: number;
        ContentTypesEnabled: boolean;
        RemoveExistingContentTypes?: boolean;
        ContentTypeBindings?: IContentTypeBinding[];
        Fields?: string[];
        FieldRefs?: IListInstanceFieldRef[];
        Views?: IListView[];
        Security?: IListSecurity;
        AdditionalSettings?: {
            DefaultContentApprovalWorkflowId?: string;
            DefaultDisplayFormUrl?: string;
            DefaultEditFormUrl?: string;
            DefaultNewFormUrl?: string;
            Description?: string;
            Direction?: string;
            DocumentTemplateUrl?: string;
            /**
             * Reader = 0; Author = 1; Approver = 2.
             */
            DraftVersionVisibility?: number;
            EnableAttachments?: boolean;
            EnableFolderCreation?: boolean;
            EnableMinorVersions?: boolean;
            EnableModeration?: boolean;
            EnableVersioning?: boolean;
            ForceCheckout?: boolean;
            Hidden?: boolean;
            IrmEnabled?: boolean;
            IrmExpire?: boolean;
            IrmReject?: boolean;
            IsApplicationList?: boolean;
            NoCrawl?: boolean;
            OnQuickLaunch?: boolean;
            Title?: string;
            ValidationFormula?: string;
            ValidationMessage?: string;
            [key: string]: string | boolean | number;
        };
    }
    export interface IListInstanceFieldRef {
        ID: string;
        DisplayName?: string;
        Required?: boolean;
        Hidden?: boolean;
    }
    export interface IContentTypeBinding {
        ContentTypeID: string;
        Name?: string;
    }
    export interface IListView {
        Title: string;
        PersonalView?: boolean;
        ViewFields?: string[];
        AdditionalSettings?: {
            ViewQuery?: string;
            RowLimit?: number;
            Paged?: boolean;
            Hidden?: boolean;
            Scope?: 0 | 1;
        };
    }
    export interface IPropertyBagEntry {
        Key: string;
        Value: string;
        Indexed?: boolean;
        Overwrite?: boolean;
    }
}
declare module "handlers/handlerbase" {
    import { Web } from "sp-pnp-js";
    /**
     * Describes the Object Handler Base
     */
    export class HandlerBase {
        private name;
        /**
         * Creates a new instance of the ObjectHandlerBase class
         */
        constructor(name: string);
        /**
         * Provisioning objects
         */
        ProvisionObjects(web: Web, templatePart: any): Promise<void>;
        /**
         * Writes to Logger when scope has started
         */
        scope_started(): void;
        /**
         * Writes to Logger when scope has stopped
         */
        scope_ended(): void;
    }
}
declare module "util/index" {
    export function ReplaceTokens(str: string): string;
    export function MakeUrlRelative(absUrl: string): string;
    export function base64EncodeString(str: string): string;
    export function isNode(): boolean;
}
declare module "handlers/composedlook" {
    import { IComposedLook } from "schema";
    import { HandlerBase } from "handlers/handlerbase";
    import { Web } from "sp-pnp-js";
    /**
     * Describes the Composed Look Object Handler
     */
    export class ComposedLook extends HandlerBase {
        /**
         * Creates a new instance of the ObjectComposedLook class
         */
        constructor();
        /**
         * Provisioning Composed Look
         *
         * @param {Web} web The web
         * @param {IComposedLook} object The Composed Look to provision
         */
        ProvisionObjects(web: Web, composedLook: IComposedLook): Promise<void>;
    }
}
declare module "handlers/customactions" {
    import { HandlerBase } from "handlers/handlerbase";
    import { ICustomAction } from "schema";
    import { Web } from "sp-pnp-js";
    /**
     * Describes the Custom Actions Object Handler
     */
    export class CustomActions extends HandlerBase {
        /**
         * Creates a new instance of the ObjectCustomActions class
         */
        constructor();
        /**
         * Provisioning Custom Actions
         *
         * @param {Web} web The web
         * @param {Array<ICustomAction>} customactions The Custom Actions to provision
         */
        ProvisionObjects(web: Web, customActions: ICustomAction[]): Promise<void>;
    }
}
declare module "handlers/features" {
    import { HandlerBase } from "handlers/handlerbase";
    import { IFeature } from "schema";
    import { Web } from "sp-pnp-js";
    /**
     * Describes the Features Object Handler
     */
    export class Features extends HandlerBase {
        /**
         * Creates a new instance of the ObjectFeatures class
         */
        constructor();
        /**
         * Provisioning features
         *
         * @param {Web} web The web
         * @param {Array<IFeature>} features The features to provision
         */
        ProvisionObjects(web: Web, features: IFeature[]): Promise<void>;
    }
}
declare module "handlers/websettings" {
    import { HandlerBase } from "handlers/handlerbase";
    import { IWebSettings } from "schema";
    import { Web } from "sp-pnp-js";
    /**
     * Describes the WebSettings Object Handler
     */
    export class WebSettings extends HandlerBase {
        /**
         * Creates a new instance of the WebSettings class
         */
        constructor();
        /**
         * Provisioning WebSettings
         *
         * @param {Web} web The web
         * @param {IWebSettings} settings The settings
         */
        ProvisionObjects(web: Web, settings: IWebSettings): Promise<void>;
    }
}
declare module "handlers/navigation" {
    import { HandlerBase } from "handlers/handlerbase";
    import { INavigation } from "schema";
    import { Web } from "sp-pnp-js";
    /**
     * Describes the Navigation Object Handler
     */
    export class Navigation extends HandlerBase {
        /**
         * Creates a new instance of the Navigation class
         */
        constructor();
        /**
         * Provisioning navigation
         *
         * @param {Navigation} navigation The navigation to provision
         */
        ProvisionObjects(web: Web, navigation: INavigation): Promise<void>;
        private processNavTree(target, nodes);
        private processNode(target, node, existingNodes);
        private deleteExistingNodes(target);
        private deleteNode(target, id);
    }
}
declare module "handlers/lists" {
    import { HandlerBase } from "handlers/handlerbase";
    import { IList } from "schema";
    import { Web } from "sp-pnp-js";
    /**
     * Describes the Lists Object Handler
     */
    export class Lists extends HandlerBase {
        private lists;
        private tokenRegex;
        /**
         * Creates a new instance of the Lists class
         */
        constructor();
        /**
         * Provisioning lists
         *
         * @param {Web} web The web
         * @param {Array<IList>} lists The lists to provision
         */
        ProvisionObjects(web: Web, lists: IList[]): Promise<void>;
        /**
         * Processes a list
         *
         * @param {Web} web The web
         * @param {IList} lc The list
         */
        private processList(web, lc);
        /**
         * Processes content type bindings for a list
         *
         * @param {IList} lc The list configuration
         * @param {List} list The pnp list
         * @param {Array<IContentTypeBinding>} contentTypeBindings Content type bindings
         * @param {boolean} removeExisting Remove existing content type bindings
         */
        private processContentTypeBindings(lc, list, contentTypeBindings, removeExisting);
        /**
         * Processes a content type binding for a list
         *
         * @param {IList} lc The list configuration
         * @param {List} list The pnp list
         * @param {string} contentTypeID The Content Type ID
         */
        private processContentTypeBinding(lc, list, contentTypeID);
        /**
         * Processes fields for a list
         *
         * @param {Web} web The web
         * @param {IList} list The pnp list
         */
        private processFields(web, list);
        /**
         * Processes a field for a lit
         *
         * @param {Web} web The web
         * @param {IList} lc The list configuration
         * @param {string} fieldXml Field xml
         */
        private processField(web, lc, fieldXml);
        /**
       * Processes field refs for a list
       *
       * @param {Web} web The web
       * @param {IList} list The pnp list
       */
        private processFieldRefs(web, list);
        /**
         *
         * Processes a field ref for a list
         *
         * @param {Web} web The web
         * @param {IList} lc The list configuration
         * @param {IListInstanceFieldRef} fieldRef The list field ref
         */
        private processFieldRef(web, lc, fieldRef);
        /**
         * Processes views for a list
         *
         * @param web The web
         * @param lc The list configuration
         */
        private processViews(web, lc);
        /**
         * Processes a view for a list
         *
         * @param {Web} web The web
         * @param {IList} lc The list configuration
         * @param {IListView} lvc The view configuration
         */
        private processView(web, lc, lvc);
        /**
         * Processes view fields for a view
         *
         * @param {any} view The pnp view
         * @param {IListView} lvc The view configuration
         */
        private processViewFields(view, lvc);
        /**
         * Replaces tokens in field xml
         *
         * @param {string} fieldXml The field xml
         */
        private replaceFieldXmlTokens(fieldXml);
    }
}
declare module "handlers/files" {
    import { HandlerBase } from "handlers/handlerbase";
    import { IFile } from "schema";
    import { Web } from "sp-pnp-js";
    /**
     * Describes the Features Object Handler
     */
    export class Files extends HandlerBase {
        /**
         * Creates a new instance of the Files class
         */
        constructor();
        /**
         * Provisioning Files
         *
         * @param {Web} web The web
         * @param {IFile[]} files The files  to provision
         */
        ProvisionObjects(web: Web, files: IFile[]): Promise<void>;
        /**
         * Get blob for a file
         *
         * @param {IFile} file The file
         */
        private getFileBlob(file);
        /**
         * Procceses a file
         *
         * @param {Web} web The web
         * @param {IFile} file The file
         * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
         */
        private processFile(web, file, webServerRelativeUrl);
        /**
         * Remove exisiting webparts if specified
         *
         * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
         * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
         * @param {boolean} shouldRemove Should web parts be removed
         */
        private removeExistingWebParts(webServerRelativeUrl, fileServerRelativeUrl, shouldRemove);
        /**
         * Processes web parts
         *
         * @param {IFile} file The file
         * @param {string} webServerRelativeUrl ServerRelativeUrl for the web
         * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
         */
        private processWebParts(file, webServerRelativeUrl, fileServerRelativeUrl);
        /**
         * Fetches web part contents
         *
         * @param {IWebPart[]} webParts Web parts
         * @param {Function} cb Callback function that takes index of the the webpart and the retrieved XML
         */
        private fetchWebPartContents;
        /**
         * Processes page list views
         *
         * @param {Web} web The web
         * @param {IWebPart[]} webParts Web parts
         * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
         */
        private processPageListViews(web, webParts, fileServerRelativeUrl);
        /**
         * Processes page list view
         *
         * @param {Web} web The web
         * @param {any} listView List view
         * @param {string} fileServerRelativeUrl ServerRelativeUrl for the file
         */
        private processPageListView(web, listView, fileServerRelativeUrl);
        /**
         * Process list item properties for the file
         *
         * @param {Web} web The web
         * @param {File} pnpFile The PnP file
         * @param {Object} properties The properties to set
         */
        private processProperties(web, pnpFile, properties);
        /**
         * Replaces tokens in a string, e.g. {site}
         *
         * @param {string} str The string
         * @param {SP.ClientContext} ctx Client context
         */
        private replaceXmlTokens(str, ctx);
    }
}
declare module "handlers/propertybagentries" {
    import { HandlerBase } from "handlers/handlerbase";
    import { IPropertyBagEntry } from "schema";
    import { Web } from "sp-pnp-js";
    /**
     * Describes the PropertyBagEntries Object Handler
     */
    export class PropertyBagEntries extends HandlerBase {
        /**
         * Creates a new instance of the PropertyBagEntries class
         */
        constructor();
        /**
         * Provisioning property bag entries
         *
         * @param {Web} web The web
         * @param {Array<IPropertyBagEntry>} entries The property bag entries to provision
         */
        ProvisionObjects(web: Web, entries: IPropertyBagEntry[]): Promise<void>;
    }
}
declare module "handlers/exports" {
    import { TypedHash } from "sp-pnp-js";
    import { HandlerBase } from "handlers/handlerbase";
    export const DefaultHandlerMap: TypedHash<HandlerBase>;
    export const DefaultHandlerSort: TypedHash<number>;
}
declare module "webprovisioner" {
    import { Schema } from "schema";
    import { HandlerBase } from "handlers/handlerbase";
    import { TypedHash, Web } from "sp-pnp-js";
    /**
     * Root class of Provisioning
     */
    export class WebProvisioner {
        private web;
        handlerMap: TypedHash<HandlerBase>;
        handlerSort: TypedHash<number>;
        /**
         * Creates a new instance of the Provisioner class
         *
         * @param {Web} web The Web instance to which we want to apply templates
         * @param {TypedHash<HandlerBase>} handlermap A set of handlers we want to apply. The keys of the map need to match the property names in the template
         */
        constructor(web: Web, handlerMap?: TypedHash<HandlerBase>, handlerSort?: TypedHash<number>);
        /**
         * Applies the supplied template to the web used to create this Provisioner instance
         *
         * @param {Schema} template The template to apply
         * @param {Function} progressCallback Callback for progress updates
         */
        applyTemplate(template: Schema, progressCallback?: (msg: string) => void): Promise<void>;
    }
}
