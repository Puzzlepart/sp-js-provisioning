# Architecture

This document is a field guide for agents working in this repository. It explains the current architecture, the provisioning flow, the shared contracts between modules, and the places where small changes can have larger effects.

## Purpose

`sp-js-provisioning` is a TypeScript library for applying JSON provisioning templates to SharePoint webs. A template is split into top-level sections such as `SiteFields`, `ContentTypes`, `Lists`, `Files`, `Navigation`, and `Hooks`. Each section is handled by a dedicated object handler.

The package exports three public symbols from `src/index.ts`:

- `Schema` - TypeScript shape for provisioning templates.
- `WebProvisioner` - runtime orchestrator that applies templates to a SharePoint web.
- `ProvisioningError` - wrapper error that records which handler failed.

The compiled package output is written to `lib/`. Source lives in `src/`.

## Runtime Model

The normal runtime path is:

1. A consumer creates a PnP `IWeb` instance for the target SharePoint web.
2. The consumer creates `new WebProvisioner(web)`.
3. The consumer calls `.setup(config)`.
4. The consumer calls `.applyTemplate(template, handlers?, progressCallback?)`.
5. `WebProvisioner` initializes logging and shared context.
6. `WebProvisioner` sorts template sections by handler order.
7. Each matching handler provisions its section sequentially.
8. Later handlers use IDs and state collected by earlier handlers through `ProvisioningContext`.

The orchestration is intentionally sequential. Many handlers depend on SharePoint artifacts created by earlier handlers, and several handlers mutate the shared context.

## Entry Points

### `src/index.ts`

Public package barrel. Keep this narrow. Any new public API should be intentionally exported here.

### `src/webprovisioner.ts`

Main orchestrator.

Important responsibilities:

- Holds the target `IWeb`.
- Stores `IProvisioningConfig`.
- Creates `DefaultHandlerMap(config)` during setup.
- Loads `context.web` from a fresh copy of the web (`Web(this.web)`), and hands every handler a fresh copy too. PnPjs 4's `select()` sets the query on the queryable itself, so a `$select` a handler left on a shared web made the next `applyTemplate` load only that property (`{sitename}` stayed literal).
- Sorts top-level template keys with `DefaultHandlerSort`.
- Filters to requested handlers if `handlers` is supplied.
- Creates, right before `SiteFields` runs, the lists that an unresolved `{listid:}` in a site column names and that the template's `Lists` declares (`createListsForSiteFields`, through `Lists.createList`; only when `SiteFields` and `Lists` both run, `SiteFields` first, and `loadContext` read the web's lists, else a warning). Records them in `context.preCreatedLists` (reset per `applyTemplate`); a list it cannot create is logged as an error and the run goes on; a title declared twice is created once. When the run fails before `Lists` reaches such a list, a warning names it.
- Executes handlers sequentially.
- Wraps failures in `ProvisioningError(currentHandler, error)`.

Important behavior:

- Unknown template sections are ignored unless a matching handler exists in `handlerMap`.
- Handler selection is by top-level schema property name.
- `progressCallback` is called before each handler starts.
- The current implementation logs before `onSetup()`, so `this.config` must be set by `.setup(config)` before `applyTemplate()`.

### `src/handlers/exports.ts`

Defines the handler registry and dependency order.

Current order:

| Sort | Handler | Why It Is Early Or Late |
| ---: | --- | --- |
| 0 | `SiteFields` | Creates or updates web fields used by content types and lists. |
| 1 | `ContentTypes` | Uses site fields and produces content types used by list bindings. |
| 2 | `Features` | Activates or deactivates web features before list and page work. |
| 2.5 | `Security` | Creates site groups and applies web permissions; runs right before `Lists` so list `Security` can grant permissions to those groups. |
| 3 | `Lists` | Creates lists, binds content types, adds list fields, field refs, and views. |
| 4 | `Files` | Uploads files and classic pages, can add web parts and list view web parts. |
| 5 | `CustomActions` | Adds user custom actions after base structures exist. |
| 6 | `ComposedLook` | Applies theme assets, usually after files are available. |
| 7 | `ClientSidePages` | Creates modern pages and page controls. |
| 8 | `PropertyBagEntries` | Sets web property bag values after core structures exist. |
| 9 | `Navigation` | Reconciles navigation after pages and lists exist. |
| 10 | `WebSettings` | Applies final web/root folder settings. |
| 11 | `Hooks` | Calls external endpoints after provisioning has completed. |

Exception: right before `SiteFields` runs, `WebProvisioner` creates the lists of the template that a site column's `{listid:}` names (`createListsForSiteFields`, through `Lists.createList`), when `SiteFields` and `Lists` both run and `SiteFields` comes first. A list's id cannot be chosen, so the list must exist before the column's XML can name it (`docs/specs/sitefields-re-apply.md`).

When adding a handler, update:

- `Handler` union type
- `DefaultHandlerMap`
- `DefaultHandlerSort`
- `Schema` if the handler is template-driven

## Core Contracts

### `Schema`

Defined in `src/schema.ts`.

`Schema` is the top-level provisioning template contract. Each property maps directly to a handler name where possible:

- `SiteFields?: string[]`
- `ContentTypes?: IContentType[]`
- `Lists?: IListInstance[]`
- `Files?: IFileObject[]`
- `ClientSidePages?: IClientSidePage[]`
- `CustomActions?: ICustomAction[]`
- `Features?: IFeature[]`
- `Security?: IWebSecurity`
- `ComposedLook?: IComposedLook`
- `PropertyBagEntries?: IPropertyBagEntry[]`
- `Navigation?: INavigation`
- `WebSettings?: IWebSettings`
- `Hooks?: IHooks[]`
- `Parameters?: Record<string, string>`
- `Version?: string`

The interface also has `[key: string]: any`, so unknown sections are type-accepted. Runtime only processes sections with matching handlers.

### `IProvisioningConfig`

Defined in `src/provisioningconfig.ts`.

Fields:

- `parameters?: Record<string, string>` - values for `{parameter:...}` token replacement.
- `spfxContext?: any` - SPFx page context used for URL token replacement.
- `logging?: { activeLogLevel?: LogLevel; prefix?: string }` - PnP logging configuration.

### `ProvisioningContext`

Defined in `src/provisioningcontext.ts`.

Shared mutable context passed to every handler:

- `web` - loaded web metadata from `await this.web()`.
- `webUrl` - absolute URL of the web (derived from the pnp web instance).
- `siteId` - site collection ID, backing the `{sitecollectionid}` token.
- `termStoreId` - default site collection term store ID, backing the `{sitecollectiontermstoreid}` token.
- `lists: { [title: string]: string }` - list title to list ID. Pre-populated in `WebProvisioner.onSetup` so `{listid:...}` resolves in handlers that run before `Lists`; the `Lists` handler reloads it, and `WebProvisioner` adds the lists it creates before `SiteFields`.
- `createdLists: { [title: string]: boolean }` - lists made in this setup (created by this provisioner, across its `applyTemplate` calls); only these get `RemoveExistingViews` and `RemoveExistingContentTypes`.
- `preCreatedLists: { [title: string]: boolean }` - lists `WebProvisioner` created before `SiteFields` in the current `applyTemplate`; `Lists` does not send their settings again.
- `listViews: { [listAndView: string]: string }` - key format is `${listTitle}|${viewTitle}`.
- `siteFields: { [internalName: string]: string }` - site field internal name to field ID.
- `createdSiteFields: { [internalName: string]: boolean }` - site columns made in this setup (by this provisioner, across its `applyTemplate` calls); a target the template changes on one of them follows the template.
- `contentTypes: { [nameOrId: string]: IContentType }` - content type entries keyed by both name and ID.

Context is the main dependency bridge between handlers. Be careful when changing key formats or timing. If context loading is async, it must be awaited before dependent decisions are made.

## Token System

Token replacement is split between two helpers.

### `TokenHelper`

Defined in `src/util/tokenhelper.ts`.

Supports context/config tokens with the patterns `{token}` and `{token:value}`:

- `{listid:List Title}` -> `context.lists[List Title]`
- `{listviewid:List Title|View Title}` -> `context.listViews[...]`
- `{webid}` / `{siteid}` -> `context.web.Id` (PnP semantics: `{siteid}` is the web ID)
- `{sitename}` -> `context.web.Title` (PnP semantics: the title of the web)
- `{sitecollectionid}` -> `context.siteId` (falls back to `context.web.Id`)
- `{sitecollectiontermstoreid}` / `{termstoreid}` -> `context.termStoreId`
- `{parameter:Name}` -> `config.parameters.Name`

The IDs are loaded once in `WebProvisioner.onSetup` (`loadContext`); each lookup fails soft with a logged warning, leaving the token unresolved.

`unresolvedTokens(text)` returns the known tokens a string holds that do not resolve now, each once, as written; `SiteFields` logs them as an error naming the column, and `WebProvisioner` uses the `{listid:}` ones to create lists before `SiteFields`.

Important limitations:

- `{listviewid:}` does not resolve before `Lists`: `loadContext` loads no views, so it resolves there only for views an earlier `applyTemplate` through the same provisioner processed.
- The regex only matches lowercase token keys (`[a-z]+`); the value part may contain anything but braces. Unknown keys and JSON-like text such as `{"a":1}` or `{Title}` are left untouched.
- It replaces a match only when the token resolves; an unresolved token is left in place. A resolved empty value does replace the token, so an empty `{parameter:Name}` yields an empty string.
- Where tokens are applied: field XML (site and list fields), list `DataRows` values (every string value whose field type is not handled specially, the `KeyColumn` value when matching, and a URL's `Url` and `Description`), file sources and web part XML, client-side page properties and server-processed content, site group titles, string properties, owners and members, and role assignment principals (web and list `Security`). Not view queries, not custom action properties and not navigation nodes (their `Url` gets the URL tokens only, `replaceUrlTokens`; `Title` none).

### URL Tokens

Defined in `src/util/index.ts` as `replaceUrlTokens`.

Supports:

- `{site}` -> site server-relative URL
- `{sitecollection}` -> site absolute URL
- `{wpgallery}` -> site absolute URL plus `/_catalogs/wp`
- `{hosturl}` -> the origin (`https://contoso.sharepoint.com`, no trailing colon): from the SPFx context's web, else `window.location`; left in place when there is neither
- `{themegallery}` -> site absolute URL plus `/_catalogs/theme/15`

URL token replacement uses `config.spfxContext` when available. Otherwise it falls back to `window._spPageContextInfo`; without either (Node with no SPFx context) the site tokens become `null`. `window` and `document` are reached through `browserWindow()`, so the URL tokens, `makeUrlRelative` and the property bag key encoding also work in Node.

## Handler Architecture

All handlers extend `HandlerBase`.

`HandlerBase` provides:

- common constructor with handler name and config
- default `ProvisionObjects` warning implementation
- `scope_started()` and `scope_ended()`
- `log_info`, `log_warn`, and `log_error`

Handler convention:

- Public method is named `ProvisionObjects`.
- Inputs are the PnP `IWeb`, the template section, and sometimes `ProvisioningContext`.
- Most handlers process items sequentially with promise `reduce`.
- Handler errors should be thrown so `WebProvisioner` can wrap them in `ProvisioningError`. SiteFields, Lists and Security log a failing item with `log_error` and carry on instead; only failures that leave nothing to do are thrown.

### `SiteFields`

File: `src/handlers/sitefields.ts`; the merge and comparison rules in `src/util/fieldxml.ts`, shared with `Lists`

Responsibilities:

- Reads the web's own columns once (`Id`, `InternalName`, `TypeAsString`, `SchemaXml`) and the columns available to it (`web.availablefields`: on a subweb also its parents'), through fresh queryables and without `top()` (the collection is not paged; `$top` truncates it). Fills `context.siteFields`.
- Replaces context/config tokens in each field XML; a known token that does not resolve is an error naming the column and the token, and the column is still created with it as written.
- Creates a missing column with `web.fields.createFieldAsXml` and records it in `context.createdSiteFields`.
- Updates an existing column in place with the list-field rules: the template merged onto the stored `SchemaXml` (`mergeFieldXml`), written only when `sameFieldXml` says it changed (`update({ SchemaXml }, 'SP.Field')`); a widening type change is applied, another is a warning and the column is left as it is; a stored GUID target (a lookup's `List`/`WebId`, a taxonomy column's `SspId`/`TermSetId`) the template changes is a warning and kept, unless the column was made in this setup; a stored non-GUID target (an old literal token), or the empty GUID as `SspId`/`TermSetId`, follows the template; a template target still holding an unresolved token never overwrites the stored one.
- Taxonomy columns: the merge keeps the stored `Customization` and `List`, `WebId`, `ShowField` (SharePoint manages them and ignores the `Customization` in a SchemaXml MERGE); an `SspId` or `TermSetId` that follows the template is set with one `update({ SspId, TermSetId }, 'SP.Taxonomy.TaxonomyField')` (only the ones that follow).
- Never pushes to list copies: a SchemaXml MERGE does not reach them, and neither does a content type update.

Key contracts:

- Site column identity is keyed by internal name. Same name with another id: a warning, the stored id stays. Same id under another name (on the web or a parent web): an error, not created. A column only a parent web has (subweb): a warning, neither created nor changed.
- Each column is processed in its own try/catch: a failure is logged with `log_error` naming the column and the operation (reading its XML, create, update), and the handler resolves. When the web's columns cannot be read, one error names the template's columns and nothing is created or changed.
- Field XML must have a top-level field element with `Name` and usually `DisplayName`.

### `ContentTypes`

File: `src/handlers/contenttypes.ts`

Responsibilities:

- Initializes JSOM with `initSpfxJsom(context.web.ServerRelativeUrl)`.
- Loads existing web content types and field links into `context.contentTypes`.
- Resolves configured content type IDs by explicit `ID` or existing name.
- Creates missing content types.
- Updates name, description, and group.
- Adds or updates field links using JSOM.
- Reorders field links to match template order.

Key contracts:

- Content types are keyed in context by both `Name` and `ID`.
- Field link existence is determined from `context.contentTypes`.
- Content type field refs depend on site fields already existing.
- List content type bindings depend on content types already existing.

Related design notes:

- Deterministic content type ID behavior is specified in `docs/specs/deterministic-content-type-ids.md`.
- The JSOM explicit-ID decision is tracked in `docs/adr/jsom-explicit-content-type-ids.md`.
- Context initialization must be awaited before create-versus-update decisions depend on `context.contentTypes`.

### `Security`

File: `src/handlers/security.ts`

Responsibilities:

- Creates site groups (`SiteGroups`), or updates the configured properties of a group that already exists with the same title.
- Adds group members (each login name is ensured on the web first) and, with `ClearExistingMembers`, removes members that are not in the template.
- Sets the group owner after the members are added, so the provisioning user still owns a new group while it is being filled. A group owner goes through CSOM (`Group.Owner` over `_vti_bin/client.svc/ProcessQuery`, `src/util/csom.ts`), a user owner through REST's `SetUserAsOwner`, which answers success for a group id and changes nothing. The owner is read back; a mismatch is logged as an error.
- Breaks role inheritance on the web when `BreakRoleInheritance` is set and the web still inherits.
- Adds role assignments on the web.

Key contracts:

- Groups are matched by title through `siteGroups.getByName` (apostrophes escaped); a 404 means the group is missing and is created.
- Group owners and role assignment principals are resolved by `resolvePrincipalId` in `src/util/security.ts`, shared with list `Security`: associated group tokens, a principal ID, a site group title or a user login name. Role definitions use `resolveRoleDefinitionId` from the same module.
- `ClearExistingMembers` is skipped (with a warning) when any configured member could not be added, so an unresolved login cannot cause an existing member to be removed.

Important behavior:

- Failures are logged per group, member, owner and role assignment and never abort the template.
- The handler does not remove groups or role assignments that are missing from the template.
- Consumers with a custom `handlerSort` must sort `Security` before `Lists`; an unknown key sorts as 99 and would run after `Lists`.

### `Lists`

File: `src/handlers/lists.ts`

Responsibilities:

- Initializes JSOM for list operations that need it.
- Loads existing lists into `context.lists`.
- Creates lists (recording them in `context.createdLists`, the lists made in this setup) or updates an existing list with only the settings the template defines.
- Adds content type bindings.
- On a list made in this setup, optionally removes the content types not bound (`RemoveExistingContentTypes`).
- Creates list fields from field XML, or updates them in place (`src/util/fieldxml.ts`): never deletes one.
- Adds or updates list field refs from web fields.
- Sets the template's views and their fields when they differ; on a list made in this setup, optionally removes the other views (`RemoveExistingViews`).
- Creates folder hierarchies (`Folders`).
- Seeds items from `DataRows` (upsert by the token-resolved `KeyColumn`; string values get token replacement).
- Applies list `Security` (breaks role inheritance only on a list that still inherits, adds role assignments).
- Updates `context.listViews`.

Processing phases inside `ProvisionObjects`:

1. Load existing lists.
2. Process each list definition.
3. Process list fields for all lists.
4. Process list field refs for all lists.
5. Process views for all lists.
6. Process folders for all lists.
7. Process data rows for all lists.
8. Process security for all lists.
9. Reload list IDs into context.

Key contracts:

- List identity is title-based in context.
- Content type binding uses `ContentTypeID`; a list content type belongs to a binding when its id is the binding's or a child of it.
- `RemoveExistingContentTypes` keeps folder content types (`0x0120`).
- "Made in this setup" is `context.createdLists`, which lives on the `WebProvisioner` instance across its `applyTemplate` calls.
- List field refs match fields by field ID, not name.
- `SourceID` for copied web fields is set to the list ID from context.

Important behavior:

- `processField` never deletes a list field (until 1.5.0 it deleted and recreated every one on every run, losing its values; `docs/specs/lists-re-apply.md`). An existing field gets the template merged onto its stored schema XML (`mergeFieldXml`: template attributes win, the type as SharePoint stores it (`Type="Lookup" Mult="TRUE"` is `LookupMulti`), identity and storage attributes stay, `Version` dropped, child elements replaced unless `keep` names them, `Customization` merged by property), sent only when it differs (`sameFieldXml`). A widening type change is applied in place; another one is a warning, and so is a new target (a lookup's list, a taxonomy field's term set) on a list not made in this setup, because SharePoint accepts them and loses or misdirects the stored values (probe 2026-10-26). Primary lookups are made before their dependents.
- `processFieldRef` logs an error when a FieldRef names a column found neither on the list nor on the web (`web.fields`: on a subweb, a parent web's column is not found).
- `createList` is the create path, shared with `WebProvisioner`'s creation of lists before `SiteFields`; a list created there (`context.preCreatedLists`) is not sent its settings again.
- `DataRows` decides add, update or skip from the key before building a row: a skipped row queries no lookup list and ensures no user. An unresolved lookup or user value is an error naming it; an existing item keeps its stored value.
- `Security` principals accept the PnP tokens `{associatedownergroupid}`, `{associatedmembergroupid}` and `{associatedvisitorgroupid}`, a principal ID, a site group name or a user login name; other tokens in a principal are replaced first. Role definitions are resolved by localized name with a fallback to the well-known role type for the English names. The resolution lives in `src/util/security.ts`, shared with the `Security` handler. Failures are logged per assignment and never abort the template.
- Several list operations catch errors and log failures without throwing. Agents should inspect logs and behavior, not just promise success.

### `Files`

File: `src/handlers/files.ts`

Responsibilities:

- Uploads files from configured source URLs.
- Replaces URL tokens and context tokens in file sources.
- Fetches file contents with browser `fetch`.
- Adds files with `addUsingPath`.
- Updates list item properties for uploaded files.
- Removes and adds classic web parts through JSOM.
- Loads web part XML from inline contents or `FileSrc`.
- Applies web part property overrides.
- Configures page list view web parts.

Runtime limitations:

- Throws when `config.spfxContext` is present.
- Uses browser APIs such as `fetch`, `Blob`, `document`, `window`, and global `SP`.
- Classic web part operations depend on JSOM.

### `ClientSidePages`

File: `src/handlers/clientsidepages.ts`

Responsibilities:

- Loads available client-side web part definitions.
- Creates or overwrites modern pages.
- Adds vertical section controls.
- Adds section, column, text, and web part controls.
- Replaces config/context tokens in web part properties and server-processed content, in each string value (and key) with `mapStrings` (`src/util/index.ts`), not in their JSON text: a token value with a quote or a backslash used to break `JSON.parse` and lose the part.
- Saves the page.
- Enables or disables comments.

Key behavior:

- `Id` values `Text` and `PageText` create `ClientsideText`.
- Other controls are matched by checking whether the component definition ID contains the configured ID, case-insensitively.
- If `Overwrite` is false and the page exists, processing is skipped.

### `CustomActions`

File: `src/handlers/customactions.ts`

Responsibilities:

- Reads existing user custom actions by title.
- Adds missing custom actions in a PnP batch.

Key behavior:

- Existing actions are detected by `Title`, not `Name`.
- Existing actions are not updated.
- The action object is passed through to `web.userCustomActions.add(...)` as-is, so SPFx extension actions are supported by setting `Location` (`ClientSideExtension.ApplicationCustomizer` or `ClientSideExtension.ListViewCommandSet.CommandBar`), `ClientSideComponentId` and `ClientSideComponentProperties` (JSON string). Command sets can be scoped with `RegistrationId`/`RegistrationType` (numeric REST enum, e.g. `1` = List). `Url` is only relevant for classic actions and is optional.
- Actions are web-scoped (`web.userCustomActions`); site-collection-scoped actions are not supported.
- No token replacement is applied to action properties.

### `Features`

File: `src/handlers/features.ts`

Responsibilities:

- Activates or deactivates web features.
- Processes features sequentially.
- Reads the active web features once and leaves a feature that is already in the requested state alone (logged at Info): SharePoint refuses to activate an active feature (500) and to deactivate an inactive one (400). `force` still re-activates an active feature; a feature listed twice follows each entry in order.

Template contract:

- `id`
- `deactivate`
- `force`

### `ComposedLook`

File: `src/handlers/composedlook.ts`

Responsibilities:

- Applies SharePoint theme URLs using `web.applyTheme`.
- Replaces URL tokens.
- Converts absolute URLs to relative URLs before applying theme assets.
- `FontSchemeUrl` and `BackgroundImageUrl` are optional; a missing one is sent as `null` (no font scheme of its own, no background image).

### `PropertyBagEntries`

File: `src/handlers/propertybagentries.ts`

Responsibilities:

- Sets web property bag entries.
- Optionally indexes property bag keys by updating `vti_indexedpropertykeys`.

Runtime limitations:

- Not supported in Node.
- Not supported in SPFx mode.
- Uses browser/global JSOM `SP.ClientContext`.

Important behavior:

- Only entries with `Overwrite` truthy are processed.
- Indexed property keys are base64 encoded as UTF-16LE-style strings by `base64EncodeString`.

### `Navigation`

File: `src/handlers/navigation.ts`; URL comparison in `src/util/navigationurl.ts`

Responsibilities:

- Reconciles the Quick Launch and the top navigation bar with the template, in place: the template owns each collection it names (replace mode); a collection it leaves out is not touched.
- Replaces URL tokens in `Url` (`replaceUrlTokens`); `Title` gets no tokens.

Important behavior:

- Each collection is read once, through a fresh queryable, with its children expanded one level below the template's depth (`select('Id', 'Title', 'Url', 'Children/Id', ...)`, `expand('Children', ...)`). The web (`Url`, `ServerRelativeUrl`) is read once, only when a URL has to be compared.
- At each level a template node is matched to the first unmatched stored node with exactly the same title. A matched node is kept (id, audiences, URL) and nothing is written for it; its children are reconciled against the template node's `Children` (none: its children are deleted). A missing node is added with its children; a stored node without a match is deleted with its subtree. Then `MoveAfter` puts each node after the one before it in the template.
- URL rule, at every level: a kept node keeps its stored URL unless `IgnoreExisting` is true; then the template's URL is set (only `Url` is sent) where it differs once both are normalised (`normaliseNavigationUrl`: web-relative and same-host absolute to server-relative, `%20`, a trailing slash, the case of the path; external URLs verbatim). A title stored more than once: the first is kept with the template's URL, the others deleted. The template object is not changed.
- Write order: the Quick Launch's adds and URL updates, then the top navigation bar's; then each collection's deletes and moves. The adds and URL updates are journaled; when one fails (SharePoint refuses a URL that names no file or folder: `-2130247147`), or a collection cannot be read, the journal is undone, the last write first (added nodes deleted, URLs set back), the touched collections are read back (SharePoint reports success for writes to missing nodes), anything left is logged as an error naming the node, and the handler throws `Navigation node '<Title>' (<URL>) was refused by SharePoint: <message>` (or `Failed to read <collection>: ...`). A failed delete or move is logged and thrown; there is no rollback once the deletes start. A level where nodes moved is read back and must be in the template's order.
- An unchanged re-apply sends no write.

### `WebSettings`

File: `src/handlers/websettings.ts`

Responsibilities:

- Replaces URL tokens in string settings.
- Updates web properties.
- Updates root folder `WelcomePage` separately when supplied.

Important behavior:

- `web.update(settings)` runs whether or not `WelcomePage` is set (up to 1.4.0 it ran only together with a welcome page, so settings without one were ignored). Nothing is sent when there is nothing to set.
- The welcome page is split off with object rest (not `object.omit`, which failed as a namespace import in Node).

### `Hooks`

File: `src/handlers/hooks.ts`

Responsibilities:

- Runs external HTTP hooks after provisioning.
- Supports `GET` and `POST`.
- For `POST`, merges `hook.Body` with `web.allProperties()`.
- Supports long-running `202 Accepted` responses by polling the `Location` header (resolved against the URL that answered, after any redirect) every 5 seconds until a non-202 response. A `202` without `Location`, or any request that fails or cannot be sent (network, CORS), fails the hook with its name.

Important behavior:

- Hooks run concurrently with `Promise.all`.
- Unsupported methods are logged and ignored.
- Hook failures include status details from JSON error responses when possible.

## PnP, JSOM, And Browser Assumptions

This library uses both PnPjs and SharePoint JSOM.

PnPjs is used for most CRUD-style operations:

- web metadata
- fields
- lists
- views
- files
- client-side pages
- custom actions
- themes
- hooks support data

JSOM is used where PnPjs does not cover the needed behavior cleanly or where classic SharePoint APIs are required:

- content type field links and ordering
- classic web parts on files/pages
- property bag indexing

Several modules assume browser globals (`src/util/index.ts` no longer does: it reaches them through `browserWindow()`):

- `window`
- `document`
- `fetch`
- `Blob`
- `_spPageContextInfo`
- global `SP`

Some handlers explicitly reject Node or SPFx modes, but not all browser/global assumptions are guarded. When making a handler work in a new runtime, audit the full call path for these globals.

## Build And Package Layout

Source:

- `src/**/*.ts`

Generated output:

- `lib/**/*.js`
- `lib/**/*.d.ts`
- `lib/**/*.js.map`

Build command:

```sh
npm run build
```

The TypeScript compiler:

- targets `es5`
- emits ES modules
- emits declarations
- writes to `lib`
- excludes `lib` and `node_modules`
- includes DOM and ES2016 libraries
- uses SharePoint typings from `@types/sharepoint`

Do not hand-edit `lib/`. Change `src/` and rebuild.

## Error And Logging Model

Logging uses `@pnp/logging`.

`WebProvisioner.onSetup()` subscribes `ConsoleListener()` once per process (a module-level flag; every `applyTemplate` used to add one) and sets `Logger.activeLogLevel` when `config.logging` is present.

Handlers use `HandlerBase` logging helpers for consistent messages:

- `log_info`
- `log_warn`
- `log_error`

Errors:

- Handler-level errors should be thrown.
- `WebProvisioner.applyTemplate` catches handler errors and throws `ProvisioningError`.
- `ProvisioningError.handler` records the handler name active at failure, `cause` what it threw or rejected with; the message is readable for an `Error`, a string or nothing (`undefined`).

Important caveat:

- Some handler internals catch errors and only log them. They log them with `log_error` (a refused in-use content type removal with `log_warn`), so a listener that collects errors sees them, but `applyTemplate` success still does not prove every sub-operation succeeded: SharePoint can also accept a request and change nothing (REST's `SetUserAsOwner` with a group id). Reading the result back is the proof (`docs/specs/integration-testing.md`).

## Extension Guidelines For Agents

When adding or modifying behavior:

1. Start with `src/schema.ts`.
   - Confirm whether the template contract changes.
   - Keep property names aligned with handler names when adding top-level sections.

2. Check handler order.
   - If the handler needs IDs or objects created elsewhere, update `DefaultHandlerSort` accordingly.
   - Do not rely on object key order in user templates.

3. Respect `ProvisioningContext`.
   - Keep key formats stable.
   - Load context before branching on existence.
   - Update context after creating new artifacts that later handlers may reference.

4. Prefer scoped changes in handlers.
   - Each handler owns one SharePoint object category.
   - Avoid cross-handler side effects unless they are represented in `ProvisioningContext`.

5. Be explicit about runtime assumptions.
   - Browser-only handlers should fail early with clear errors.
   - SPFx restrictions should be documented in code and specs.

6. Verify with the tests ([`docs/TESTING.md`](./TESTING.md)).
   - `npm run build`, `npm test`, `npm run typecheck:tests`
   - the handler's integration suite, which applies its fixture twice to a fresh scratch web and reads the result back: `npm run test:integration -- <Handler>`
   - a template by hand on the reserved site: `npm run provision -- --package <path> --twice --strict` (`debug/README.md`)

7. Keep docs close to architectural decisions.
   - Use `docs/specs/` for behavior and implementation plans.
   - Use `docs/adr/` when a decision should survive beyond one PR.

## Known Risk Areas

- Content type creation and update semantics are sensitive to SharePoint ID rules and JSOM/PnP differences.
- Handler order is a contract. Changing sort values can break token resolution, content type bindings, or page/list references.
- Token replacement is partial and regex-limited. Do not assume arbitrary token values are supported. Data row string values are now token-replaced, so a literal `{lowercasekey}` in seeded content is only safe if the key is unknown to `TokenHelper`.
- Some errors are swallowed after logging, especially inside site column, list, file, and content type sub-operations.
- Browser globals are used in utility and handler code, so Node compatibility is uneven.
- `.npmignore` and `.gitignore` affect whether docs and generated files are visible to consumers. The repository tracks docs; `lib/` remains generated and ignored by git.
