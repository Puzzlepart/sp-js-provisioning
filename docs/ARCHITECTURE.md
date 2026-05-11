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
- Loads `context.web` with `await this.web()`.
- Sorts top-level template keys with `DefaultHandlerSort`.
- Filters to requested handlers if `handlers` is supplied.
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
| 3 | `Lists` | Creates lists, binds content types, adds list fields, field refs, and views. |
| 4 | `Files` | Uploads files and classic pages, can add web parts and list view web parts. |
| 5 | `CustomActions` | Adds user custom actions after base structures exist. |
| 6 | `ComposedLook` | Applies theme assets, usually after files are available. |
| 7 | `ClientSidePages` | Creates modern pages and page controls. |
| 8 | `PropertyBagEntries` | Sets web property bag values after core structures exist. |
| 9 | `Navigation` | Rebuilds navigation after pages/lists may exist. |
| 10 | `WebSettings` | Applies final web/root folder settings. |
| 11 | `Hooks` | Calls external endpoints after provisioning has completed. |

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
- `lists: { [title: string]: string }` - list title to list ID.
- `listViews: { [listAndView: string]: string }` - key format is `${listTitle}|${viewTitle}`.
- `siteFields: { [internalName: string]: string }` - site field internal name to field ID.
- `contentTypes: { [nameOrId: string]: IContentType }` - content type entries keyed by both name and ID.

Context is the main dependency bridge between handlers. Be careful when changing key formats or timing. If context loading is async, it must be awaited before dependent decisions are made.

## Token System

Token replacement is split between two helpers.

### `TokenHelper`

Defined in `src/util/tokenhelper.ts`.

Supports context/config tokens with the pattern `{token:value}`:

- `{listid:List Title}` -> `context.lists[List Title]`
- `{listviewid:List Title|View Title}` -> `context.listViews[...]`
- `{webid:...}` -> `context.web.Id`
- `{siteid:...}` -> `context.web.Id`
- `{sitecollectionid:...}` -> `context.web.Id`
- `{parameter:Name}` -> `config.parameters.Name`

Important limitations:

- The regex only matches lowercase token keys and a limited value character set.
- It replaces matches only when a value exists.
- Site, web, and site collection IDs currently all resolve to `context.web.Id`.

### URL Tokens

Defined in `src/util/index.ts` as `replaceUrlTokens`.

Supports:

- `{site}` -> site server-relative URL
- `{sitecollection}` -> site absolute URL
- `{wpgallery}` -> site absolute URL plus `/_catalogs/wp`
- `{hosturl}` -> current browser protocol, host, and port
- `{themegallery}` -> site absolute URL plus `/_catalogs/theme/15`

URL token replacement uses `config.spfxContext` when available. Otherwise it falls back to browser globals such as `_spPageContextInfo` and `window`.

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
- Handler errors should be thrown so `WebProvisioner` can wrap them in `ProvisioningError`.

### `SiteFields`

File: `src/handlers/sitefields.ts`

Responsibilities:

- Loads existing web fields into `context.siteFields`.
- Processes field XML strings.
- Replaces context/config tokens in field XML.
- Parses XML using `xml-js`.
- Updates existing fields by `Name`.
- Creates missing fields with `web.fields.createFieldAsXml`.

Key contract:

- Site field identity is keyed by internal name.
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

### `Lists`

File: `src/handlers/lists.ts`

Responsibilities:

- Initializes JSOM for list operations that need it.
- Loads existing lists into `context.lists`.
- Ensures or creates lists.
- Adds content type bindings.
- Optionally removes existing content type bindings.
- Deletes and recreates list fields from field XML.
- Adds or updates list field refs from web fields.
- Adds, updates, removes, and reorders view fields.
- Updates `context.listViews`.

Processing phases inside `ProvisionObjects`:

1. Load existing lists.
2. Process each list definition.
3. Process list fields for all lists.
4. Process list field refs for all lists.
5. Process views for all lists.
6. Reload list IDs into context.

Key contracts:

- List identity is title-based in context.
- Content type binding uses `ContentTypeID`.
- `RemoveExistingContentTypes` keeps folder content types containing `0x0120`.
- List field refs match fields by field ID, not name.
- `SourceID` for copied web fields is set to the list ID from context.

Important behavior:

- `processField` deletes a list field by ID before recreating it. This is intentional for fields that cannot be updated cleanly, but it is destructive for list-local field state.
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
- Replaces config/context tokens in web part properties and server-processed content.
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

### `Features`

File: `src/handlers/features.ts`

Responsibilities:

- Activates or deactivates web features.
- Processes features sequentially.

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

File: `src/handlers/navigation.ts`

Responsibilities:

- Rebuilds Quick Launch and Top Navigation Bar trees.
- Deletes existing nodes in the target collection.
- Adds configured nodes recursively.
- Replaces URL tokens.

Important behavior:

- Existing nodes are deleted before adding configured nodes.
- If a configured node title matches one existing node and `IgnoreExisting` is not true, the existing URL is reused.
- Children are processed recursively by calling the same tree processor on the child collection.

### `WebSettings`

File: `src/handlers/websettings.ts`

Responsibilities:

- Replaces URL tokens in string settings.
- Updates web properties.
- Updates root folder `WelcomePage` separately when supplied.

Important behavior:

- If `WelcomePage` is absent, the current implementation does not call `web.update(settings)`. Be careful if changing this behavior, because it may expose settings that were previously ignored.

### `Hooks`

File: `src/handlers/hooks.ts`

Responsibilities:

- Runs external HTTP hooks after provisioning.
- Supports `GET` and `POST`.
- For `POST`, merges `hook.Body` with `web.allProperties()`.
- Supports long-running `202 Accepted` responses by polling the `location` header every 5 seconds until a non-202 response.

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

Several modules assume browser globals:

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

`WebProvisioner.onSetup()` subscribes `ConsoleListener()` and sets `Logger.activeLogLevel` when `config.logging` is present.

Handlers use `HandlerBase` logging helpers for consistent messages:

- `log_info`
- `log_warn`
- `log_error`

Errors:

- Handler-level errors should be thrown.
- `WebProvisioner.applyTemplate` catches handler errors and throws `ProvisioningError`.
- `ProvisioningError.handler` records the handler name active at failure.

Important caveat:

- Some handler internals catch errors and only log them. This means `applyTemplate` success does not always prove every sub-operation succeeded.

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

6. Verify with build and targeted SharePoint smoke tests.
   - `npm run build`
   - relevant handler scenario on a fresh site
   - relevant handler scenario on an existing site

7. Keep docs close to architectural decisions.
   - Use `docs/specs/` for behavior and implementation plans.
   - Use `docs/adr/` when a decision should survive beyond one PR.

## Known Risk Areas

- Content type creation and update semantics are sensitive to SharePoint ID rules and JSOM/PnP differences.
- Handler order is a contract. Changing sort values can break token resolution, content type bindings, or page/list references.
- Token replacement is partial and regex-limited. Do not assume arbitrary token values are supported.
- List field deletion/recreation can lose list-specific state.
- Some errors are swallowed after logging, especially inside list, file, and content type sub-operations.
- Browser globals are used in utility and handler code, so Node compatibility is uneven.
- `.npmignore` and `.gitignore` affect whether docs and generated files are visible to consumers. The repository tracks docs; `lib/` remains generated and ignored by git.
