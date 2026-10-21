# sp-js-provisioning

[![npm version](https://img.shields.io/npm/v/sp-js-provisioning.svg)](https://www.npmjs.com/package/sp-js-provisioning)
[![license](https://img.shields.io/npm/l/sp-js-provisioning.svg)](./LICENSE)

Template-based SharePoint Online provisioning in pure JavaScript/TypeScript. Apply a JSON
template to a web — site fields, content types, site groups and permissions, lists (fields,
views, folders, seeded rows, permissions), files, client-side pages, navigation, taxonomy and
more — from the browser
(SharePoint Framework) or from Node. Built on [PnPjs](https://pnp.github.io/pnpjs/) and used by
[Prosjektportalen 365](https://github.com/Puzzlepart/prosjektportalen365) for template-driven
project setup.

## Installation

```shell
npm install sp-js-provisioning @pnp/sp @pnp/core @pnp/logging @pnp/queryable --save
```

[PnPjs](https://pnp.github.io/pnpjs/) v4 (`@pnp/sp` 4.x) is a **peer dependency**: the `IWeb`
you pass to `WebProvisioner` must come from a v4 `spfi(...)`, and it has to come from *your*
copy of `@pnp/sp`. PnPjs classes carry protected members, so two physically distinct copies of
the package are not assignable to each other and the build would fail on the `WebProvisioner`
constructor. Declaring the `@pnp/*` packages as peers lets the package manager enforce the
single instance. npm 7+ installs them for you; with `legacy-peer-deps` or Yarn 1, install them
explicitly as shown above.

## Quick start

```ts
import { LogLevel } from '@pnp/logging'
import { WebProvisioner, Schema } from 'sp-js-provisioning'

const template: Schema = {
  WebSettings: {
    WelcomePage: 'SitePages/Home.aspx'
  },
  Lists: [
    {
      Title: 'Checklist',
      Description: '',
      Template: 100,
      ContentTypesEnabled: false,
      AdditionalSettings: { EnableVersioning: true },
      DataRows: {
        KeyColumn: 'Title',
        UpdateBehavior: 'Skip',
        Rows: [{ Title: 'First checkpoint' }]
      }
    }
  ]
}

// `web` is a PnPjs IWeb for the site you are provisioning
const provisioner = new WebProvisioner(web).setup({
  spfxContext: this.context, // in SPFx; in Node, pass a minimal page context (see debug/)
  logging: { prefix: '(MySolution)', activeLogLevel: LogLevel.Info },
  parameters: { MyParameter: 'value' } // resolved by {parameter:MyParameter}
})

await provisioner.applyTemplate(template, undefined, (handler) => {
  console.log(`Running handler ${handler}`)
})
```

`applyTemplate(template, handlers?, progressCallback?)` applies every section of the template
in a fixed order (see below). Pass an array of handler names as the second argument to run a
subset, e.g. `['SiteFields', 'ContentTypes', 'Lists']`.

Running from Node works with any PnPjs v4 web, for example one set up with `@pnp/nodejs` (install
it yourself; it is not a dependency of this package). [`debug/`](./debug) has a ready-made runner
that signs in with a Playwright session or an app certificate (`debug/lib/auth.ts`).

## The template

A template is a JSON object (TypeScript type `Schema`) where each top-level key is handled by
a dedicated handler, applied in this order:

| Order | Section              | Provisions                                                                                                                                                                                                   |
| ----- | -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 1     | `Taxonomy`           | Term groups, term sets and terms (ids, sort order, custom properties). `UpdateExistingTerms` fills missing custom properties onto existing terms                                                             |
| 2     | `SiteFields`         | Site columns from field XML                                                                                                                                                                                  |
| 3     | `ContentTypes`       | Content types with fixed ids, field refs (by name) and field link order                                                                                                                                      |
| 4     | `Features`           | Web feature activation and deactivation                                                                                                                                                                      |
| 5     | `Security`           | Site groups (properties, owner, members) and role inheritance and role assignments on the web                                                                                                                |
| 6     | `Lists`              | Lists and libraries: settings, content type bindings, list fields, field refs (site columns added to the list, by id), views, folder hierarchies, seeded rows (`DataRows`) and list permissions (`Security`) |
| 7     | `Files`              | Files with web parts, properties and overwrite control                                                                                                                                                       |
| 8     | `CustomActions`      | User custom actions, including SPFx extension registrations                                                                                                                                                  |
| 9     | `ComposedLook`       | Theme (composed look)                                                                                                                                                                                        |
| 10    | `ClientSidePages`    | Modern pages with sections and web parts                                                                                                                                                                     |
| 11    | `PropertyBagEntries` | Property bag values                                                                                                                                                                                          |
| 12    | `Navigation`         | Quick launch and top navigation nodes                                                                                                                                                                        |
| 13    | `WebSettings`        | Web properties such as welcome page and master page                                                                                                                                                          |
| 14    | `Hooks`              | HTTP callbacks after provisioning                                                                                                                                                                            |

[`schema.json`](./schema.json) is a JSON Schema for the template format — reference it from
your template files for editor validation and IntelliSense:

```json
{
  "$schema": "https://raw.githubusercontent.com/Puzzlepart/sp-js-provisioning/main/schema.json"
}
```

See [`sample-schemas/`](./sample-schemas) for a small end-to-end example and
[`docs/ARCHITECTURE.md`](./docs/ARCHITECTURE.md) for a deep dive into every handler,
the token system and the error model.

### Seeded list rows (`DataRows`)

Lists can be seeded with items. Rows are upserted by `KeyColumn`: when a row with the same
key value exists, it is updated (`UpdateBehavior: "Overwrite"`, the default) or left alone
(`"Skip"`), so re-applying a template does not duplicate content. Field values are mapped by
the field's type — lookup and user values resolve to item/user ids, taxonomy values
(`{ termId, label }`) write the hidden note field, and string values get token replacement.

### Site groups and web permissions (`Security`)

```json
"Security": {
  "SiteGroups": [
    {
      "Title": "{sitename} Styringsgruppe",
      "Description": "Styringsgruppen for {sitename}",
      "Owner": "{associatedownergroupid}",
      "AllowMembersEditMembership": false,
      "OnlyAllowMembersViewMembership": false,
      "Members": ["kari.nordmann@contoso.no"],
      "ClearExistingMembers": false
    }
  ],
  "BreakRoleInheritance": false,
  "RoleAssignments": [
    { "Principal": "{sitename} Styringsgruppe", "RoleDefinition": "Lese" }
  ]
}
```

Site groups are matched by title, so re-applying a template updates an existing group instead
of creating a duplicate. Members are login names (UPN, email or claims, e.g. an Entra ID group
claim) and are ensured on the web before they are added. `ClearExistingMembers` removes members
that are not in `Members`, but only when every listed member could be added. `Owner` is resolved
like a role assignment principal (below) and may be the group itself; without it SharePoint
keeps the provisioning user as owner.

`BreakRoleInheritance`, `CopyRoleAssignments`, `ClearSubscopes` and `RoleAssignments` work as for
lists, but on the web. Inheritance is only broken when the web still inherits, so re-running a
template with `ClearSubscopes` does not wipe list permissions set since. The root web of a site
collection always has unique permissions and takes role assignments directly.

`Security` runs right before `Lists`, so list permissions can grant access to the groups it
creates. If you pass your own `handlerSort` to `WebProvisioner`, give `Security` a lower value
than `Lists`.

### List permissions (`Security`)

```json
"Security": {
  "BreakRoleInheritance": true,
  "CopyRoleAssignments": false,
  "ClearSubscopes": false,
  "RoleAssignments": [
    { "Principal": "{associatedownergroupid}", "RoleDefinition": "Full kontroll" },
    { "Principal": "{associatedmembergroupid}", "RoleDefinition": "Lese" }
  ]
}
```

A principal can be one of the associated-group tokens above, a principal id, a site group
name (including groups created by `Security` in the same template) or a user login name. Other
tokens such as `{sitename}` and `{parameter:Name}` are replaced first. Role definitions are
resolved by their name as it exists on the web (localized), with a fallback to the well-known
role types for the English names `Full Control`, `Design`, `Edit`, `Contribute` and `Read`.

### Tokens

Strings in field XML, data row values, file sources and client-side page properties can
reference values that are only known at provisioning time. Both `{token}` and `{token:value}`
forms are supported; unresolved or unknown tokens are left untouched. The URL tokens
`{site}` / `{sitecollection}` are the exception: they need a page context (SPFx, or a minimal
one in Node — see [`debug/`](./debug)) and resolve to `null` without one.

| Token                                                                                   | Resolves to                                                                          |
| --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| `{listid:List Title}`                                                                   | Id of a list on the web                                                              |
| `{listviewid:List Title\|View Title}`                                                   | Id of a list view                                                                    |
| `{webid}` / `{siteid}`                                                                  | Id of the web being provisioned                                                      |
| `{sitename}`                                                                            | Title of the web being provisioned                                                   |
| `{sitecollectionid}`                                                                    | Id of the site collection                                                            |
| `{sitecollectiontermstoreid}` / `{termstoreid}`                                         | Id of the default term store — enables managed metadata site columns from a template |
| `{parameter:Name}`                                                                      | Value from `IProvisioningConfig.parameters`                                          |
| `{site}` / `{sitecollection}`                                                           | Server-relative / absolute URL (in URLs)                                             |
| `{associatedownergroupid}` / `{associatedmembergroupid}` / `{associatedvisitorgroupid}` | Associated group (in `Security` principals and group owners)                         |

## Configuration

`setup(config)` accepts:

- `parameters` — string map used by `{parameter:...}` tokens.
- `spfxContext` — the SPFx context; used for URL token resolution in the browser.
- `logging` — `prefix` and `activeLogLevel` for the built-in `@pnp/logging` console listener.

Errors are thrown as `ProvisioningError`, carrying the name of the handler that failed (`handler`)
and what it failed with (`cause`).
Sub-operations such as individual data rows, folders, site groups, group members and role
assignments log failures and continue, so one bad row does not abort the template.

## Development

```shell
git clone https://github.com/Puzzlepart/sp-js-provisioning.git
cd sp-js-provisioning
nvm use        # Node version from .nvmrc
npm ci         # .npmrc enables legacy peer dependency resolution
npm run build  # tsc → lib/
npm test       # node:test unit tests
npm run typecheck:tests  # type-checks the tests together with src
```

`npm run provision` runs a template against a real site from Node, signed in from `.env`
(`.env.example`) — see [`debug/README.md`](./debug/README.md) for setup. Changes are documented in
[`CHANGELOG.md`](./CHANGELOG.md).

## Contributing

Contributions are welcome. Please open an issue or submit a pull request on the
[GitHub repository](https://github.com/Puzzlepart/sp-js-provisioning).

## License

[MIT](./LICENSE) © Prosjektportalen. This repository continues the earlier
`pnp-js-provisioning` project.
