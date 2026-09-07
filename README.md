# sp-js-provisioning

[![npm version](https://img.shields.io/npm/v/sp-js-provisioning.svg)](https://www.npmjs.com/package/sp-js-provisioning)
[![license](https://img.shields.io/npm/l/sp-js-provisioning.svg)](./LICENSE)

Template-based SharePoint Online provisioning in pure JavaScript/TypeScript. Apply a JSON
template to a web — site fields, content types, lists (fields, views, folders, seeded rows,
permissions), files, client-side pages, navigation, taxonomy and more — from the browser
(SharePoint Framework) or from Node. Built on [PnPjs](https://pnp.github.io/pnpjs/) and used by
[Prosjektportalen 365](https://github.com/Puzzlepart/prosjektportalen365) for template-driven
project setup.

## Installation

```shell
npm install sp-js-provisioning --save
```

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

Running from Node works with a PnPjs web set up via `@pnp/nodejs` — see
[`debug/`](./debug) for a ready-made runner with MSAL certificate auth.

## The template

A template is a JSON object (TypeScript type `Schema`) where each top-level key is handled by
a dedicated handler, applied in this order:

| Order | Section              | Provisions                                                                                                                                                                                                   |
| ----- | -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 1     | `Taxonomy`           | Term groups, term sets and terms (ids, sort order, custom properties). `UpdateExistingTerms` fills missing custom properties onto existing terms                                                             |
| 2     | `SiteFields`         | Site columns from field XML                                                                                                                                                                                  |
| 3     | `ContentTypes`       | Content types with fixed ids, field refs (by name) and field link order                                                                                                                                      |
| 4     | `Features`           | Web feature activation and deactivation                                                                                                                                                                 |
| 5     | `Lists`              | Lists and libraries: settings, content type bindings, list fields, field refs (site columns added to the list, by id), views, folder hierarchies, seeded rows (`DataRows`) and list permissions (`Security`) |
| 6     | `Files`              | Files with web parts, properties and overwrite control                                                                                                                                                              |
| 7     | `CustomActions`      | User custom actions, including SPFx extension registrations                                                                                                                                                  |
| 8     | `ComposedLook`       | Theme (composed look)                                                                                                                                                                                        |
| 9     | `ClientSidePages`    | Modern pages with sections and web parts                                                                                                                                                                     |
| 10    | `PropertyBagEntries` | Property bag values                                                                                                                                                                                          |
| 11    | `Navigation`         | Quick launch and top navigation nodes                                                                                                                                                                        |
| 12    | `WebSettings`        | Web properties such as welcome page and master page                                                                                                                                                          |
| 13    | `Hooks`              | HTTP callbacks after provisioning                                                                                                                                                                            |

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
name or a user login name. Role definitions are resolved by their name as it exists on the
web (localized), with a fallback to the well-known role types for the English names
`Full Control`, `Design`, `Edit`, `Contribute` and `Read`.

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
| `{sitecollectionid}`                                                                    | Id of the site collection                                                            |
| `{sitecollectiontermstoreid}` / `{termstoreid}`                                         | Id of the default term store — enables managed metadata site columns from a template |
| `{parameter:Name}`                                                                      | Value from `IProvisioningConfig.parameters`                                          |
| `{site}` / `{sitecollection}`                                                           | Server-relative / absolute URL (in URLs)                                             |
| `{associatedownergroupid}` / `{associatedmembergroupid}` / `{associatedvisitorgroupid}` | Associated group (in `Security` principals)                                          |

## Configuration

`setup(config)` accepts:

- `parameters` — string map used by `{parameter:...}` tokens.
- `spfxContext` — the SPFx context; used for URL token resolution in the browser.
- `logging` — `prefix` and `activeLogLevel` for the built-in `@pnp/logging` console listener.

Errors are thrown as `ProvisioningError`, carrying the name of the handler that failed.
Sub-operations such as individual data rows, folders and role assignments log failures and
continue, so one bad row does not abort the template.

## Development

```shell
git clone https://github.com/Puzzlepart/sp-js-provisioning.git
cd sp-js-provisioning
nvm use        # Node version from .nvmrc
npm install    # .npmrc enables legacy peer dependency resolution
npm run build  # tsc → lib/
```

`npm run provision` runs a template against a real site from Node — see
[`debug/README.md`](./debug/README.md) for setup. Changes are documented in
[`CHANGELOG.md`](./CHANGELOG.md).

## Contributing

Contributions are welcome. Please open an issue or submit a pull request on the
[GitHub repository](https://github.com/Puzzlepart/sp-js-provisioning).

## License

[MIT](./LICENSE) © Prosjektportalen. This repository continues the earlier
`pnp-js-provisioning` project.
