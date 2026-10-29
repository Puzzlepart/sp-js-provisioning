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
| 2     | `SiteFields`         | Site columns from field XML: created when missing, updated in place on a re-apply (see Re-applying site columns)                                                                                             |
| 3     | `ContentTypes`       | Content types with fixed ids, field refs (by name) and field link order                                                                                                                                      |
| 4     | `Features`           | Web feature activation and deactivation                                                                                                                                                                      |
| 5     | `Security`           | Site groups (properties, owner, members) and role inheritance and role assignments on the web                                                                                                                |
| 6     | `Lists`              | Lists and libraries: settings, content type bindings, list fields, field refs (site columns added to the list, by id), views, folder hierarchies, seeded rows (`DataRows`) and list permissions (`Security`) |
| 7     | `Files`              | Files with web parts, properties and overwrite control                                                                                                                                                       |
| 8     | `CustomActions`      | User custom actions, including SPFx extension registrations                                                                                                                                                  |
| 9     | `ComposedLook`       | Theme (composed look)                                                                                                                                                                                        |
| 10    | `ClientSidePages`    | Modern pages with sections and web parts                                                                                                                                                                     |
| 11    | `PropertyBagEntries` | Property bag values                                                                                                                                                                                          |
| 12    | `Navigation`         | Quick Launch and top navigation nodes, reconciled in place: the template owns each collection it names                                                                                                       |
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
(`"Skip"`), so re-applying a template does not duplicate content. The key is matched with its
tokens replaced, as the item holds it. Field values are mapped by the field's type — lookup and
user values resolve to item/user ids, taxonomy values (`{ termId, label }`) write the hidden note
field, and string values (a URL's `Url` and `Description` too) get token replacement. A lookup or
user value that matches nothing is logged as an error and left out: a new item gets the values
that resolved, an existing one keeps what it holds in that column.

### Re-applying lists

A template can be applied again to a site that holds data (PP365 re-runs project setup and
re-imports packages), and that changes no data ([spec](./docs/specs/lists-re-apply.md)):

- A list field in `Fields` that exists is updated in place, never deleted: the template merged onto
  what SharePoint stores (what the template declares wins; what it leaves out keeps its value). A widening type change (`Text` to `Note`, `Lookup` to `LookupMulti`,
  `User` to `UserMulti`) is applied; another type change is a warning and leaves the field as it
  is. A lookup pointed at another list, or a taxonomy field at another term set, follows the
  template on a list made in this setup; elsewhere it is a warning and keeps its target, because
  the stored values belong to it.
- `RemoveExistingViews` and `RemoveExistingContentTypes` act on lists made in this setup: created
  by this provisioner, across its `applyTemplate` calls (a template and then its extensions). On a
  list that existed they do nothing. Neither removes a view the template names, a personal view,
  the default view or the view behind a page's list web part. (PnP Framework removes on every run.)
- An existing list gets only the settings the template defines; one it leaves out keeps its value.
- The views the template names are the template's: set when they differ from it.
- List `Security` breaks inheritance only on a list that still inherits; role assignments are
  added, never removed.

### Re-applying site columns

`SiteFields` follows the same rules as list fields ([spec](./docs/specs/sitefields-re-apply.md)):

- A site column is identified by its internal name (`Name`). One the web has is updated in place:
  the template merged onto what SharePoint stores (what the template declares wins; what it leaves
  out keeps its value, so an index, formatter or description an admin added survives), and nothing
  is written when that changes nothing. A widening type change (`Text` to `Note`, `Lookup` to
  `LookupMulti`, `User` to `UserMulti`) is applied; another type change is a warning and leaves the
  column as it is.
- A lookup pointed at another list, or a taxonomy column at another term set or term store, is a
  warning and keeps its target, unless the column was made in this setup. A target an earlier run
  stored as a literal token (or the empty GUID SharePoint stores for a literal term store id)
  follows the template, so the column heals.
- Site column changes do not reach the column's existing list copies: the engine never pushes
  them (SharePoint does not either, through the column or through a content type update).
- A column with the template column's id under another name is an error and is not created. On a
  subweb, a column a parent web has is neither created nor changed (a warning).
- A column SharePoint refuses is logged as an error, and the other columns and the handlers after
  `SiteFields` still run.

### Navigation

```json
"Navigation": {
  "QuickLaunch": [
    { "Title": "Hjem", "Url": "" },
    {
      "Title": "Prosjektinformasjon",
      "Url": "SitePages/Prosjektinformasjon.aspx",
      "Children": [{ "Title": "Fasesjekkliste", "Url": "Lists/Fasesjekkliste" }]
    },
    { "Title": "Notatblokk", "Url": "#", "IgnoreExisting": true }
  ]
}
```

The template owns each collection it names (`QuickLaunch`, `TopNavigationBar`); one it leaves
out is not touched. A run that succeeds leaves exactly the template's nodes in the collection, in
its order, children included ([spec](./docs/specs/navigation-re-apply.md)):

- A node whose title is already there at the same level (exact title, case included) is kept as
  it is: its id, its audiences and its URL. Nothing is written for it, so re-applying an unchanged
  template changes nothing.
- A missing node is added; a node the template does not declare is deleted, with its children
  (SharePoint's default nodes, a node a site owner added). A kept node without `Children` in the
  template loses the children it has.
- `IgnoreExisting: true` sets the template's `Url` on a node that exists, where it differs (`%20`
  or a space, a trailing slash, the case of a SharePoint path, and a web-relative or absolute form
  of the same URL count as the same). Without it an existing node keeps its stored URL, at every
  level: SharePoint changes it when the target page is renamed, and an owner may have. A title
  stored more than once gets the template's URL either way.
- `Url` may be server-relative, web-relative (`SitePages/Home.aspx`, resolved against the web),
  absolute, or `#`. It takes the URL tokens only (`{site}` is the site collection, also on a
  subweb); `Title` takes none.
- SharePoint refuses a node whose URL names no existing file or folder. The run then fails with
  `Navigation node '<Title>' (<Url>) was refused by SharePoint: ...`, and the nodes and URLs the
  run had added or changed are undone, so both collections are left as they were. A page, file or
  list deleted since the last run takes its node with it, so the next run fails on that node.
- SharePoint returns the titles of its own nodes (`Hjem`, `Dokumenter`) in the language of the
  account that runs the template, so they match only for an account with the template's
  language; otherwise the node is added again and the stored one deleted.

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
forms are supported; unresolved or unknown tokens are left untouched. In site column XML, a known
token that cannot be resolved is also logged as an error naming the column and the token (the
column is still created with it as written). The URL tokens
`{site}` / `{sitecollection}` are the exception: they need a page context (SPFx, or a minimal
one in Node — see [`debug/`](./debug)) and resolve to `null` without one.

| Token                                                                                   | Resolves to                                                                          |
| --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------ |
| `{listid:List Title}`                                                                   | Id of a list on the web (in `SiteFields` also one the same template creates)         |
| `{listviewid:List Title\|View Title}`                                                   | Id of a list view (not in `SiteFields`, see below)                                   |
| `{webid}` / `{siteid}`                                                                  | Id of the web being provisioned                                                      |
| `{sitename}`                                                                            | Title of the web being provisioned                                                   |
| `{sitecollectionid}`                                                                    | Id of the site collection                                                            |
| `{sitecollectiontermstoreid}` / `{termstoreid}`                                         | Id of the default term store — enables managed metadata site columns from a template |
| `{parameter:Name}`                                                                      | Value from `IProvisioningConfig.parameters`                                          |
| `{site}` / `{sitecollection}`                                                           | Server-relative / absolute URL (in URLs)                                             |
| `{associatedownergroupid}` / `{associatedmembergroupid}` / `{associatedvisitorgroupid}` | Associated group (in `Security` principals and group owners)                         |

`SiteFields` runs before `Lists`, so `{listid:}` there resolves for lists that exist when the run
starts, and for a list the same template's `Lists` declares: `WebProvisioner` creates that list,
with its settings, before `SiteFields` (when both run and `SiteFields` comes first).
`{listviewid:}` does not resolve before `Lists` (for example in `SiteFields`), except for views an
earlier `applyTemplate` through the same provisioner processed.

## Configuration

`setup(config)` accepts:

- `parameters` — string map used by `{parameter:...}` tokens.
- `spfxContext` — the SPFx context; used for URL token resolution in the browser.
- `logging` — `prefix` and `activeLogLevel` for the built-in `@pnp/logging` console listener.

Errors are thrown as `ProvisioningError`, carrying the name of the handler that failed (`handler`)
and what it failed with (`cause`).
Sub-operations such as individual site columns, data rows, folders, site groups, group members
and role assignments log failures and continue, so one bad row does not abort the template.

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
