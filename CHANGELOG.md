# Changelog

All notable changes to this package are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
For releases older than 1.3.6, see the git history.

## [Unreleased]

For 1.6.0 (the JSOM exit, [spec](./docs/specs/jsom-exit.md)): every handler runs in Node and under
SPFx alike. ContentTypes, Taxonomy and PropertyBagEntries reach SharePoint through CSOM over HTTP,
Files through REST, and the package no longer depends on `spfx-jsom`. In PP365 a hub import of a
package with `Files` or `PropertyBagEntries` no longer stops there; a re-run of an unchanged
template writes no content type, term set, file or property; a field ref to a missing column, a
failing file and a property bag write a modern site refuses are logged errors that no longer stop
the template. Classic web parts in `Files` are no longer supported (Removed). Templates written for
1.5 work unchanged otherwise.

### Changed

- `ContentTypes` no longer needs JSOM: it reaches SharePoint through CSOM over HTTP
  (`_vti_bin/client.svc/ProcessQuery`, through the web's own PnPjs pipeline), so it runs in Node and
  the browser alike ([spec](./docs/specs/jsom-exit.md), [ADR](./docs/adr/csom-over-http.md)). A
  content type is still created with the template's id. What it does differently on a site:
  - A content type the web has gets only what differs: a name, description or group, a missing
    field link, a changed `Required` or `Hidden`, a changed order. An unchanged one is not sent
    `Update(true)`, which pushed it, and every list and child content type that inherits from it,
    again on each re-run of a hub import.
  - A field ref whose site column is not on the web is logged as an error, naming it, and the
    content type gets its other columns; it used to stop the template.
  - A field ref finds its column among the columns the web can use (on a subweb, its parent webs'
    too): by internal name, else by id, else by title.
  - The template's `ContentTypes` array is no longer sorted in place.
- `Taxonomy` no longer needs JSOM either: it reaches the default term store through CSOM over HTTP,
  with the template's ids as before. What it does differently:
  - It reads the store once, in one request, instead of one request per term, and writes a term set
    only when something is missing or its custom sort order differs: the sort order used to be set
    and committed on every run. A term set's writes are one commit, all or nothing.
  - A term store that cannot be read stops the template with the reason. A failed read used to be
    taken for a missing term set or term, which it then tried to create.
- `Files` runs under an SPFx context and in Node: it used to stop every template that had it
  whenever `config.spfxContext` was set, as PP365 always sets it (a `pp-testprosjekt` hub import
  stopped there). The source is read as bytes, so a binary file is no longer corrupted, and a
  source on another origin is fetched without credentials, as a public URL such as
  raw.githubusercontent.com requires; a source on the web's own origin is read through REST, which
  an app-only sign-in may read where the file's own URL answers 401. A file that already holds the
  source's bytes is not uploaded again, a property is written only when it differs, and
  `Overwrite: false` leaves an existing file (and its properties) as it is. A file that fails is
  logged as an error, naming it, and the next is tried.
- `PropertyBagEntries` runs under an SPFx context and in Node, and writes PnP's way: an entry whose
  key the web lacks is written even without `Overwrite` (it used to be written only with
  `Overwrite: true`), an existing one only with `Overwrite` and another value, and `Indexed` adds
  the key to the web's indexed keys instead of replacing them. SharePoint refuses the write on a
  site that blocks custom script (modern sites do by default): that is logged as an error naming
  the cause, and the template goes on instead of stopping.

### Removed

- The `spfx-jsom` dependency, and with it the SharePoint Framework packages it pulled in: no
  handler uses JSOM. PP365 keeps its own copy for its own tasks.
- Classic web parts in `Files` (`WebParts`, `RemoveExistingWebParts`, a web part's `ListView`): no
  template of PP365 or prosjektportalen-hosting uses them, and they need a classic page and custom
  script. A file that has them is provisioned, and an error says they were left out.

### Development

- Integration testing Phase 2, part 2 ([spec](./docs/specs/integration-testing.md)): `npm run verify`
  compares a template or a package with a web, read-only, and reports each difference
  (`debug/lib/verify.ts`, `debug/lib/readback/site.ts`): tokens replaced as the engine does for each
  handler, and what a re-apply makes true (lists, views and members as a superset, data rows by
  their key, a menu node's URL only with `IgnoreExisting`; `--exact` for a fresh web).
  `npm run inspect` prints what a web holds of what the handlers write, as JSON.
  `provision --verify` runs `verify` after applying, and `provision --mode project` applies a
  package's project template and extensions as the project setup's `ApplyTemplate` and `Hooks`
  steps do (`--extensions`). Every integration suite's `applyTwice` now also asserts that `verify`
  finds nothing; `tests/unit/verify.test.ts` covers its rules.
- Integration testing Phase 2, part 3: `npm run test:integration -- packages`
  (`tests/integration/packages.test.ts`) applies the prosjektportalen-hosting packages twice, each
  case on its own scratch web (`pp-testprosjekt` as a hub import, a cloud template and a project
  setup, in both locales; three templates of other packages; `INTEGRATION_PACKAGES=all` the rest),
  and holds the errors it logs and the differences `verify` finds to an exact list of what the Node
  run cannot do until the JSOM exit (`packages.expected.ts`: what each list lacks). It found two
  `verify` rules to correct (a view query is compared as SharePoint stores it, without the
  whitespace around an attribute's `=`; a list field ref without an `ID` is noted instead of
  reported as a missing column) and three such field refs in `pp-testprosjekt` (fixed in its
  1.1.1). `npm run test:integration -- all` adds `everything.test.ts`: `all.node.json`, every
  Node-capable fixture in one template (`npm run build:fixtures`), applied twice and verified with
  `--exact`; the `parameters` suite covers `{parameter:}`. A DataRows unit test sends DateTime
  values the same in Europe/Oslo, UTC and Asia/Tokyo. `provision` and `verify` fill a fixture's
  `@@user@@`. The `spjs-handlers` and `spjs-release` skills; the coverage floors rise to 80/84/79.
- The JSOM exit, steps 1 and 2 ([spec](./docs/specs/jsom-exit.md)): `src/util/csom.ts` builds and
  sends CSOM requests (`CsomRequest`, `executeQuery`, `CsomError`); `tests/helpers/csom.ts` reads a
  request back as lines and answers it like SharePoint (`parseCsom`, `FakeCsom`). The ContentTypes
  suite runs (`contenttypes.json`, `.b.json`): content types with the template's ids (a child of
  Item, its child, a child of Document), their links, flags and order, a second apply that sends no
  CSOM request, B's changes reaching the child and the bound list, a missing column, a refused
  name. `verify` asked SharePoint for content type ids with `0X` and found none; fixed. The runner
  still leaves ContentTypes out by default until the packages list counts it (step 5).
- The JSOM exit, step 5: every handler runs in Node, so the runner applies all of them but
  Taxonomy, which it applies only with `--taxonomy` (or `--handlers`): it writes to the tenant's
  term store, which every site of the tenant shares. The packages suite never applies a package's
  Taxonomy section, and its list (`packages.expected.ts`) now holds what a scratch web lacks that a
  PP365 hub has, and what the reserved site refuses. `all.node.json` is `all.json`, with the
  content types; `debug/stubs/spfx-jsom.ts`, its `paths` mappings and `@types/sharepoint` are gone.
  The packages suite makes a package's missing parent content type first (PP365's `Prosjekt` for
  `Prosjekt (Vei)`), as it makes menu targets, and `INTEGRATION_PACKAGES_REPORT=true` prints every
  line to write the list anew. The coverage floors rise to 90/85/86.
- The JSOM exit, step 4: the Files suite uploads a text and a binary file byte for byte from the
  scratch web and from a local server (another origin), sets a property, uploads nothing on a
  second apply, uploads a changed source again, leaves `Overwrite: false` alone and logs a failing
  source and classic web parts. The PropertyBagEntries suite proves, on the reserved site, which
  blocks custom script, one error naming the cause and nothing written; where custom script is
  allowed it applies twice and reads the entries back (it skips here). `failures.json`'s file
  source now fails on the site itself instead of on example.org.
- The JSOM exit, step 3: the Taxonomy suite runs against the tenant's term store, in a term group of
  its own (`SPJS Itest`, `taxonomy.json` and `.b.json`), created by its first run and left in place:
  nothing is deleted from the term store. It proves the ids, custom properties and sort order, a
  second apply that only reads, a sort order changed and back, and that `UpdateExistingTerms` does
  not overwrite a custom property. `verify` compares a term set's custom sort order, and a term's
  custom properties with `UpdateExistingTerms` (values, description and openness with `--exact`).
- Integration testing Phase 3, tenant CI: the `Integration` workflow
  (`.github/workflows/integration.yml`) runs the suites against the reserved site with
  prosjektportalen365's CI app. It starts only by hand from `main`, in the protected environment
  `tenant` (approval, the certificate, the user and the site as its secrets), one run at a time,
  with inputs for the suites, the hosting packages and keeping the scratch webs.
  `npm run test:integration:ci` also writes the HTML page (`test-results/integration/`), which a
  second job publishes to the test report site as the project's Integration column. The log and
  the artifact of a public repository are public, so `redact()` now also replaces certificates,
  `onmicrosoft.com` domains, the site in a `/sites/...` path, e-mail addresses and logins, the
  configured user, app and tenant, and every GUID that is not in the repository's own files, and
  `npm run check:reports` fails the job before the upload when a report still holds such a value.
  The first approved run (2026-11-10), also the first in certificate mode, was green.
- prettier also checks `debug/lib`: `.prettierignore`'s `lib/` (the build output) matched it too,
  so `npm run lint:check` never looked at the harness. The pattern is `/lib/` and the files are
  formatted.

## [1.5.1] - 2026-11-10

A handler that fails with something other than an `Error`, such as a refused JSOM request, is
described in words in the `ProvisioningError` and the log instead of «[object Object]» and
«undefined». `schema.json` is generated from `src/schema.ts`, so it describes every section,
including what 1.5.0 changed and the `Taxonomy`, `DataRows` and `Folders` it lacked.

### Fixed

- A handler that failed with something other than an `Error` is described in words. A refused JSOM
  request (the `ContentTypes` handler's `executeQueryAsync` rejects with `{ sender, args }`) made
  the `ProvisioningError` message «[object Object]» and the log line «Code execution scope ended
  with error: undefined». Both now read the request's `args.get_message()`, a SharePoint REST
  error's `message.value`, or else a short JSON of the object (`readableErrorMessage` in
  `src/provisioningerror.ts`).
- `schema.json` describes what 1.5.0 changed, copied from the TSDoc in `src/schema.ts`: list
  `Fields`, `RemoveExistingViews` and `RemoveExistingContentTypes`, `SiteFields`, `Navigation` and
  its nodes (`IgnoreExisting`, `Children`), web and list `Security`, role assignments, site groups
  and the optional `ComposedLook` URLs. An editor using the schema showed none of it, so the
  re-run rules were only in the README. `SiteFields` gets a TSDoc of its own, and
  `BreakRoleInheritance`'s says it is only done while the list or web still inherits.
- A view's `CustomFormatter` (a string) and `ViewType2` (`TILES`, `COMPACTLIST`, `KANBAN` or
  `MODERNCALENDAR`, as PnPjs types it) are in `IListView.AdditionalSettings` (`src/schema.ts` and
  `schema.json`), so a TypeScript template no longer needs a cast for a cards view or a view
  formatter. `Lists` already sent both again right after creating a view, since SharePoint
  ignores them on create; a unit test now covers that path.

### Changed

- `schema.json` is generated from `src/schema.ts` (`npm run build:schema`) instead of kept by hand,
  so it now also describes the sections the hand-made one lacked: `Taxonomy` (term groups, term
  sets, terms), a list's `DataRows` and `Folders`. A hook's `Headers` and `Body` are described as
  the JSON a template holds (an object of header strings, an object for the body) instead of DOM
  types. Properties stay optional, as before. Every fixture and every prosjektportalen-hosting
  template validates against it.

### Development

- The `@pnp/core`, `@pnp/logging`, `@pnp/queryable` and `@pnp/sp` devDependencies are 4.22.0,
  the newest PnPjs, so the build and the tests use what a new install of the peers gets; the
  peer range stays `^4.21.0`. The gulp-era `merge` and `merge2` devDependencies are gone (nothing
  imports them; `merge2` stays in the lock for `globby`).
- The unit tests' report as a page in the browser: `npm run test:ci` also writes
  `test-results/unit/index.html` with a node:test HTML reporter ported from prosjektportalen365
  (`tests/reporters/html.ts`, every text through `redact()`), and on each push to `main` the new
  `publish-report` job in `ci.yml` puts it on prosjektportalen365's private test report site, at
  `sp-js-provisioning/` ([spec](./docs/specs/test-report-site.md)). The publisher,
  `.tasks/publish-test-report.js`, is prosjektportalen365's current one, extended and shared byte
  for byte: `--project` writes only the project's folder, never the site's front page, a guard
  refuses any other change before the push, a failed publish is also written to the job summary,
  and prosjektportalen365's own publishes add an "Also on this site" link to each project's page
  once they use the shared file. Its tests run with `npm run test:tasks`, in CI too. Until the
  deploy key and the variables are set up, the job only says so.
- The unused `commitlint` devDependency and its `package.json` config are gone (no hook or CI ran
  it, and the config it extended was no longer installed).
- Integration testing Phase 2, part 1 ([spec](./docs/specs/integration-testing.md)):
  `.tasks/generate-schema.js` (`typescript-json-schema` 0.68.0) writes `schema.json`, and
  `npm run check:schema` fails in CI when it is out of date; `tests/unit/schema.test.ts` validates
  the fixtures and the hosting templates with `ajv` (now an explicit devDependency), and
  `tests/unit/fixture-coverage.test.ts` requires every schema property in a fixture or a reason it
  is not. `tests/contract/` checks what the handlers send through the real PnPjs queryables with a
  transport that answers in the process (`tests/helpers/transport.ts`): SiteFields, DataRows, the
  navigation reconcile and WebProvisioner's handler order; `npm test` runs them.

---

## [1.5.0] - 2026-11-04

Re-applying a template to a site in use is now safe for `Lists`, `SiteFields` and `Navigation`:
list fields, site columns and menu nodes that exist are updated in place instead of deleted and
recreated, so item values, node ids and what an admin set survive; an unchanged re-run writes no
site column and no menu node; a failure is logged naming what failed, and no longer stops the
template at the first site column or leaves half a menu. Templates written for 1.4.0 work unchanged
on a first run; what a re-run does differs, see Changed.

### Added

- New top-level `Security` section and handler, run right before `Lists`
  ([spec](./docs/specs/site-groups-and-web-security.md)):
  - `SiteGroups` creates SharePoint site groups, or updates the group when one with the same
    title exists, with description, membership settings, members and owner.
    `ClearExistingMembers` removes members that are not in the template.
  - `BreakRoleInheritance`, `CopyRoleAssignments`, `ClearSubscopes` and `RoleAssignments` apply
    permissions on the web, with the same principal and role definition rules as list
    `Security`. Inheritance is only broken when the web still inherits.
  - List `Security` can grant permissions to groups created by the same template.
- `{sitename}` token, resolving to the title of the web being provisioned (PnP semantics), e.g.
  for per-site group names such as `{sitename} Styringsgruppe`.

### Changed

- Re-applying a template to a site that holds data changes no data
  ([spec](./docs/specs/lists-re-apply.md)):
  - A list field in `Fields` that exists is updated in place instead of deleted and recreated, so
    its values, its place in views and content types and what an admin set that the template does
    not mention survive (a re-run of PP365 project setup emptied the lookup columns of the standard
    templates). A widening type
    change is applied; another type change is a warning and leaves the field as it is. A lookup
    pointed at another list, or a taxonomy field at another term set, follows the template on a
    list made in this setup and is otherwise a warning that keeps the stored target.
  - `RemoveExistingViews` and `RemoveExistingContentTypes` act on lists made in this setup (created
    by this provisioner, across its `applyTemplate` calls, so a template's extensions still count)
    and do nothing on a list that existed. They never remove a view the template names, a personal
    view, the default view or the view behind a page's list web part; template views keep their
    ids. PnP Framework removes on every run.
  - An existing list is updated with only the settings the template defines: `lists.ensure` sent an
    empty `Description` and `ContentTypesEnabled: false` for what it left out.
  - A template view's fields are set only when they differ; one field that cannot be added no
    longer leaves the view without the rest.
  - List `Security` breaks inheritance only on a list that still inherits, as web `Security` does.
- Re-applying `SiteFields` follows the list-field rules
  ([spec](./docs/specs/sitefields-re-apply.md)):
  - An existing site column is no longer rewritten with the template's XML on every run (which
    dropped what the template left out: an index, a formatter, a description). The template is
    merged onto the stored schema and written only when it changes something; a child element the
    template declares (`CHOICES`, `Default`, `Formula`) still replaces the stored one, so a choice
    an admin added to a column whose template lists `CHOICES` is reset, as before. A widening type
    change is applied; another type change is a warning and leaves the column as it is. A lookup or
    taxonomy column the template points at another list, term set or term store is a warning and
    keeps its stored target (the rest of the template is still merged), unless the column was made
    in this setup. Changes do not reach existing list copies of the column; the engine does not
    push them.
  - A site column that SharePoint refuses is logged as an error, and the other columns and the
    handlers after `SiteFields` still run; the template used to stop there.
  - A column with the template column's id under another internal name is an error and is not
    created (SharePoint refused it with "The local device name is already in use"). On a subweb, a
    column a parent web already has is a warning and is neither created nor changed (SharePoint
    refused it as a duplicate).
  - A known token that cannot be resolved in a site column's XML is logged as an error naming the
    column and the token; it is still written as before, except over a list, web, term store or
    term set the column already points at. A target an earlier run stored as a literal token heals
    on the next run.
- `{listid:}` in a site column now resolves to a list of the same template: `WebProvisioner`
  creates those lists, with their settings, before `SiteFields` (when `Lists` runs too, after
  `SiteFields`). The column used to hold the literal token after the first run, silently, and so
  did the copies made of it in that run.
- A list FieldRef naming a column found neither on the list nor on the web is logged as an error;
  it was skipped silently.
- Re-applying a template reconciles the navigation in place instead of deleting and re-adding
  every node ([spec](./docs/specs/navigation-re-apply.md)): a node whose title exists at the same
  level is kept with its id, audiences and URL, and nothing is written for it, so an unchanged
  re-apply changes nothing; missing nodes are added, nodes the template does not declare are
  deleted and the nodes are moved into the template's order. A kept child now keeps its stored URL
  as a top-level node does: a template change to an existing child's URL needs `IgnoreExisting`,
  which sets the template's URL only where it differs.
- `DataRows`: an unresolved Lookup or User value (and each unresolved entry of a LookupMulti or
  UserMulti) is logged as an error naming the column and the value; a new item gets the values
  that resolved and an existing one keeps what it holds in that column. A lookup or user request
  that may succeed later (throttled, unavailable) still fails the row, so the next run writes it
  whole; a taxonomy term without `termId` or `label` is logged instead of skipped silently. A LookupMulti entry that
  matched nothing used to make SharePoint refuse the whole row, and a single Lookup was dropped
  without a word. More than one item with the same key, and a row without a key value, are
  warnings.
- Tokens such as `{sitename}` and `{parameter:Name}` are now replaced in list `Security` role
  assignment principals before they are resolved.
- Principal and role definition resolution moved from the `Lists` handler to
  `src/util/security.ts`, shared by list and web `Security`. Site group names with an apostrophe
  are now escaped in the lookup url.
- `{hosturl}` resolves to the origin (`https://contoso.sharepoint.com`), from the SPFx context's
  web first, then the page's location. It used to end in a colon (`${protocol}//${host}:${port}`
  with an empty port), so a template that worked around that may need a change.
- The URL tokens, `makeUrlRelative` and the base64 encoding of indexed property bag keys work
  without a browser (`window`, `document`, `btoa`): in Node they no longer throw a
  `ReferenceError`, so no browser shims are needed there. In a browser nothing changes.
- `object.omit` is no longer a dependency: nothing uses it since `WebSettings` splits off the
  welcome page with object rest, so consumers no longer install it.

### Fixed

Found by the tests of the [integration testing spec](./docs/specs/integration-testing.md),
Phases 0b and 1; each fix has the test that exposed it.

- A `$select` left on the web passed to `WebProvisioner` no longer leaks into the next run.
  PnPjs 4's `select()` sets the query on the queryable itself, so after a handler read the web
  (ClientSidePages and Files read `ServerRelativeUrl`, web `Security` reads
  `HasUniqueRoleAssignments`) the next `applyTemplate` on the same web loaded only that property:
  `{sitename}`, `{webid}` and `{siteid}` stayed unresolved. PP365's setup wizard applies
  the template and then each extension to the same web. `WebProvisioner` now reads the web, and
  hands each handler the web, as a fresh queryable.
- `{associatedownergroupid}`, `{associatedmembergroupid}` and `{associatedvisitorgroupid}` on a web
  that has no such group of its own (one that inherits permissions) fail with an error naming it,
  instead of a role assignment sent with `principalid=undefined`.
- `Features` leaves a feature that is already in the requested state alone. SharePoint refuses to
  activate an active feature (500) and to deactivate an inactive one (400), so applying a template
  a second time stopped there with a `ProvisioningError`. `force` still re-activates an active
  feature, and a feature listed twice follows each entry in order.
- `WebSettings` applies its settings when the template has no `WelcomePage`; it used to update the
  web only together with a welcome page. It sends no update when there is nothing to set.
- `WebSettings` with a `WelcomePage` no longer fails with "omit is not a function" where
  `object.omit` is loaded as a namespace (Node); the welcome page is split off with object rest.
- `Hooks`: a POST hook answered with `202 Accepted` now finishes when its `Location` answers
  another 2xx; it used to wait forever. A poll request that fails now fails the hook instead of
  escaping as an unhandled rejection. A relative `Location` is resolved against the URL that
  answered (in a browser it was resolved against the SharePoint page), and a `202` without a
  `Location`, or with one that is not a URL, fails the hook, naming it, instead of polling `null`.
- `DataRows` matches a `KeyColumn` value with its tokens replaced, as the item holds it: a key such
  as `{sitename} Styringsgruppe` added the row again on every run. A skipped row no longer queries
  lookup lists or ensures users (nor fails when they throw), and a URL `{ Url }` without
  `Description` is described by its resolved Url instead of the raw one; a given `Description`
  gets token replacement.
- A template view without `AdditionalSettings` that already exists on the list (an OOTB view
  named in the template) is no longer sent an empty update, which SharePoint refused (400), so its
  fields were not set either.
- `FieldRefs` `AdditionalProperties` no longer set the column's title to its internal name, drop
  attributes, or send a stale schema version; field ref ids are matched whatever their braces and
  case.
- `Hooks`: a request that cannot be sent (a network or CORS failure) fails the hook with its name,
  as an error status does, instead of a bare "Failed to fetch".
- `ClientSidePages` replaces tokens in each string of a web part's properties and server-processed
  content instead of in their JSON text. A `{sitename}` or `{parameter:...}` value with a quote or
  a backslash (a web titled `Rehab "Gamle skole"`) made the text invalid JSON, and the page was
  saved without the part.
- `ComposedLook` applies a look without `FontSchemeUrl` (it threw a `TypeError`); `FontSchemeUrl`
  and `BackgroundImageUrl` are optional in `IComposedLook`, as `schema.json` already had them.
- `Navigation`: a node SharePoint refuses (a URL that names no existing file or folder) no longer
  leaves the menu with only the nodes added before it: nothing is deleted until every node is
  accepted, this run's adds and URL changes are undone, and the error names the node and its URL
  (`Navigation node '<Title>' (<Url>) was refused by SharePoint: ...`). The Quick Launch and the
  top navigation bar are written one after the other, so a failure in one no longer lets the other
  carry on unreported, and a refusal in either leaves both as they were. The template object is no
  longer changed (`node.Url` was overwritten with the stored URL).

- `Security`: a site group as `Owner` (an associated group token, a group title, or the group
  itself) is now applied. It goes through CSOM (`Group.Owner` over
  `_vti_bin/client.svc/ProcessQuery`), because REST's `SetUserAsOwner` answered success for a group
  id and changed nothing; the first tenant run found every group still owned by the provisioning
  user while the log said otherwise. The owner is read back, and an owner SharePoint did not change
  is logged as an error. A user owner still goes through `SetUserAsOwner`.
- Failures that `Lists` and `ClientSidePages` log and carry on after are logged as errors instead
  of Info: a content type binding, a list field, a view's fields, a list's root folder, a folder,
  a data row, a page control. A content type removal SharePoint refuses because the content type
  is in use is a warning. Listeners that collect errors (PP365's install log) now see them.
- `ProvisioningError` has a readable message whatever the handler failed with, keeps `handler`,
  adds the original error as `cause`, and is an `instanceof ProvisioningError` in the ES5 build
  (`lib/`). A `PropertyBagEntries` rejection used to make its constructor throw a `TypeError`,
  losing the handler name.
- `Files` and `PropertyBagEntries` fail with an `Error` that names the reason and is logged by the
  handler, instead of a bare string and an empty rejection. A template with either section still
  stops there when an SPFx context is set, as before.
- The console listener is subscribed once per process instead of on every `applyTemplate`. PP365's
  setup wizard applies a template and then each extension, and printed every line once per earlier
  call.
- `applyTemplate` no longer throws a `TypeError` when `setup()` was given no `logging`, or was not
  called.

### Development

Tooling only; nothing here ships in the package. Plan:
[integration testing spec](./docs/specs/integration-testing.md), Phases 0a, 0b and 1.

- Offline CI (`.github/workflows/ci.yml`, job `Build and offline tests`) on every pull request and
  push to `main`: `lint:check`, build, `typecheck:tests`, `test:ci` (spec + JUnit to
  `test-results/unit.xml`), `test:coverage` (source-mapped, floors 79/83/79 for lines, branches
  and functions, which only rise), `check:pack` (the package holds `lib/index.js` and every handler) and
  `check:secrets` (no tenant URL, cookie, token or certificate in any committed or new text
  file, docs and specs included). eslint's `quotes` rule allows `"''"` (it conflicted with prettier), and
  `.prettierignore` keeps `lib/` out of the formatting check.
- `tests/fixtures/templates/failures.json`: a list field that cannot be created, a `Files` and a
  `PropertyBagEntries` entry, for checking that failures surface (`--strict` exits 1 for each),
  and a site column SharePoint refuses followed by a valid one (`sitefields.test.ts`).
- Integration tests against a tenant (Phase 1): `npm run check:tenant`, `npm run test:integration`
  (`-- <Handler ...>`, `--keep`, `:ci`), one suite per handler in `tests/integration/` applying a
  fixture twice to a scratch web (a subweb of the reserved site with its own associated groups) and
  comparing the read-backs; the JSOM handlers' suites skip. `npm run cleanup` removes leftovers (a
  plan unless `--apply`); `provision` and `cleanup` refuse a site outside `INTEGRATION_SITE_URL`
  unless `--allow-site` repeats it; `cleanup` never removes the web's own associated groups.
  Each attempt at a request the harness sends is bounded at 120 s (a throttled one still waits
  `Retry-After`), and on a CI runner the log and the JUnit reports pass through `redact()`. The runner no longer installs browser shims.
- Agent docs: `AGENTS.md` (and `CLAUDE.md`), `docs/TESTING.md`, the `spjs-testing` and
  `spjs-provision` skills in `.claude/skills`, checked by `npm run check-skills` in CI. The
  gulp-era `.vscode/launch.json`, `.vscode/tasks.json` and `settings.example.js` are gone.

- `package-lock.json` is regenerated: it still resolved every `@pnp/*` package to 3.17.0, so
  `npm ci` refused the tree and the runner ran a v3 `@pnp/nodejs` over v4 queryables.
- The local runner (`npm run provision`) signs in from `.env` (`.env.example`): locally with the
  session cookies of PP365's Playwright sign-in, in CI with an app certificate
  (`CI_TENANT`/`CI_CLIENT_ID`/`CI_CERT_BASE64`), ported from PP365's integration tests.
  `debug/provision.settings.ts` and the client-secret option are gone (SharePoint Online refuses
  app-only tokens from a secret), and so is the `@pnp/nodejs` devDependency (replaced by
  `@azure/msal-node`, `node-forge` and `dotenv`).
- The runner takes its handler list from `DefaultHandlerSort`, so a template's `Security` and
  `WebSettings` sections are applied again (its own list had left them out), and it checks
  `--handlers` against the engine.
- New runner flags: `--parameters` for `{parameter:...}` tokens, `--strict` (exit 1 when a
  handler logged an error), `--twice` (apply twice); `--site` defaults to `INTEGRATION_SITE_URL`.
- `--mode cloud` now matches PP365's `provisionCloudTemplateHubDependencies`: lists with hub
  configuration rows and Taxonomy are kept, `listContent` source lists are left out.
- `tests/fixtures/templates/smoke.json` for a first live run, and `tests/unit/runner.test.ts`
  for the runner and its sign-in, including the cookie pipeline through real PnPjs with a fake
  transport.

## [1.4.0] - 2026-09-21

### Changed

- Upgraded to PnPjs v4 (`@pnp/core`, `@pnp/logging`, `@pnp/queryable` and
  `@pnp/sp` 4.21.0). The `IWeb` passed to `WebProvisioner` must now come from a PnPjs v4
  `spfi(...)`; a v3 web is no longer accepted. Node 18.12 or newer is required.
- `@pnp/nodejs` moved to `devDependencies`. It was only ever used by the local debug runner,
  never by `src`, so consumers no longer pull `@azure/msal-node` and `@pnp/graph` into their
  bundles. Install it yourself when provisioning from Node.
- The default term store id behind `{sitecollectiontermstoreid}` is now read directly from
  `_api/v2.1/termstore` (the endpoint PnPjs v3's `TermStore` used), because PnPjs v4 has no
  `taxonomy` module. The resolved value is unchanged.
- Result shapes follow PnPjs v4, which resolves with the endpoint payload rather than a
  `{ data, <object> }` wrapper: `lists.add`, `fields.createFieldAsXml`, `navigation.add`,
  `files.addUsingPath` and `web.ensureUser`. This is internal to the handlers — the public API
  of `WebProvisioner` is unchanged.
- Batched `CustomActions` requests are awaited together with the batch itself
  (`Promise.all([execute(), ...requests])`), so a failing custom action is reported and no
  request promise is left unobserved when the batch itself fails.
- `@pnp/core`, `@pnp/logging`, `@pnp/queryable` and `@pnp/sp` are now
  `peerDependencies` (`^4.21.0`) instead of hard dependencies. The published typings expose
  `IWeb` in the `WebProvisioner` constructor, and PnPjs classes carry protected members, so two
  physically distinct copies of `@pnp/sp` are not assignable to each other. npm 7+ installs the
  peers automatically; with `legacy-peer-deps` or Yarn 1 install them explicitly.

### Fixed

- List fields are re-created against the right url again: the braced `ID` from the field schema
  XML (`{guid}`) is normalised before it is handed to `fields.getById`, which brace-wraps the
  value itself in v4. A braced id produced a double-braced url, so the delete/update fell
  through to the handler's catch and the field kept its generated title.
- The list lookups behind the `{listid:...}` tokens (`WebProvisioner` setup and the `Lists`
  handler) ask for `.top(5000)`. Without an explicit page size SharePoint returns the first 100
  lists only, and tokens for anything beyond that stayed unresolved.

### Added

- Unit tests (`npm test`) covering the PnPjs v4 result shapes, the term store url, the
  `CustomActions` batch lifecycle and the list lookup paging, run with the built-in Node test
  runner through `tsx` — no test framework added.
- `npm run typecheck:tests` type-checks the test files together with `src`. The test
  `tsconfig.json` previously inherited `include: ["src/**/*.ts"]`, so it never compiled a
  single test and emitted into `lib/`.

## [1.3.15] - 2026-09-04

### Added

- New tokens: `{sitecollectiontermstoreid}` (alias `{termstoreid}`) resolves to the id of the
  default term store, so managed metadata site columns can be provisioned from a template, and
  `{sitecollectionid}` now resolves to the actual site collection id. The site collection id,
  term store id and existing list ids are loaded once at setup, so tokens also resolve in
  handlers that run before `Lists`.
- Tokens can be written without a value part (`{siteid}` as well as `{siteid:...}`), and string
  values in list `DataRows` (plus `URL` field values) get token replacement, so seeded content
  can reference the site it is provisioned into.
- List permissions via `Security` on a list: break role inheritance and add role assignments.
  Principals accept `{associatedownergroupid}` / `{associatedmembergroupid}` /
  `{associatedvisitorgroupid}`, a principal id, a site group name or a user login name; role
  definitions resolve by localized name with a fallback to the well-known role types for the
  English names.
- SPFx extension registrations in `CustomActions` (`ClientSideComponentId`,
  `ClientSideComponentProperties`, `RegistrationId`, `RegistrationType`, `Sequence`).
- Node debug runner (`npm run provision`) for applying templates to a real site from the
  command line, with MSAL app-only auth (client secret or certificate). See `debug/README.md`.

### Changed

- `{sitecollectionid}` previously resolved to the web id. No known template used it.
- Unknown token keys are always left untouched, so JSON-like text and column placeholders in
  seeded content are never altered. A token that resolves to an empty value is now replaced
  with the empty value instead of being left in place.

## [1.3.14] - 2026-06-25

### Fixed

- Content type bindings are idempotent on re-apply: a content type that is already bound to
  the list is skipped instead of failing the binding step, and binding errors are logged with
  the underlying message.

## [1.3.13] - 2026-06-24

### Fixed

- Folder provisioning is idempotent on re-apply: existing folders are detected and only the
  missing ones are created, instead of failing when a folder already exists.

## [1.3.12] - 2026-06-09

### Added

- `UpdateExistingTerms` on a term set fills missing custom properties onto terms that already
  exist, so term properties can be added to term sets provisioned by an earlier version.

### Fixed

- Removing a content type that is still in use by existing items no longer aborts
  provisioning; the failure is logged and the run continues.

## [1.3.11] - 2026-06-08

### Added

- `Folders` on a list or library: provision a nested folder hierarchy.

## [1.3.10] - 2026-06-08

### Added

- `DataRows` on a list: seed list items from the template, upserted per row by `KeyColumn`
  with `UpdateBehavior` `Overwrite` or `Skip`. Field values are mapped by field type
  (lookup, user, taxonomy, URL, choice, boolean, date).

## [1.3.9] - 2026-06-04

### Added

- `Taxonomy` handler: provision term groups, term sets and terms with fixed ids, custom sort
  order and custom properties.

### Fixed

- Re-applying view formatters is resilient and surfaces errors instead of failing silently.

## [1.3.8] - 2026-05-13

### Fixed

- Content type field refs that already exist as field links are skipped instead of failing.

## [1.3.7] - 2026-05-12

### Fixed

- Additional view settings are applied when a view is created, not only when it is updated.
- The build output is included in the npm package.

## [1.3.6] - 2026-05-12

### Added

- Content types are created and updated with fixed, deterministic ids from the template.

[Unreleased]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.5.1...HEAD
[1.5.1]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.4.0...v1.5.0
[1.4.0]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.15...v1.4.0
[1.3.15]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.14...v1.3.15
[1.3.14]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.13...v1.3.14
[1.3.13]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.12...v1.3.13
[1.3.12]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.11...v1.3.12
[1.3.11]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.10...v1.3.11
[1.3.10]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.9...v1.3.10
[1.3.9]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.8...v1.3.9
[1.3.8]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.7...v1.3.8
[1.3.7]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.6...v1.3.7
[1.3.6]: https://github.com/Puzzlepart/sp-js-provisioning/compare/v1.3.5...v1.3.6
