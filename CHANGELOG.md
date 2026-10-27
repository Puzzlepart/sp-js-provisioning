# Changelog

All notable changes to this package are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
For releases older than 1.3.6, see the git history.

## [Unreleased]

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
  `test-results/unit.xml`), `test:coverage` (source-mapped, floors 73/78/74 for lines, branches
  and functions, which only rise), `check:pack` (the package holds `lib/index.js` and every handler) and
  `check:secrets` (no tenant URL, cookie, token or certificate in any committed or new text
  file, docs and specs included). eslint's `quotes` rule allows `"''"` (it conflicted with prettier), and
  `.prettierignore` keeps `lib/` out of the formatting check.
- `tests/fixtures/templates/failures.json`: a list field that cannot be created, a `Files` and a
  `PropertyBagEntries` entry, for checking that failures surface (`--strict` exits 1 for each).
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
