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

- Tokens such as `{sitename}` and `{parameter:Name}` are now replaced in list `Security` role
  assignment principals before they are resolved.
- Principal and role definition resolution moved from the `Lists` handler to
  `src/util/security.ts`, shared by list and web `Security`. Site group names with an apostrophe
  are now escaped in the lookup url.

### Development

Tooling only; nothing here ships in the package. Plan:
[integration testing spec](./docs/specs/integration-testing.md), Phase 0a.

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
