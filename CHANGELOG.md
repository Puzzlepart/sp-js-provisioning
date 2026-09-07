# Changelog

All notable changes to this package are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
For releases older than 1.3.6, see the git history.

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
