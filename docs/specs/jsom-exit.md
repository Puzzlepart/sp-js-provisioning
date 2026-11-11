# Spec: The JSOM Exit

Status: Implemented (2026-11-11: steps 1 to 5 done, green against the reserved site in cookie mode
and in the `Integration` workflow in certificate mode; Verified once PP365 has taken 1.6.0 and the
project setup and hub import below have run, see Verification Plan)

Owner: Remi Blom-Ohlsen

Related:

- [Integration testing](./integration-testing.md), D3 and Phase 4: this spec
- [Deterministic content type ids](./deterministic-content-type-ids.md) and its
  [ADR](../adr/jsom-explicit-content-type-ids.md), superseded by [CSOM over HTTP](../adr/csom-over-http.md)
- `src/util/csom.ts` (CSOM over HTTP, used since 1.5.0 for a site group's owner)
- PP365 `releases/1.15`: `PackageInstaller.ts` (hub import), `ApplyTemplate` (project setup)

## Problem

Four handlers reach SharePoint through JSOM, the browser's client object model
(`SP.ClientContext`, loaded by `spfx-jsom`):

- **ContentTypes** creates content types with the template's ids, adds and orders their field
  links, and updates them with `update(true)`. Every PP365 project setup runs it (two of PP365's
  own templates have `ContentTypes`), and every hub import of a hosting package (14 of 26 files).
- **Taxonomy** creates a term group, term sets and terms with the template's ids in the default
  term store (7 hosting files).
- **PropertyBagEntries** and the web parts of **Files** use a global `SP.ClientContext`, which an
  SPFx page does not have. Both refuse to run whenever `config.spfxContext` is set, which every
  consumer does: a template with either stops (`Files` throws) or loses the part. A hub import of
  `pp-testprosjekt` stops at `Files` in PP365 1.15.

What follows from it:

- None of the four runs in Node, so none has an integration test. The packages suite holds an
  exact list of what that costs (`tests/integration/packages.expected.ts`): every content type
  binding of every hosting package, and the columns, view fields and rows that depend on them.
- `Files` reads its source with `fetch(src, { credentials: 'include' })` and `response.text()`:
  a source on another origin with `Access-Control-Allow-Origin: *` (raw.githubusercontent.com,
  which `pp-testprosjekt` uses) is refused by the browser, and a binary file is corrupted.
- `PropertyBagEntries` writes only entries with `Overwrite: true` (`pp-testprosjekt`'s entry has
  none, so it would never be written) and replaces the web's indexed keys with the template's.
- `ContentTypes` sends `update(true)` on every run, for every content type, changed or not; on a
  hub that pushes each one to all its children again. One field ref to a column that is not there
  fails the whole section, and since 1.5.0 `SiteFields` goes on after a refused column.
- The package ships `spfx-jsom`, which loads SharePoint's JSOM scripts into the page.

## Goals

- No handler needs JSOM. All 14 run in Node and under an SPFx context alike, through the web's
  own PnPjs pipeline.
- `ContentTypes`, `Taxonomy`, `PropertyBagEntries` and `Files` have fixtures, read-backs and
  suites that run (not skip), and the packages suite's list shrinks to what a scratch web lacks.
- Re-applying is safe and quiet for the four as for the others since 1.5.0: an unchanged second
  apply writes nothing.
- `spfx-jsom` and `@types/sharepoint` leave the package.

## Non-Goals

- Nested terms, term labels in other languages, deleting or renaming terms, term sets or content
  types.
- The term store's REST API (`_api/v2.1/termstore`) for writes: it answers 403 to a browser
  session on tenants that enforce "OAuth only" (PP365 `SPDataAdapter.hasTermStorePermission`).
- Classic web parts in `Files` (decision J7).
- PP365's own use of `spfx-jsom` (its setup tasks), and PP365's adoption of 1.6.0.

## Constraints

- `src/` compiles with TypeScript 4.0.7 to ES5; `@pnp/*` stays a single peer copy.
- CSOM answers 200 with the error in the body. A request runs its actions in order and stops at
  the first that fails; what the actions before it did stays (there is no rollback).
- The term store is tenant-wide: a test that writes to it is seen by every site. Nothing in the
  runner or the suites may write to it without an explicit choice (J5).
- Custom script is blocked on modern sites (`DenyAddAndCustomizePages`): SharePoint refuses a
  property bag write there, whichever API asks (spike S5).
- An app-only (certificate) sign-in may be refused by the term store; the suite then skips and
  names it, as reachability does for the read.

## Spike (2026-11-10, on scratch webs of the reserved site, cookie sign-in)

| | Question | Answer |
|---|---|---|
| S1 | Does REST create a content type with the id it is given? | No. `POST _api/web/contenttypes` with `Id: { StringValue: '0x01005B1F…' }` (PnPjs 4.22 `contentTypes.add`) created `0x0100D8B5…`: SharePoint ignores the id and makes a child of Item with one of its own. The ADR's conclusion stands. |
| S2 | Does CSOM over HTTP? | Yes. `ContentTypes.Add` with a `ContentTypeCreationInformation` (`Id`, `Name`, `Description`, `Group`, no parent) through `_vti_bin/client.svc/ProcessQuery` returns the asked id, also for a child (`<id>01`). An id that exists answers "duplicate content type" (`Microsoft.SharePoint.SPException`); `GetById` of a missing id answers `IsNull: true`. |
| S3 | Field links, `Required`/`Hidden`, order, children? | `FieldLinks.Add` (a `FieldLinkCreationInformation` with the field), `SetProperty`, `Update(true)` and `FieldLinks.Reorder` all work in ProcessQuery, and `Update(true)` pushes links, flags and order to the child content type. `Reorder` puts the named links first, in that order, and the rest after. |
| S4 | The term store? | `TaxonomySession.GetTaxonomySession().GetDefaultSiteCollectionTermStore()` reads through ProcessQuery (default language, languages, online); `GetGroup`/`GetTermSet` of a missing id answer `IsNull: true`. Writes were not tried in the spike (J5). |
| S5 | The property bag? | `AllProperties.SetFieldValue` + `Web.Update()` through ProcessQuery is refused on the scratch web with `System.UnauthorizedAccessException` ("Access is denied"): custom script is blocked there. |

## Decisions

| | Decision | Rationale |
|---|---|---|
| J1 | Writes that REST cannot do go through CSOM over HTTP (`processQuery`): creating a content type with its id, field links and their order, `Update(true)`, the term store, the property bag. Reads use REST where it has them (content types and their links, the property bag). | S1 to S5. ProcessQuery is what JSOM itself sends, so behaviour and permissions are the same as today; it uses the web's sign-in (SPFx in a browser, cookies or a bearer token in Node), as the group owner fix has since 1.5.0. |
| J2 | One small request builder in `src/util/csom.ts` (object paths, actions, typed parameters, the response by action id, an error with SharePoint's type name). No new dependency. | The requests are few and regular; the XML stays in one place and is pinned by unit and contract tests. |
| J3 | Each handler reads first and writes only what differs: a content type's name, description and group, its missing links, changed flags and a changed order; missing groups, term sets and terms, missing custom properties (`UpdateExistingTerms`) and a changed sort order; property bag entries whose value differs. An unchanged re-apply sends no write. | As Lists, SiteFields and Navigation since 1.5.0. `Update(true)` on an unchanged content type pushed it to all its children on every run. |
| J4 | A content type field ref whose column is not on the web is logged as an error and left out; the content type gets its other links. A content type that cannot be created or updated still throws. | `SiteFields` logs a refused column and goes on (1.5.0); the content type that names it then stopped the whole template. |
| J5 | The term store is written only on purpose. The runner applies `Taxonomy` only with `--taxonomy` (or when `--handlers` names it); the packages suite never applies it (the hosting packages write to the tenant's own `Prosjektportalen` group). The Taxonomy suite uses one term group of its own, `SPJS Itest` with fixed ids, created by the first run and left in place: nothing is deleted from the term store (the user, 2026-11-10). | The site guard protects sites, not the term store, which every site of the tenant shares. |
| J6 | `PropertyBagEntries` follows PnP: an entry is written when its key is missing, and overwritten only with `Overwrite: true`. `Indexed` adds the key to the web's indexed keys, keeping the ones that are there. A write SharePoint refuses is logged as an error that names the cause (custom script is blocked) and the template goes on. | Today only `Overwrite: true` entries are written and the indexed keys are replaced. On a modern site the write is refused, and stopping there would leave `Navigation` and `WebSettings` unapplied. |
| J7 | `Files` uploads and sets properties through REST. A source on the web's own origin is read through the web's pipeline, any other with a plain `fetch` without credentials, both as bytes. A file that fails is logged as an error and the next is tried. `WebParts`, `RemoveExistingWebParts` and a web part's `ListView` are no longer supported: a file that has them is uploaded and an error says so (the user had no preference, 2026-11-10). | No template of PP365 or prosjektportalen-hosting uses them, they need a classic page and custom script, and they cannot be tried on the reserved site. Kept, they would be a port nobody can verify. |
| J8 | `spfx-jsom`, `@types/sharepoint`, the Node stub and its `paths` mappings go. `config.spfxContext` no longer turns `Files` and `PropertyBagEntries` off. | Nothing needs them. |
| J9 | Released as 1.6.0. PP365 takes it after 1.15.0, with its own follow-ups (Rollout Notes). | New capability and a removed one (J7). |

## Proposed Approach

### `src/util/csom.ts`

`CsomRequest` builds one ProcessQuery request:

- object paths: `current()` (`SP.RequestContext.Current`), `property(parent, name)`,
  `method(parent, name, ...parameters)`, `staticMethod(typeId, name, ...parameters)`;
- actions: `instantiate(path)`, `query(path, properties)`, `set(path, name, value)`,
  `call(path, name, ...parameters)`;
- parameters: strings, booleans, integers, `null`, `guid(value)`, `ref(path)`, arrays of strings,
  and `valueObject(typeId, properties)` (the creation informations); every text is XML-escaped.

`execute(web, request)` posts it and returns the results by action id; an `ErrorInfo` becomes a
`CsomError` with SharePoint's message, type name and code. `processQuery` and
`setGroupOwnerToGroupRequest` stay.

### ContentTypes

1. Read the web's content types with their links (REST, as today) and the internal names of the
   web's site columns.
2. In id order, for each: create it when its id is not there (S2); set `Name`, `Description` and
   `Group` where they differ; add the links that are missing (a column that is not on the web:
   J4), set `Required` and `Hidden` where the template has them and they differ; `Update(true)`
   once, only when something was set. A new content type's inherited links are read before its
   field refs are compared.
3. `Reorder` and `Update(true)` only when the named links are not already first, in the
   template's order.
4. The context's content type cache is updated as today.

### Taxonomy

1. One read: the default term store's language, and for the group, every term set and every term
   whether it is there (`IsNull`), a term set's `CustomSortOrder`, and a term's `CustomProperties`
   when `UpdateExistingTerms` is set.
2. One write per term set that needs one: the group and the term set where missing (name, id,
   the store's default language, `Description`, `IsOpenForTermCreation`), missing terms with
   their custom properties, missing custom properties on existing terms (`UpdateExistingTerms`,
   never overwriting), the sort order when it differs, then `CommitAll`.
3. A term store that is not there, or a refused write, throws as today (`Taxonomy` runs first, so
   an import fails before anything else is made).

### PropertyBagEntries

Read the bag (REST), write the entries J6 selects with `SetFieldValue`, the merged
`vti_indexedpropertykeys` and `Web.Update()` in one request. Nothing to write: no request.

### Files

Per file: read the source (J7), `files.addUsingPath` with `Overwrite` (default `true`),
`Properties` on the file's item. `Overwrite: false` on a file that exists: left as it is.

### Lists, runner, tests

- `Lists` no longer initialises JSOM (it never used it).
- Runner: no handler is "browser-only"; `Taxonomy` is opt-in (J5). `verify` and `inspect` are
  unchanged (they already read all four).
- Fixtures and suites: `contenttypes.json` (+ `.b.json`), `taxonomy.json`,
  `propertybagentries.json`, `files.json`; the four suites apply twice, read back, run `verify`,
  and assert that the second run sends no write. `all.node.json` becomes `all.json` with the
  content types (not `Taxonomy`, J5; not files and property bag entries either, as it turned out:
  `files.json` needs a local server and a source in the web, and the property bag needs custom
  script, so they stay with their suites, like `hooks.json`). `packages.expected.ts` is
  regenerated: what is left is what a scratch web lacks (PP365's own columns and content types)
  and what the reserved site refuses (the property bag).
- Unit tests for the request builder (every parameter kind, escaping, the response, an error) and
  each handler with a fake transport; contract tests pin each request on the wire.

## Steps

1. This spec and the request builder.
2. ContentTypes.
3. Taxonomy.
4. PropertyBagEntries and Files.
5. `spfx-jsom` out, the runner, `all.json`, the packages list, docs, 1.6.0's changelog.

One commit per step; the suites run against the reserved site at each.

Status, 2026-11-11: all five steps are implemented and documented. No handler uses JSOM, and
`spfx-jsom` and `@types/sharepoint` are gone. The runner applies every handler but Taxonomy, which
it applies only with `--taxonomy` (J5). Against the reserved site, cookie mode:
`npm run test:integration -- all` 66 tests, 64 pass, 2 skipped (both need custom script), 0 fail,
14 minutes, `everything` with the content types included; `-- packages` 9 of 9 against the
regenerated list, 9 minutes; each of the four suites green at least twice. Offline: 413 tests,
coverage 94.1/87.2/88.7 (floors raised to 90/85/86).

What the packages showed once the four ran:

- The content types of the packages are created, and the lists bind them; what is left is PP365's
  own: site columns a content type or list names (left out with an error each, and the view fields
  and rows that need them) and content types a list binds. A PP365 hub has both.
- `pp-veiprosjekt`'s `Prosjekt (Vei)` is a child of PP365's `Prosjekt`
  (`0x0100805E9E4FEAAB4F0EABAB2600D30DB70C`): on a scratch web SharePoint refuses it ("the parent
  content type does not exist") and the template stops, as it should (J4). The suite makes the
  parent first, as it makes the menu's targets.
- `pp-testprosjekt`'s property bag entries are refused (custom script blocked), one error per run.
- Files with a raw.githubusercontent.com source upload in Node.

The first `Integration` run after the exit (2026-11-11, run 42972561472, certificate mode,
`suites=all`, `packages=testprosjekt`): 75 tests, 71 pass, 2 skipped (custom script), 2 fail.
ContentTypes, Taxonomy (both tests: the CI app may read and write the term store), the property
bag refusal, `everything` and the nine package cases pass in certificate mode. The two Files tests
failed: SharePoint answers an app-only token on a file's own URL with 401, so a source on the web's
origin could not be read that way (with cookies and under SPFx it can). `Files` now reads such a
source through REST (`getFileByUrl`, then `$value`), and the URL itself only for what is not a file
in a library (a `_layouts` resource).

The second run (2026-11-11, run 42977511130, on that fix): 75 tests, 73 pass, 2 skipped (custom
script), 0 fail, 26 minutes, every suite of the four and the nine package cases green in
certificate mode; the report check passed, and the public log holds neither the site's host nor an
e-mail address.

What the Files and PropertyBagEntries runs showed:

- A web's document library is `Delte dokumenter` on a Norwegian site, and SharePoint answers a
  folder that is not there with 403 ("Ingen tilgang"), not 404: the Files fixture uses
  `SiteAssets`, which has that URL in every language.
- A text file with æøå and CRLF and a binary file come through byte for byte, from the web (read
  through its pipeline) and from another origin (a local server, no credentials).
- The property bag write is refused on the reserved site, as the spike found:
  `System.UnauthorizedAccessException` ("Access is denied", `E_ACCESSDENIED`); the handler logs it
  with the cause and the template goes on. Where custom script is allowed is proven by the unit
  tests only.

What the Taxonomy runs showed (the signed-in user, cookies):

- The first run created the group `SPJS Itest`, the term set and its four terms with the
  template's ids, custom properties, description and sort order, in three requests (the read, the
  group, the term set); the second only read. The user may write to the term store.
- `_api/v2.1/termstore` returns a term's custom properties with `$select=properties` and a term
  set's sort order with `$select=customSortOrder` (neither without a `$select`): `readTaxonomy` and
  `verify` use them.

## Acceptance Criteria

- `git grep -n "spfx-jsom\|SP\.ClientContext\|executeQueryAsync" -- src` finds nothing, and
  `package.json` has neither `spfx-jsom` nor `@types/sharepoint`.
- `npm run test:integration -- all` runs the ContentTypes, Taxonomy, PropertyBagEntries and Files
  suites (none skips as "JSOM handler"), each green twice in a row, the second apply of each
  writing nothing.
- A content type is created with the template's id, on a fresh web and as a child; a changed
  name, a new link, a changed `Required` and a changed order each reach SharePoint, and a child
  content type follows.
- The term group `SPJS Itest` holds the fixture's term set and terms with the template's ids,
  custom properties and sort order; a second apply leaves the store untouched.
- On a web that blocks custom script, `PropertyBagEntries` logs one error naming the cause and the
  template goes on; where it is allowed, the entry and its indexed key are there.
- `Files` uploads a text file and a binary file byte for byte, from the web's origin and from
  another, and sets its properties.
- The packages suite's list has no entry caused by a handler that could not run.
- Under an SPFx context (PP365), a template with `Files` and `PropertyBagEntries` no longer stops.

## Verification Plan

- Each step: `npm run build`, `npm test`, `npm run typecheck:tests`, `npm run lint:check`,
  `npm run test:coverage`, `npm run check:schema`, then the step's suite against the reserved
  site and `npm run test:integration -- all` at the end.
- `npm run test:integration -- packages` with the regenerated list.
- ~~The `Integration` workflow once in certificate mode (`suites=all`): which of the four an app
  may do is recorded here.~~ Done 2026-11-11 (run 42977511130): all four, the term store included.
- In PP365, after its bump: a project setup (ContentTypes), a hub import of `pp-testprosjekt`
  (all four) and the same import again.

## Open Questions

- ~~May PP365's CI app write to the term store, or does the Taxonomy suite skip in certificate
  mode?~~ It may: both Taxonomy tests passed in certificate mode (run 42972561472).
- ~~Is the signed-in test user a term store administrator?~~ It may write to the term store: the
  first Taxonomy run created the group (2026-11-10).

## Rollout Notes

- 1.6.0 changes what consumers see: `Files` and `PropertyBagEntries` run under SPFx; a property
  bag entry without `Overwrite` is written when missing; a refused property bag write and a
  failing file are logged and no longer stop the template; an unchanged content type is not
  pushed to its children again; classic web parts in `Files` are refused with an error.
- PP365 follow-ups when it takes 1.6.0: `PackageInstaller` can stop stripping what it stripped
  for JSOM; a `pp-testprosjekt` hub import completes, so the package's README (what the import
  stops at) and PP365's guide change; `spfx-jsom` stays for PP365's own tasks.
- prosjektportalen-hosting: a package may use `Files` and `PropertyBagEntries` on a hub only
  with `minPPVersion` of the PP365 version that ships 1.6.0.
