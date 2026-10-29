# Spec: SiteFields re-apply, same-template lookups and failing columns

Status: Verified (2026-10-29)

Owner: Remi Blom-Ohlsen

Related: [Lists re-apply](./lists-re-apply.md) (the list-field rules this spec applies to site
columns), [integration testing](./integration-testing.md) (the pinned `{listid:}` case),
navigation-re-apply.md (written alongside, same release). Ships in 1.5.0 with Lists and Security.

## Problem

`SiteFields` creates a site column from field XML when its internal name is missing and otherwise
sends the template's XML as the column's new `SchemaXml`. It was written for the first run. PP365
applies it again when an admin re-runs project setup, and on every hub package import, update and
cloud publish (`PackageInstaller._provisionHub`, `PackageInstaller.ts:535-571`, 578-690), through
one `WebProvisioner` per setup. What that did before this spec, from the closing analysis
(2026-10-29) as corrected by the tenant probe (Evidence below):

1. **`{listid:}` naming a list the same template creates is stored as the literal token, silently**
   (pinned `{ todo }` in `sitefields.test.ts`). `SiteFields` (sort 0) runs before `Lists` (3), and
   `{listid:}` resolves only for lists that exist when `applyTemplate` starts (`loadContext`).
   SharePoint accepts a lookup site column with a literal `List` (probe SF6), the run logs nothing,
   and every copy made of the column in the same run is broken: ContentTypes' links and pushes, and
   list FieldRefs, which SharePoint refuses at list level ("Value does not fall within the expected
   range", SF6). Only a second run pointed the site column itself at the list. `TokenHelper` leaves
   an unknown token in place on purpose (JSON, `{Title}`), which is also what made the miss silent.
   No consumer template has a lookup site column today; README promises `{listid:}` in field XML.
2. **Every re-apply rewrote every existing column with the template's XML.** No read of the stored
   schema, no comparison, no type or target rule:
   - every column was written on every run (a GET for `FieldTypeKind` and a MERGE; each MERGE
     raises the column's `Version` by 2, SF1);
   - SharePoint replaces a site column's schema wholesale (SF2): what the XML leaves out is
     dropped or reset, except `Group` and `SourceID`: `Indexed` back to false, `CHOICES` emptied,
     `Default`, `CustomFormatter`, `Description` removed. A hub admin's index, formatter or
     description on a package column was lost on every import;
   - a type change or a new lookup target was applied as given;
   - an unresolved token was written over a working lookup target: a SchemaXml MERGE re-targets a
     lookup (SF6), so `{listid:X}` after X was renamed broke the column.

   Two claims of the analysis are wrong and are not problems: a taxonomy column does **not** lose
   its `TextField` (the vei hub's `GtVeiSubject` and `GtVeiTopic`), and an unresolved
   `{sitecollectiontermstoreid}` cannot damage a stored `SspId` through this call. SharePoint
   ignores a taxonomy column's whole `Customization` in a SchemaXml MERGE (SF3). The hosting docs'
   statement that a literal term store token makes the taxonomy create fail
   (`prosjektportalen-hosting/docs/CLOUD-TEMPLATE-COMPATIBILITY.md:111-114`) is also wrong: the
   create is accepted and stores the empty GUID as `SspId` (SF3a).
3. **One failing column aborted the whole template.** `processSiteField` rethrew, `WebProvisioner`
   wrapped it in a `ProvisioningError`, and nothing after SiteFields (the first REST handler) was
   applied. PP365 works around it by stripping "blocked" columns before `applyTemplate`
   (`PackageInstaller.ts:553-559`, `CompatibilityService.ts:190-199`).
4. **Identity was by internal name only, and only against `web.fields`.** A column with the same
   id and another name took the create path and was refused with SharePoint's cryptic "The local
   device name is already in use (0x80070055)" (SF5), aborting the run. On a subweb `web.fields`
   lists only that web's own columns (SF4): a template column a parent web has was refused as a
   duplicate ("Fant et duplisert feltnavn", SF4b).
5. Smaller: no `scope_ended` on success; a list FieldRef that names a column found neither on the
   list nor on the web did nothing and logged nothing.

## Goals

- A second apply of an unchanged template writes no site column (`Version` unchanged).
- Template changes reach existing site columns where they are safe: what the template says is
  merged onto what the column holds; a narrowing type change is a Warning and the column is left as
  it is; a new target is a Warning and the stored target is kept, unless the column was made in
  this setup (the rest of the template is still merged).
- A lookup site column to a list of the same template points at that list on the first run, and
  so do the copies made of it in that run.
- A known token that cannot be resolved is an Error naming the column and the token, never written
  over a target the column has; an old column holding such a literal heals on the next run.
- A column that fails is an Error naming it and the operation; the others and the handlers after
  SiteFields still run.
- The pinned SiteFields `{ todo }` flips to a plain test.

## Non-Goals

- **No push.** Site column changes do not reach the column's existing list copies, and the engine
  does not push them (CSOM `UpdateAndPushChanges`): a SchemaXml MERGE does not push (SF1), and
  ContentTypes' every-run `update(true)` does not re-sync list field schema either (SF7). A later
  opt-in is possible, but a push would carry narrowing type changes and re-targets into lists that
  hold values (the Lists probe saw silent value loss).
- **No PnP Framework parity** for lookups (creating them without `List` and fixing them after
  Lists). The engine pre-creates the target list instead (S1).
- **`{listviewid:}` before Lists is documented, not fixed** (S1d): `loadContext` does not load
  views, so it resolves before Lists only for views an earlier `applyTemplate` through the same
  provisioner processed (`context.listViews` lives on the provisioner).
- Repairing list copies that earlier engine runs made from a broken column. None are known in
  consumers; they are fixed by hand.
- Taxonomy re-targeting to an existing term set on the tenant (the tenant has no term set; covered
  by unit tests and SF3g).
- ContentTypes (JSOM, browser-only) and PP365 changes (Rollout Notes).

## Constraints

- `src/` stays TypeScript 4.0 / ES5; PnPjs 4's `select()` mutates its queryable, so every read uses
  a fresh one.
- Log and carry on, as Lists does: nothing new may stop the run; `--strict` and `assertClean` still
  see every Error.
- Handler order is a documented contract (ARCHITECTURE "Current order"); the S1 pre-pass is the one
  stated exception.
- Only calls the probe saw SharePoint accept: `createFieldAsXml`, a SchemaXml MERGE, a
  `TaxonomyField` property MERGE, `lists.add`, reads of `fields` and `availablefields`.
- PP365 bundles the engine at build time; templates written for 1.4.0 must keep working on the
  first run.

## Proposed Approach

### S1: lists a site column names, created before SiteFields

`WebProvisioner.createListsForSiteFields` runs inside the handler chain, right before `SiteFields`,
so a handler before it (`Taxonomy`) that fails leaves no list behind. It runs only when
`SiteFields` and `Lists` both run and `SiteFields` comes first (the `handlers` filter and a custom
`handlerSort` are respected). It collects the `{listid:X}` tokens still unresolved in `SiteFields`
(TokenHelper `unresolvedTokens`, S1b) and creates the lists among them that the template's `Lists`
declares by `Title`, through `Lists.createList`, the create path
`processList` uses, factored out: `lists.add` with `Description`, `Template`,
`ContentTypesEnabled` and `AdditionalSettings`, then `context.lists[X]` and
`context.createdLists[X]`. Each creation is logged at Info.

- `Lists` then finds the list existing. `context.preCreatedLists` (reset per `applyTemplate`) tells
  `processList` not to send the settings again; bindings, fields, views, rows and security follow
  as for any list, and the list counts as made in this setup (`RemoveExistingViews`,
  `RemoveExistingContentTypes`).
- A list that cannot be created is an Error naming it; the run goes on as without the step (`Lists`
  creates it; SiteFields logs the token). A title the template's `Lists` declares twice is created
  once, from its first entry.
- When a handler between `SiteFields` and `Lists` (`ContentTypes`, `Features`, `Security`) fails,
  the run fails as before, and a Warning names each list the step created that `Lists` did not
  reach: it exists without its fields, views and content types, and a later run (a new
  provisioner) no longer counts it as made in this setup.
- When `loadContext` could not read the web's lists, `context.lists` is stale and a list that exists
  would look missing: the step creates nothing and logs a Warning.
- A token naming a list the template does not declare is left to S1b.

### S1b: unresolved tokens are errors

`TokenHelper.unresolvedTokens(text)` returns the known tokens (`listid`, `listviewid`, `webid`,
`siteid`, `sitename`, `sitecollectionid`, `sitecollectiontermstoreid`, `termstoreid`, `parameter`)
that do not resolve now, each once, as written. Unknown keys are not tokens and are not reported.
SiteFields logs one Error per column naming the column and the tokens. The column is still created
with the literal, as before (ContentTypes and list FieldRefs still find it by name). On a column
the web already has, a target holding the literal never overwrites the stored value (S2).

### S1c: old literal targets heal

On a column the web has, a stored target (`List`, `WebId`, `SspId`, `TermSetId`) that is not a
GUID (a literal an earlier run stored), or the empty GUID as a taxonomy column's `SspId` or
`TermSetId` (what SharePoint stores for a literal term store id, SF3a), counts as unset and follows
the template:

- a lookup's `List` / `WebId` through the SchemaXml MERGE (SharePoint re-targets it, SF6);
- a taxonomy column's `SspId` / `TermSetId` through the `TaxonomyField` property MERGE
  (`update({ SspId }, 'SP.Taxonomy.TaxonomyField')`, likewise `TermSetId`), the only call that sets
  them (SF3g, SF3h).

A taxonomy target the template gives as something other than a GUID that S1b does not report (a
token in another case, `{SiteCollectionTermStoreId}`, or one TokenHelper does not know) cannot be
set: SharePoint would store the empty GUID (SF3a). SiteFields logs an Error naming the column, the
target and the value, on create as on update, and keeps the stored value; it never reports such a
column as healed.

### S2: the list-field rules on site columns, without a push

One read of the web's columns (`Id`, `InternalName`, `TypeAsString`, `SchemaXml`), with the columns
available to the web (`Id`, `InternalName`, S5), both fresh queryables, without `top()` (SF4: the
collection is not paged and `$top` truncates it silently). For a column the web has:

- `mergeFieldXml` (shared with Lists, `src/util/fieldxml.ts`): the template's attributes win, the
  stored identity and storage are kept (`ID`, `Name`, `StaticName`, `SourceID`, ...), attributes and
  children the template leaves out are kept, `Version` is dropped, and the type is written as
  SharePoint stores it (`Type="Lookup" Mult="TRUE"` as `LookupMulti`; this moved into
  `mergeFieldXml` from Lists).
- Nothing is written when `sameFieldXml(merged, stored)`; otherwise one
  `update({ SchemaXml }, 'SP.Field')` (the type argument saves PnPjs' extra GET).
- `typeChange`: a widening change (`Text` to `Note`, `Lookup` to `LookupMulti`, `User` to
  `UserMulti`) is applied; any other is a Warning naming both types, and the column is left as it
  is.
- Targets, per target: the template leaves it out or says the same: nothing; the template's value
  still holds an unresolved token: the stored value stays (Error per S1b); stored unset (S1c):
  follows; the column was made in this setup (`context.createdSiteFields`, across the provisioner's
  `applyTemplate` calls, like `createdLists`): follows; otherwise a Warning, and the stored target
  stays (list copies made before point at it).
- **Taxonomy columns** (`TaxonomyFieldType`, `TaxonomyFieldTypeMulti`): the merge keeps the stored
  `Customization` whole and the attributes SharePoint manages (`List` = TaxonomyHiddenList,
  `WebId`, `ShowField`, which it sets to the site's language), so an unchanged re-import writes
  nothing; a SchemaXml MERGE would ignore the `Customization` anyway (SF3). `SspId` and `TermSetId`
  follow the target rule above through the `TaxonomyField` property MERGE.
- Never a push: list copies keep their schema (SF1, SF7). Template changes to a site column reach
  only list copies made after them.

### S3: a failing column is logged, the handler resolves

Each column is processed in its own try/catch. An XML that cannot be read: "Failed to read a site
column's XML, it is skipped"; a refused create or update: "Failed to create|update site column
<Name> (<DisplayName>): <SharePoint's message>". The handler logs `scope_ended` and resolves. When
the web's columns cannot be read at all, one Error names the template's columns and nothing is
created or changed (it cannot tell create from update). `Lists.processFieldRef` logs an Error when a
FieldRef names a column found neither on the list nor on the web (it did nothing, silently, before
and after 1ad8181).

### S4: identity by internal name

As before and as PP365's CompatibilityService assumes. Same name with another id: Warning, updated
in place keeping its stored id (SharePoint ignores `ID` in a MERGE anyway, SF5). Same id under
another name: Error, not created (SharePoint would refuse it with 0x80070055, SF5).

### S5: subwebs

A template column found only among the web's available columns (a parent web has it) is neither
created nor changed: a Warning says a parent web has it. The template changes no other web than the
one it is applied to. A parent web's column with the template column's id under another name is an
Error, not created (S4).

## Decisions (2026-10-29)

| # | Question | Decision |
|---|---|---|
| S1 | `{listid:X}` naming a list the same template creates | Create X before SiteFields: only lists an unresolved `{listid:}` in a SiteFields XML names and the template's `Lists` declares (by Title); the create path of `Lists.processList`, factored out; `context.lists` and `createdLists` set; Info; a failed create is an Error and the run continues; only when SiteFields and Lists both run and SiteFields sorts first |
| S1b | A known token still unresolved in a site column | The column is still created with the literal; an Error names the column and the token |
| S1c | A stored non-GUID target (an old literal) | Counts as unset and follows the template, so old broken columns heal |
| S1d | `{listviewid:}` before Lists | Documented as not resolvable before Lists; no code change |
| S2 | Re-apply of an existing site column | The list-field rules without a push: one read; `mergeFieldXml`; no write when `sameFieldXml`; widening applied, other type changes a Warning; a stored GUID target the template changes: Warning and keep; a stored non-GUID target or a column created in this run follows; an unresolved template target never overwrites; `update(..., 'SP.Field')`; never push |
| S3 | A failing site column | Error naming the column and the operation, carry on, resolve; `scope_ended` on success; `processFieldRef` Error for a column found nowhere |
| S4 | Identity | By internal name. Same name + other id: Warning, stored id kept. Same id + other name: Error, not created |
| S5 | Subwebs (after the probe) | Identity also from `web.availablefields`; a column only in a parent web is neither created nor changed, with a Warning |

Refinements from the probe (no new decision needed): the taxonomy merge keeps the stored
`Customization` and the SharePoint-managed `List`, `WebId` and `ShowField`; the empty GUID is unset
for `SspId` and `TermSetId`; healing them uses the `TaxonomyField` property MERGE; no `top()` on
fields; same id + other name is detected from the engine's own read.

Defaults set while implementing (open to change, see Open Questions):

- The empty GUID counts as unset only for `SspId` and `TermSetId`, not for a lookup's `List` or
  `WebId`.
- A parent web's column with the template column's id under another name is an Error (S4 applied to
  S5).
- When the web's columns cannot be read, nothing is created or changed, with one Error.
- A list created before SiteFields is not sent its settings again by `Lists`
  (`context.preCreatedLists`).
- Only targets are protected from unresolved tokens on update; another attribute with an unresolved
  token (`{sitename}` in a `Description`) is written as it stands, as on create, with the Error.

## Evidence: tenant probe (2026-10-29)

A throwaway probe on five scratch webs below the reserved site (not in the repository; all removed,
no leftovers), nb-NO site, cookie session:

| Case | SharePoint | Consequence |
|---|---|---|
| SF1: SchemaXml MERGE on site columns (Choice, Text, Lookup) with list copies made the FieldRef way and an item | accepted; site columns changed (`Version` 1 to 3); list copies, their `Version` and the item unchanged | No push; each MERGE raises `Version` by 2, so "Version unchanged" proves no write |
| SF2: MERGE with only `Type`, `ID`, `Name`, `StaticName`, `DisplayName` | accepted; `Description`, `CustomFormatter`, `Default`, `Format`, `FillInChoice`, `ShowInNewForm` dropped, `Indexed` reset to false, `CHOICES` emptied; `Group`, `SourceID` kept | `mergeFieldXml` is required |
| SF3a: taxonomy create with a literal `SspId` | accepted; `SspId` stored as `00000000-0000-0000-0000-000000000000`; SharePoint adds `TextField`, `List` (TaxonomyHiddenList), `WebId`, and sets `ShowField` Term1033 to Term1044 | Empty GUID = unset; keep the managed attributes |
| SF3b-f: taxonomy MERGE without `TextField`, with a literal `SspId`, with another `TermSetId` or `SspId` | accepted; `DisplayName` applied; `SspId`, `TermSetId`, `TextField` unchanged whatever is sent | `TextField` is not lost; Customization is ignored in a MERGE |
| SF3g/h: `update({ TermSetId } \| { SspId }, 'SP.Taxonomy.TaxonomyField')` | accepted and applied, `TextField` kept | The only way to heal or re-target a taxonomy column |
| SF4: `_api/web/fields` without `$top` (root web, 541 fields), with `$top=100` | all 541, no `nextLink`; `$top=100` gives 100, no `nextLink` | No `top()`; on a subweb `fields` has only its own 11, `availablefields` 552 |
| SF4b: create on a subweb with a root column's name / id | refused: "duplicate field name" / 0x80070055 | S5 |
| SF5a: MERGE carrying another `ID` under the same name | accepted; stored id kept | S4 Warning is the engine's own |
| SF5b: create with an existing id and another name | refused, 500, code -2147024811, "The local device name is already in use (0x80070055)", in English | Detect from the read; Error, no create |
| SF6: lookup site column with a literal `List`; FieldRef copy; MERGE with the GUID | create accepted with the literal; copy refused (400 "Value does not fall within the expected range"); MERGE re-targets; later copies correct | S1, S1b, S1c hold for lookups |
| SF7: ContentType `Update(true)` (REST and CSOM, with a field link touched) after site column MERGEs | accepted; list field `Title`, `CHOICES`, `Version` and the item unchanged | ContentTypes is no hidden push |

Side findings for a future REST ContentTypes handler (out of scope): REST `contentTypes.add`
ignores the requested id; the REST field link add was refused; CSOM `FieldLinks.Add` works.

## Verification Result (2026-10-29)

- Offline: `npm test` 258 tests, all pass (35 of them new for this spec: `sitefields.test.ts` 19,
  `webprovisioner.test.ts` 5, `fieldxml.test.ts` 5, `tokenhelper.test.ts` 3, `lists.test.ts` 3; the
  rest includes the Navigation spec's, written in the same working tree); `build`,
  `typecheck:tests`, eslint and prettier on the changed files, `check:secrets` green;
  `test:coverage` 80.14 / 83.68 / 79.95 measured alone, 80.36 / 84.02 / 80.21 with both specs (floors
  raised from 73 / 78 / 74 to 79 / 83 / 79; `sitefields.ts` 98.6 % lines).
  Mutating the taxonomy keep list, the token rule, the empty-GUID rule, the made-in-this-setup rule
  or the parent-web check each turns at least one unit test red.
- `npm run test:integration -- SiteFields` against the reserved site: 7 tests, all pass. Then
  `npm run test:integration -- SiteFields Lists`: reachability 6 pass; the suites 25 tests (Lists 11,
  Lists (DataRows) 7, SiteFields 7), 25 pass, 0 fail, 0 `todo`, 0 skipped; about 4 minutes; no
  throttling line; nothing left behind. A Navigation run of the other spec was active on the same
  site at the time; neither disturbed the other.
- What the tenant showed: the second apply of the eleven-type fixture wrote no column (versions
  equal, each column logged unchanged); a `Description` set by hand (a property MERGE, `Version` 1
  to 2) survived a re-apply unwritten; Note to Text gave exactly one Warning and left the column;
  the same-template lookup and its FieldRef copy pointed at the list on the first run, the list kept
  its template description (created before SiteFields, its settings not sent again), and the second
  run wrote nothing; a vei-style taxonomy column (`Description=""`, `ShowField="Term1033"`, no
  `TextField`, `SspId` as a token, a term set that does not exist) read back with the real term
  store id, SharePoint's own `TextField`, and the same `Version` after a second apply; SharePoint
  refused the Calculated column whose formula names a missing column, the run logged one Error and
  created the Text column after it.
- PP365 (throwaway script on a scratch web, removed afterwards): Standardmal and the
  EnkelVenstremeny extension through one provisioner, then again with a new provisioner. Run 1
  created the 5 site columns (5 `createfieldasxml`, no MERGE); run 2 wrote none (no POST or MERGE
  on `_api/web/fields`, each column logged unchanged, `Version` 1), so the Calculated
  `GtRiskFactor` and `GtRiskFactorPostAction` compare as the same. No SiteFields Error or Warning.
- Review (2026-10-29), five reviewers and an adversarial check of each finding, fixed with a unit
  test that fails without the fix: the pre-pass ran before `Taxonomy` and could leave an empty list
  when an earlier handler failed (now right before `SiteFields`, with a Warning for a list `Lists`
  did not reach); a title declared twice was created twice (a false Error); a taxonomy target that
  is no GUID was reported as healed (now an Error, nothing sent). Afterwards `npm test` 265 tests,
  all pass; `test:coverage` 80.50 / 84.07 / 80.41.
- Full integration run after the fixes (`npm run test:integration`): reachability 6 of 6; the
  suites 52 tests, 51 pass, 0 `todo`, 1 skipped (custom script blocked), 0 fail; nothing left
  behind.
- Not run: the vei hub template imported twice through PP365's PackageInstaller.

## Acceptance Criteria

- `sitefields.test.ts` (integration), against scratch webs:
  - the eleven-type fixture applied twice: same ids, and the second apply writes nothing (every
    column's `Version` unchanged, each logged as unchanged);
  - B over A changes only the Text column (only its `Version` rises), B again writes nothing, A
    again restores it;
  - a `Description` set by hand on a column whose template has none survives a re-apply, unwritten;
  - a narrowing type change (Note to Text) gives exactly one Warning and leaves the type;
  - the former `{ todo }`: a lookup to a list of the same template points at it on the first run, a
    FieldRef copy made in the same run too, the list has its template settings, a clean run, and
    the second run writes nothing;
  - a taxonomy site column (`SspId` `{sitecollectiontermstoreid}`, any `TermSetId`, no `TextField`)
    is unchanged by a second apply (`Version`, `SspId`, `TermSetId`, `TextField`);
  - `failures.json` with `--handlers SiteFields`: SharePoint refuses a Calculated column whose
    formula names a missing column; one Error names it, the run does not stop, and the next column
    is created.
- Unit tests for each rule (`sitefields.test.ts`, `webprovisioner.test.ts`, `tokenhelper.test.ts`,
  `lists.test.ts`, `fieldxml.test.ts`).
- The Lists suites stay green; `build`, `typecheck:tests`, `test`, `test:coverage` green; coverage
  floors only rise.

Replaced assertions: `sitefields.test.ts` "every column's `SchemaXml` version rises on the second
run" asserted the old every-run rewrite; it is now "the second apply writes nothing (versions
unchanged)". A deliberate behaviour change (S2), not a deleted assertion. The first test's title
changed to match.

## Verification Plan

- `npm test`, `npm run typecheck:tests`, `npm run build`, eslint and prettier on the changed files,
  `npm run test:coverage`.
- `npm run test:integration -- SiteFields Lists` against the reserved site.
- PP365's Standardmal (with an extension) applied twice on a scratch web: the second run writes no
  site column (the Calculated `GtRiskFactor` FieldRefs compare as the same).
- By hand, later: the vei hub template imported twice through PP365's PackageInstaller on a test
  hub: the second import writes no site column and logs no Warning; a hub admin's formatter on a
  package column survives (unless the template states it, see Open Questions).

## Open Questions

- A stored GUID target whose list or term set no longer exists (an owner deleted the list, the
  template creates it again): Warning and keep leaves the column pointing at a dead list. Should a
  target that no longer exists count as unset? Not decided; today: kept, with the Warning.
- Consumer XML often states empty values (`CustomFormatter=""` on the vei hub's User columns): the
  merge keeps only what the template leaves out, so an admin's formatter there is still reset by
  each import. Treat an empty template attribute as "not stated"?
- An unresolved token in a non-target attribute on update (`{sitename}` in a `Description`) is
  written as it stands, with an Error. Keep the stored value instead?
- On a subweb, a list FieldRef to a column only a parent web has is now an Error (Lists reads
  `web.fields`, the web's own). Reading `availablefields` there would add the parent's column to the
  list, as PnP does; a behaviour change for Lists, not done here.
- The PP365 project setup sends engine logs only to the console (`ApplyTemplate/index.ts:31-37`): a
  failing site column that used to stop setup is now an Error line nobody sees there. Hub import
  shows Error lines in its advanced log. PP365 may want to surface Errors.

## Rollout Notes

- **1.5.0, Changed** (with Lists): re-applying a template no longer rewrites unchanged site columns;
  changes are merged onto what the column holds; narrowing type changes and new targets are
  Warnings; a failing column no longer stops the run; `{listid:}` in a site column resolves to a list
  of the same template; unresolved tokens in site columns are Errors (runs that were clean can now
  log an Error, and `--strict` and PP365's install log see it).
- **Hub import behaviour** (PP365 `PackageInstaller`, on every import, update and cloud publish):
  package site columns are no longer written when unchanged, a hub admin's additions the template
  does not state survive, a type mismatch is no longer applied (Warning), and a column that fails
  no longer aborts the import.
- **PP365 follow-ups** (separate change, when PP365 bumps to 1.5.0; it bundles the engine at build
  time, so older hubs and projects keep today's behaviour):
  - CompatibilityService says a same-name column with another type or id "will be overwritten"
    (`CompatibilityService.ts:170-189`): with S2 and S4 the engine keeps the stored type or id and
    logs a Warning. The wording should say so.
  - "Blocked" same-id/other-name columns (`CompatibilityService.ts:190-199`, stripped by
    `PackageInstaller.ts:553-559`) can stop being stripped: the engine logs an Error and creates
    nothing, without aborting `applyTemplate`.
- The README's token table and ARCHITECTURE's handler order gain the S1 exception and S1d; the
  integration-testing matrix row and the pin list change (doc notes for the docs change).
