# Spec: Re-applying Lists Without Data Loss

Status: Verified (2026-10-27)

Owner: Remi Blom-Ohlsen

Related: [integration testing](./integration-testing.md) (D4 pinned these defects),
[site groups and web security](./site-groups-and-web-security.md) (the web-level guard this spec
mirrors for lists). Ships in 1.5.0 with Security.

## Problem

Templates are applied to sites that already hold data, and the Lists handler was written for the
first run. PP365 applies a template again when an admin re-runs project setup ("Run project setup",
`ProjectSetupCustomAction.ts:29-30`: `skipAlreadySetupCheck`, `forceTemplate`), when a setup that
failed reloads, and when a package is imported or updated on the hub
(`PackageInstaller._provisionHub`, `PackageInstaller.ts:535-571`). It applies the template and then
each selected extension through one `WebProvisioner` instance (`ApplyTemplate/index.ts:31-76`).

What a second apply does today, each pinned by a `{ todo }` test in Phase 1 or found by the closing
analysis (2026-10-26):

1. **List `Fields` are deleted and recreated on every run** (`processField`). The values in the
   column are lost, and so are the column in views the template does not name, its content type
   links, its place in forms, and anything an admin set on it. A create that fails after the delete
   leaves no column at all. The standard project and program templates declare 9 lookups as list
   Fields in 6 lists; a re-run of setup empties the 6 that have no dependent lookup (SharePoint
   refuses to delete a primary lookup with dependents, so those 3 survive, with an Error per
   re-run). Forskningsmal declares 30 more. On the hub, PP365 1.14's CompatibilityService strips
   existing list Fields from a re-import as "blocked" (`CompatibilityService.ts:203-264`, 384-402),
   so template changes never reach those columns, and the check fails open (a throw logs a Warning
   and strips nothing). A DataRows `Skip` row loses its value in such a column for the same reason.
2. **`RemoveExistingViews` deletes every view on every run**, the template's own included, and adds
   them again: new ids (links and web parts break), user views gone, page web part views gone, and
   one refused delete stops the Lists handler. PnP Framework has the same every-run meaning; no
   consumer template sets the flag.
3. **`RemoveExistingContentTypes` removes every unbound content type on every run**, so content
   types a customer added since are removed (one in use is refused). It matches by substring
   (`includes`) while bindings match by prefix. 75 consumer list entries set it.
4. **List settings are overwritten** by `lists.ensure`, which sends PnPjs' defaults
   (`Description: ""`, `EnableContentTypes: false`) for what the template omits: every hub package
   import clears the descriptions of 15 hub lists (Prosjekter, Prosjektstatus, ...).
5. **Template views churn and can be left half-filled**: every run removes all view fields and adds
   them again; one field that cannot be added leaves the view without the rest.
6. **List `breakRoleInheritance` is sent on every run** (`processListSecurity`). SharePoint ignores
   it on a list that already has unique permissions (probe 2026-10-26: also with
   `clearsubscopes=true`, an item keeps its unique permissions), but the engine depends on that, and
   it warns about assignments without `BreakRoleInheritance` even on lists that are already unique.
   Six hub lists in the hosting packages set list Security.
7. **DataRows**:
   - a `KeyColumn` value with a token is matched raw while the item holds the resolved value, so
     the row is added again on every run;
   - every row's payload is built before the key decides it is skipped, so a skipped row still
     queries lookup sources and calls `ensureUser`, and fails when they throw;
   - a URL `{ Url }` without `Description` gets the raw, unresolved Url as its description;
   - a single Lookup value that matches no item is dropped without a log line;
   - a LookupMulti entry that matches no item is sent as `null`, and SharePoint refuses the whole
     row (400) with a message that names neither the column nor the value. All 91 LookupMulti
     strings in the hosting packages are Norwegian titles of Prosjektkolonner items, so one renamed
     or English title loses a whole Porteføljevisninger row.
8. **FieldRef `AdditionalProperties` is broken**: it sets the column's title to its internal name,
   drops attributes whose new value is undefined, and sends the `SchemaXml` read before the Title
   update, with its old `Version`.

## Goals

- A second apply of an unchanged template changes no data, ids or user-made views, and sends no
  delete.
- Template changes still reach existing sites where SharePoint can apply them in place without
  losing data; where it cannot, the run says so (Warning) and leaves the site as it is.
- Every value the template asked for and the engine could not write is logged as an Error naming
  the list, the row, the column and the value.
- Every pinned Lists `{ todo }` flips to a plain test.

## Non-Goals

- Matching lists by URL instead of title (a renamed or differently localised list still gets a new
  twin). Recorded for a later spec.
- Removing duplicates the KeyColumn defect already created, or reconciling role assignments
  (PnP's `Remove: true` is a later, additive change).
- Taxonomy on the JSOM path (D3); the taxonomy list Field case below is verified with the
  integration suite once a fixture term set exists.
- PP365 changes. The CompatibilityService strip and the CopyListData KeyColumn projection
  (pp-veiprosjekt) are follow-ups in PP365 (Rollout Notes).

## Constraints

- Templates written for 1.4.0 must keep working on the first run: PP365's first setup applies
  extensions over lists the main template just created (Forskningsmal and Veimal replace
  Fasesjekkliste's content type and set views on the OOTB "Dokumenter" library).
- `src/` stays TypeScript 4.0 / ES5; PnPjs 4's `select()` mutates its queryable (read through fresh
  instances).
- A failure on one list or column is logged and the handler carries on, as elsewhere in the
  handler; nothing new may stop the whole Lists handler.

## Proposed Approach

### Lists made in this setup

`ProvisioningContext` gets `createdLists`, the titles `processList` creates. The context lives on
the `WebProvisioner` instance and is not reset by `applyTemplate`, so a list the main template
creates counts as made in this setup for the extensions applied after it through the same
instance, as PP365 does. A list that existed before (the OOTB "Dokumenter", or any list on a
re-run) is not.

### List Fields: update in place, never delete (1)

Read the list's fields once per list (`Id`, `InternalName`, `TypeAsString`, `SchemaXml`), fresh
queryable. For each template field, matched by normalised `ID`:

- **Missing**: create as today (`createFieldAsXml`, then `Title` from `DisplayName`), primary
  lookups before their dependents (`FieldRef=`), whatever the template order.
- **Present, same type**: merge the template onto the stored `SchemaXml` and send one
  `update({ SchemaXml }, 'SP.Field')` (the type argument saves PnPjs' extra read). Template
  attributes win; the storage attributes the template does not carry (`ColName`, `RowOrdinal`,
  `SourceID`, `Name`, `StaticName`) are kept; `Version` is dropped; child elements the template
  declares (`CHOICES`, `Default`, `Formula`, `Validation`) replace the stored ones; `Customization`
  properties are merged by name, keeping the stored `TextField` (the taxonomy note field). No write
  when the merged XML equals the stored one after normalisation (attribute order, `TRUE`/`true`,
  GUID braces and case). The `DisplayName = Name` swap stays on create only.
- **Present, widening type change** (`Text` to `Note`, `Lookup` to `LookupMulti`, `User` to
  `UserMulti`): merged in place as above; SharePoint keeps the values (probe: `Text` to `Note`,
  `Lookup` to `LookupMulti`; `User` to `UserMulti` by analogy, not yet probed). The type is
  written as SharePoint stores it: `Type="Lookup" Mult="TRUE"` merges as `LookupMulti`.
- **Present, any other type change** (narrowing, such as multi to single, or a different kind):
  Warning naming the list, column and both types; the column is left as it is. Probe: SharePoint
  accepts `LookupMulti` to `Lookup` and silently keeps one value of two.
- **Another target** (a lookup's `List` or `WebId`, a taxonomy field's `TermSetId` or `SspId`): on
  a list made in this setup, which holds no values yet, the field follows the template (a content
  type can bring a site column with another target: the vei hub's Fasesjekkliste binds the
  content type with PP365's `GtProjectPhase` and declares the same field on the vei term set).
  Elsewhere a Warning, the stored target is kept and the rest is merged. Probe: SharePoint accepts
  a lookup re-target and the stored ids then point at other items in the new list.
- **Not found by ID, but the internal name exists with another ID**: Warning, no create.
- **A read that fails other than 404**: Error, nothing written for that list's fields.

The merge helper lives in `src/util` and replaces `addFieldAttributes` for the FieldRef path too
(8): `AdditionalProperties` are merged onto `SchemaXml` read after the `Title`/`Required`/`Hidden`
update, without `Version`, without touching `DisplayName`, and an undefined value no longer removes
an attribute.

### Views (2, 5)

- `RemoveExistingViews` acts only on lists made in this setup: after the template views are
  processed, delete the list's other views, one at a time, each failure logged as a Warning. Never
  delete a view the template names, a hidden view whose URL is outside the list's root folder (a
  page's list view web part, PnP parity), a personal view, or the default view unless a template
  view became default. On any other list it does nothing.
- The template owns the views it names, on every run (the extensions' first run needs it on the
  OOTB "Dokumenter"): settings are updated in place, and view fields are set only when they differ
  from the template's, so an unchanged view is not touched. When they differ the view gets exactly
  the template's fields; a field that cannot be added is logged as an Error naming the view and the
  field, and the others are still added.

### Content types (3)

`RemoveExistingContentTypes` removes unbound content types only on lists made in this setup.
Bindings and removal match the same way: a list content type belongs to a binding when its id is
the binding's id or a child of it (prefix). Each removal is caught: a content type in use (the
500 SharePoint answers; its message is localised) is a Warning, as today, any other failure an
Error.

### List settings (4)

An existing list is updated with only what the template defines: `Description` and
`ContentTypesEnabled` when the template has them, and `AdditionalSettings`. A list the template
creates gets them as today (`lists.ensure` with the template's values).

### List security (6)

As web Security does: read `HasUniqueRoleAssignments` on its own queryable; break inheritance
(with `CopyRoleAssignments` and `ClearSubscopes`) only when the list still inherits, otherwise log
at Info. A failed read is an Error and inheritance is not broken; the role assignments are still
added. Role assignments stay additive. The "assignments without BreakRoleInheritance" Warning
fires only when the list inherits.

### DataRows (7)

- **Key first**: resolve the key value with the token helper, then OData-escape it, and match on
  that. Decide add, update or skip before building the payload; a skipped row builds nothing. The
  Skip log line names the resolved key. A Warning when more than one item matches the key (the
  first is updated, as today) and when `KeyColumn` is set but a row has no value for it (it is
  added).
- **URL**: resolve the `Url` once and use the result as the default `Description`; a given
  `Description` gets token replacement too (DataRows strings get tokens; ARCHITECTURE's token reach
  gains "URL Description").
- **Unresolved lookups** (single Lookup, LookupMulti entries, and likewise single User and UserMulti
  entries): an Error naming the list, the row (index and key), the column and each value, and
  distinguishing "no item has ShowField = value" from "the column has no lookup list". A new item
  gets the values that resolved; an existing item (Overwrite) is updated without that column, so the
  stored value is kept. The row is never refused because of one value that matches nothing, but a
  request that may succeed later (429, 503, 504, no answer) fails the row as in 1.4.0, so the next
  run writes it whole (a `Skip` row would otherwise never get the value). A taxonomy term without a
  `termId` or `label`, or a column without its note field, is an Error too.

## Decisions (2026-10-26)

| # | Question | Decision |
|---|---|---|
| L1 | An existing list Field | Update in place, never delete; Warning and no change for a narrowing type change, or a new target (lookup list, term set) on a list not made in this setup |
| L2 | `RemoveExistingViews` | Only on lists made in this setup, never a template view |
| L3 | A lookup value that matches no item | Error; a new item gets what resolved, an existing item keeps its stored value |
| L4 | Scope | Also list settings, `RemoveExistingContentTypes`, template views, FieldRef `AdditionalProperties` |
| L5 | `RemoveExistingContentTypes` | Only on lists made in this setup |
| L6 | Template views on later runs | The template owns the views it names: updated when they differ, never half-filled (an extension's first run needs it) |
| L7 | Release | In 1.5.0 with Security |
| L8 | Widening type changes | Applied in place (`Text` to `Note`, Lookup or User single to multi), per the probe (2026-10-27) |

## Evidence: tenant probe (2026-10-26)

A throwaway probe on a scratch web below the reserved site (not in the repository), list fields
created with `createFieldAsXml` and one item with values, then a `SchemaXml` MERGE per case:

| Case | SharePoint | Value |
|---|---|---|
| Same type, merged onto the stored XML, no `Version` (DisplayName, Description, MaxLength) | accepted, same id, applied | kept |
| The template's XML as is (no `ColName`, `RowOrdinal`, `SourceID`) | accepted, `ColName` kept | kept |
| A stale `Version` | accepted, no conflict | kept |
| `Text` to `Note` | accepted, type changed, same id | kept |
| `CHOICES` replaced | accepted | kept |
| `Lookup` to `LookupMulti` | accepted | kept (`[3]`) |
| `LookupMulti` to `Lookup`, item with `[1, 3]` | accepted | **`1`: one value lost silently** |
| Lookup `List` re-targeted to another list | accepted, re-targeted | **id `2` kept, now another item** |
| Primary lookup merged while a dependent exists; the dependent (ReadOnly) merged | accepted, dependent intact | — |
| Primary lookup with a dependent deleted (today's recreate) | **refused (500)** | — |
| `breakroleinheritance(false, clearsubscopes=true)` on a list with unique permissions | accepted and ignored: the item keeps its unique permissions | — |
| Taxonomy list field without `TextField` | not run: the probe found no term set | — |

## Verification Result (2026-10-27)

- Offline: `lint:check`, `build`, `typecheck:tests`, 179 unit tests, `test:coverage` (75.01 / 79.50 /
  75.99, floors raised to 73/78/74), `check:pack`, `check:secrets`, `check-skills` green.
- `npm run test:integration` against the reserved site: 51 tests, 48 pass, 2 `todo` (Navigation
  and SiteFields, other specs), 1 skipped (custom script blocked), 0 fail; nothing left behind. The
  Lists suites: 18 tests, all pass, the 8 former pins among them.
- PP365's standard project template (`Standardmal`, `SiteFields` and `Lists`) and the Forskningsmal
  extension applied through one `WebProvisioner` on a scratch web, values written into three list
  Fields lookups (single, multi, a primary with dependents), then the setup run again with a new
  provisioner: every value kept, no `DELETE`, no error from a list field. The re-run sent 6 field
  MERGEs: 4 are FieldRef Title updates (as before), 2 are `GtDeliveryStakeholders`, which the
  template and Forskningsmal declare with different display names (a conflict in the templates).
  The errors that run logged were content types and view columns a blank web does not have (the
  ContentTypes handler is JSOM).
- A closing review (four reviewers, each finding verified by a second agent) found 10 issues, all
  fixed with tests: attribute values were written back unescaped (`&`, `<`, line breaks from an
  admin's column formatting or description), a braced term set id was rewritten on every re-import,
  `Type="Lookup" Mult="TRUE"` was merged onto a LookupMulti, every content type removal failure was
  a Warning, a transient lookup failure left a `Skip` row without its value for good, untested
  `createdLists` across `applyTemplate` calls, and three test and text gaps. The integration suite
  then found a view without `AdditionalSettings` sent an empty update (400), fixed.

## Acceptance Criteria

- The pinned tests pass as plain tests: list Fields keep their values (`lists.test.ts`), no
  `breakroleinheritance` on an already-unique list, `RemoveExistingViews` keeps view ids, and the
  five DataRows pins (`lists-datarows.test.ts`).
- A second apply of `lists.json` sends no `DELETE` (recording web), and the read-backs are equal.
- `lists.b.json` (new) changes list Field display names, a description and a `MaxLength`; applied
  over `lists.json` the columns keep their ids and values and show the changes; `lists.json`
  applied again restores the display names and the `MaxLength`, and the description B added stays
  (the merge keeps what a template leaves out). `CHOICES` replacement is covered by the unit tests
  and the probe.
- A narrowing type change and a lookup re-target in an inline template give one Warning each and
  change nothing; `Text` to `Note` is applied with the value kept.
- A view a user added survives a re-apply with `RemoveExistingViews`; on a list made in this setup
  only the template views remain, and a template view with the OOTB title keeps `AllItems.aspx`
  (integration). A page's list web part view, a personal view and the default view are kept (unit,
  `removeOtherViews`).
- An existing list keeps a description the template does not set; a content type a customer added
  survives a re-apply with `RemoveExistingContentTypes`.
- `AdditionalProperties` on a FieldRef is applied without changing the column's title (an inline
  template in `lists.test.ts`).
- Unit tests for each rule above; coverage floors only rise; the offline CI sequence is green.

## Verification Plan

- `npm test`, `npm run typecheck:tests`, `npm run lint:check`, `npm run build`, `npm run test:coverage`.
- `npm run test:integration -- Lists` (both suites) twice, then the full run.
- `npm run provision -- --package ../prosjektportalen-hosting/packages/pp-forskningsprosjekt --twice --strict`
  and the PP365 standard project template with the Forskningsmal extension, on a scratch web:
  the second run logs no Error, deletes nothing, and the list Field values written between the runs
  are kept.
- A taxonomy list Field case once a fixture term set exists (Phase 2), or by hand on a site with
  PP365's term sets: the stored `TextField` survives the merge.

## Open Questions

- Should an unresolved lookup on an existing item clear the column instead of keeping it? (Decided:
  keep; revisit if PP365 asks for it.)

## Rollout Notes

- **1.5.0, Changed**: re-applying a template no longer deletes list Fields, views or content types
  outside the rules above; existing lists keep settings the template omits; unresolved lookups are
  logged as Errors (PP365 install logs gain lines where values were silently missing before).
- **PP365** (separate change, when it bumps to 1.5.0): CompatibilityService can stop classifying
  existing list Fields and FieldRefs as "blocked" and stripping them. PP365 bundles the engine at
  build time, so this ships in the same PP365 release that bumps the dependency; hubs and projects
  on older PP365 keep today's behaviour.
- **PP365 CopyListData** drops the KeyColumn from pp-veiprosjekt's projected rows
  (`manifest.json` projects `GtSortOrder,Title,GtProjectPhase,GtVeiAnchored`, the key is
  `GtDataRowUniqueKey`), so every row is added again on each run; the new Warning makes it visible,
  the fix belongs in PP365 or the package.
- The README and ARCHITECTURE describe `RemoveExistingViews` and `RemoveExistingContentTypes` as
  acting on lists made in this setup, a deliberate departure from PnP Framework.
