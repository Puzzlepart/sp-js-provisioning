# Spec: Navigation re-apply (reconcile in place, all or nothing)

Status: Verified (2026-10-29)

Owner: Remi Blom-Ohlsen

Related: [integration testing](./integration-testing.md) (pinned the refused-node defect),
[Lists re-apply](./lists-re-apply.md) (the same goal for lists), the SiteFields re-apply spec
written on the same branch. Ships in 1.5.0 with Lists and SiteFields.

## Problem

The Navigation handler came from pnp-js-provisioning in 2018 and was written for one run on a
fresh site whose default menu had to go. For each collection the template names (Quick Launch,
top navigation bar) it read the nodes, deleted every one of them, then added the template's nodes
in order (`src/handlers/navigation.ts` at `cdf7886`, lines 54-118). PP365 applies a template
again when an admin re-runs project setup, when CreateParent applies the parent template, and
for each extension after the main template, all through one `WebProvisioner`. Three defects:

1. **A refused node leaves half a menu** (pinned `{ todo }`, `tests/integration/navigation.test.ts`
   at `cdf7886`, 175-209). SharePoint refuses a node whose URL resolves to a server-relative path
   that is no file or folder (probe N1). By then every old node was deleted: the collection held
   only the nodes added before the refused one, its children and every later node were never
   added, and WebSettings and Hooks did not run. The error named neither the node nor the URL:
   the POST goes to the collection, and PnPjs's message carries only SharePoint's text.
2. **Every re-apply deleted and rebuilt the menu.** Each node got a new id, an owner's change of
   a child URL or of the order was lost, audiences (`AudienceIds`) were lost, and N unchanged nodes
   cost 2N writes. Every URL was validated again on add, so a page or list deleted since the last
   run made an unchanged re-run fail, and defect 1 truncated the menu there. The existing-title
   rule (a node whose title exists once keeps that node's URL unless `IgnoreExisting`) saw only
   the top level, since children were read after their parent was re-added, and it wrote the
   stored URL into the caller's template (`node.Url = ...`).
3. **The two collections were written in parallel** (`Promise.all`). When one rejected, the
   handler failed at once while the other went on deleting and adding in the background, and an
   error from the second was never reported.

## Goals

- An unchanged re-apply sends no write: one read per collection, nothing else (and one read of
  the web when a node sets `IgnoreExisting`).
- A node the template keeps keeps its id, its audiences and its stored URL (unless
  `IgnoreExisting`), at every level.
- A refused node fails the run and leaves both collections as they were; the message names the
  node and the URL sent.
- A successful run ends as before: the template's titles and order, nothing else in a collection
  the template names (PP365's first setup on a fresh site removes SharePoint's default nodes).
- The pinned `{ todo }` flips to a plain test.

## Non-Goals

- A keep-existing mode (a node the template does not declare is kept). Recorded as a later
  option, `Navigation.KeepExistingNodes?: boolean` (not PnP's `RemoveExistingNodes`, whose default
  is the other way round). It would allow additive extensions and let PP365 keep owners' links on
  a re-run without its localStorage workaround.
- PnP Framework's external-link retry (a refused node added again with `IsExternal`): a REST
  create ignores `IsExternal` and validates anyway (N3d), and the retry would hide broken links.
- `TokenHelper` tokens in `Title` or `Url` (`{parameter:}`, `{sitename}`, ...): `Url` keeps the URL
  tokens only, `Title` none. No consumer uses engine tokens in navigation.
- `{site}`: unchanged, the site collection's server-relative URL (on a subweb not the web).
- Case-insensitive title matching, pairing duplicated titles in order, matching by URL.
- `IsVisible`: REST cannot store `false` (N3d).
- Validating every URL before writing: it would copy SharePoint's undocumented rules (N1) and
  still miss a target deleted in between.

## Constraints

- `src/` stays TypeScript 4.0 / ES5; PnPjs 4's `select()` sets the query on the queryable itself,
  so every read goes through a fresh one.
- REST `add` only appends; `MoveAfter(nodeId, previousNodeId)` is the only way to reorder, has no
  "first" position, and checks neither collection nor level (N4).
- SharePoint answers success for a read, update or delete of a node it does not have (X1).
- A deleted node cannot be restored: a node added again gets a new id and no audiences (N6). So
  nothing is deleted until every add and URL update of both collections was accepted.
- SharePoint returns its own nodes' titles in the signed-in account's language (N5).
- A failure in the handler stops the template, as before (PP365 shows a failed setup task).

## Proposed Approach

### Read

Each collection is read once, through a fresh queryable, with its children expanded to the
template's depth plus one level, so the children of the deepest template nodes are known too:
`select('Id', 'Title', 'Url', 'Children/Id', ...)` with `expand('Children', 'Children/Children',
...)` (N7). A node SharePoint returns without `Children` at a level the template needs is read on
its own (a fallback; the tenant run showed it is not needed).

### Match

At each level, each template node is matched to the first stored node not matched yet whose title
is exactly the same (case included). A matched node is kept: nothing is written for it unless its
URL has to change. Its children are reconciled the same way against the template node's
`Children`; a template node without `Children` owns an empty child collection, so a kept node
loses the children it had. A template node without a match is added, with its children. A stored
node without a match is deleted (its subtree goes with it, N6): SharePoint's defaults, a stale
node, a node an owner added, a duplicate.

### URL of a kept node (at every level)

- Without `IgnoreExisting` a kept node keeps its stored URL. SharePoint rewrites the URL when the
  target page is renamed (N3b), and an owner may have changed it.
- With `IgnoreExisting: true` the template's URL (tokens replaced) is set when it differs from the
  stored one once both are normalised. Only `Url` is sent, so the node keeps its audiences (N3c).
- A title stored more than once at the level: the first is kept and gets the template's URL, as
  `IgnoreExisting` would (the old handler's fallback); the others are deleted.
- The template object is never changed.

### URL normaliser (comparison only)

`normaliseNavigationUrl(url, { serverRelativeUrl, origin })` in `src/util/navigationurl.ts`, built
on what SharePoint stored (N1): an empty URL is the web and `#` the web followed by `#`; a
web-relative URL is resolved against the web; an absolute URL on the web's origin becomes
server-relative; in a server-relative URL the path is decoded (`%20` and a space are the same),
loses a trailing slash and is compared without case (the read-back has the target's own case);
a query string or fragment is kept as is; any other absolute URL (another host,
`http://linkless.header/`) is external and compared verbatim. The web's URLs are read once, and
only when a URL has to be compared. The result is never sent.

### Write order, journal and rollback

1. Quick Launch: read, adds (depth first, in template order), then URL updates.
2. Top navigation bar: the same.
3. Quick Launch: deletes, then moves. Top navigation bar: deletes, then moves.

Steps 1 and 2 are journaled in one journal for the handler: each add (the top-most node only:
deleting it removes the children added under it) and each URL update with the URL it replaced.
When any of them fails (SharePoint's refusal is `-2130247147, Microsoft.SharePoint.SPException`,
N2; any other failure is treated the same), the journal is undone, the last write first: an added
node is deleted, a changed URL is set back (a refused add or update leaves nothing to undo, N1 and
N3a). Then each collection that had journaled writes is read back, since SharePoint reports
success for a write to a node it does not have (X1): a node that is still there, or a URL that was
not set back, is logged with `log_error` naming it, as is a rollback step that fails. Then the run
fails with:

`Navigation node '<Title>' (<URL sent>) was refused by SharePoint: <SharePoint's message>`

SharePoint's message comes from the response body's `odata.error.message.value` (in the signed-in
account's language, N2); without one, the error's own message. A collection that cannot be read
fails the run the same way (`Failed to read the top navigation bar: ...`), after the rollback.
Nothing was deleted or moved, so both collections are as they were.

Steps 3 run only when steps 1 and 2 succeeded. A failed delete or move is logged with `log_error`
naming the node and the level (`Failed to delete navigation node 'X' (12) from the Quick Launch:
...`, `Failed to move navigation node 'A' after 'B' in the children of 'P' in the Quick Launch:
...`; SharePoint's own text for a failed move is only "Unknown error", N4) and fails the run as it
is: deletes cannot be undone, so there is no rollback after step 3 starts.

### Moves

After its deletes, a level holds only template nodes: the kept ones in their stored order, then
the added ones in template order. For each template node `t[i]` (`i >= 1`) that is not already
right after `t[i-1]`, `MoveAfter(t[i], t[i-1])` is sent to that level's collection (for children:
`getById(parent).children`), and the local order is updated. Both ids come from the same read or
from this run's adds at that level (N4: a wrong id silently moves a node to another level or
collection). The invariant `t[0..i]` contiguous holds after each step, and with only template
nodes left the whole level is in template order, `t[0]` first, without a "move to first" (N4). A
level where something moved is read back (`select('Id')`); another order fails the run with
`The nodes in <level> are not in the template's order after moving them: ...`.

### Logging

Info per collection at the end: `Navigation nodes in the Quick Launch: 1 added, 0 URL(s) updated,
2 deleted, 1 moved.` Info when a rollback starts. Errors as above; `scope_ended(error)` logs the
error that fails the run.

## Decisions (2026-10-29)

| # | Question | Decision |
|---|---|---|
| N1 | Re-apply | Reconcile in place, replace mode: match by exact title at the same level (first unmatched), keep matched nodes, add missing, delete unmatched (owner nodes too), `MoveAfter` into the template order, recurse into kept parents. An unchanged run writes nothing; ids and audiences of kept nodes stay. No keep-existing mode now (`Navigation.KeepExistingNodes` later) |
| N2 | A refused node | Write order adds, URL updates, deletes, moves; a journal undoes this run's adds and URL updates; then `Error("Navigation node '<Title>' (<resolved URL>) was refused by SharePoint: <message>")`. A failed rollback step is an Error naming the node. The run fails (as before, and as the pin requires) |
| N3 | Two collections | Quick Launch, then top navigation bar, sequentially, one journal: a refusal in either leaves both as they were |
| N4 | Existing URL | At every level a kept node keeps its stored URL unless `IgnoreExisting`; with it, the template's URL is set only where it differs after normalisation. Exact-case titles. A title stored twice: the first is kept with the template's URL. The template object is never changed |
| N5 | Kept nodes | Not written, so `AudienceIds` stay (and `IsVisible`, which REST cannot set to false anyway) |
| N6 | Tokens | `Url`: `replaceUrlTokens` only; `Title`: none; a normaliser for comparison only; `{site}` unchanged |
| N7 | Docs | `IgnoreExisting` explained in README, ARCHITECTURE and the `schema.ts` JSDoc |
| — | Defaults set without a decision | A template node without `Children` owns an empty child collection (a kept node loses its children, as before); deletes before moves; the moved level is read back |

## Evidence: tenant probe (2026-10-29)

Three throwaway probes, each on its own STS#3 scratch web below the reserved site (nb-NO, 1044),
all removed afterwards; nothing in the repository. `{web}` is the scratch web's server-relative
URL, `{site}` the site collection's.

| # | Question | SharePoint did | So the handler |
|---|---|---|---|
| N1 | Which add URLs are refused, and how are accepted ones stored? | Refused (500, nothing left behind): a missing page or file (web-relative, server-relative or same-host absolute), another web's `_layouts`, `/`, another site collection, a library's localised title as its URL. Accepted and stored: web-relative resolved against the web; `''` and `{web}/` as `{web}`; `#` as `{web}#`; same-host absolute as server-relative; `%20` decoded; query and fragment kept; external URLs verbatim (spaces kept); the read-back in the target's own case | compares through the normaliser; undoes only earlier writes (a refused add leaves nothing) |
| N2 | Is the refusal's code language independent? | `-2130247147, Microsoft.SharePoint.SPException` on every refusal, whatever `Accept-Language` or the web's language; the message is in the account's UI language and names the resolved URL, not the node | builds its own message from the title and the URL sent, with SharePoint's text after it |
| N3a | Is a `Url` MERGE to a missing target refused? | Yes, same status and code; the node keeps its URL | journals URL updates and sets them back |
| N3b | Can a node whose target was deleted be updated? | The premise does not occur: a recycled file, page or list takes its nodes with it (a child node and a list's "Siste" child too); a renamed page updates the stored URL | a re-run re-adds such a node, which is refused while the target is missing (the integration suite asserts it); a kept node keeps its stored URL |
| N3c | `AudienceIds`? | Plain array MERGE accepted and stored unchecked; kept through a Title-only MERGE; `[]` reads `null` | sends only `Url` when it updates |
| N3d | `IsVisible: false`, `IsExternal` on create | Every `IsVisible: false` (MERGE or create) reads back `true`; a create with `IsExternal: true` and a missing target is refused | no visibility state to keep; no external retry |
| N4 | `MoveAfter` | Moves `nodeId` into the collection in the URL after `previousNodeId`; no call moves to the first position (`0`, `-1`, omitted: 500 "Ukjent feil"; `null`: 400); ids from another collection or level move the node there without an error, even under itself | chains `MoveAfter(t[i], t[i-1])` after the deletes, with sibling ids from one read, names the node in its message, reads the level back |
| N5 | Titles and language | `Title` is in the signed-in account's UI language (`TitleResource`), not the web's; REST ignores `Accept-Language`. SharePoint's own nodes have translations (Hjem/Home, Dokumenter/Documents, ...); added nodes have one title in every language | exact-title matching depends on who runs the setup (Rollout Notes) |
| N6 | Subtrees and ids | Deleting a parent deletes its subtree; a node added again gets a new id; any collection's `getById(id)` reaches any node | deletes no child of a deleted node; rolls back only the top-most add; addresses every node through its collection's `getById` |
| N7 | One read | `select(Id,Title,Url,...,Children).expand('Children')` returns two levels, `Children/Children` more; unknown `$select` names are ignored silently | one read per collection (verified three levels deep by the integration suite) |
| X1 | A node id that does not exist | `GET` answers 200 `odata.null`, `MERGE` and `DELETE` succeed; only calls under a missing parent's children answer 404 | reads back after the rollback and after moves instead of trusting success |

## Acceptance Criteria

- The pinned test passes as a plain test: a refused node fails the run with handler `Navigation`
  and a message naming the node and the URL, and the Quick Launch is unchanged, ids included.
- A second apply of `navigation.json` keeps every node id and sends no write (recording web); the
  only navigation requests are one read per collection, three levels deep.
- An owner's node is removed, the order an owner changed (top level and children) is restored with
  moves only, and a kept node keeps its id, its audiences and the child URL an owner changed.
- `navigation.b.json` over `navigation.json`: kept nodes keep their ids; a kept child keeps its
  stored URL; an `IgnoreExisting` node, top level or child, gets the template's URL; a kept node
  without `Children` in B loses its children.
- A refusal in the top navigation bar leaves the Quick Launch's adds and URL changes undone.
- A node whose target was recycled is gone with it, and re-applying the unchanged template fails
  naming that node and changes nothing else.
- Unit tests for each rule (call log of a fake web); the template object is unchanged; the
  normaliser is table-tested.

## Verification Plan

- `npm test`, `npm run typecheck:tests`, `npm run build`, eslint and prettier on the changed files,
  `npm run test:coverage`.
- `npm run test:integration -- Navigation` (twice).
- PP365: Standardmal with the EnkelVenstremeny extension through one `WebProvisioner` (as the
  setup wizard applies them), then again with a new provisioner, on a scratch web: no Navigation
  Error or Warning, and the Quick Launch after the second run equals the first (same ids, titles,
  URLs and order). The second run is not write-free: each `applyTemplate` reconciles the menu to
  its own template (see Rollout Notes).

## Verification Result (2026-10-29)

- Offline: `tests/unit/navigation.test.ts` 25 tests (a fake web with one store for both
  collections and an ordered call log), `tests/unit/util.test.ts` 21 normaliser cases, all pass;
  `npm test` 258 tests, 258 pass (both specs of the branch); `npm run build` and
  `npm run typecheck:tests` green; eslint and prettier clean on the changed files;
  `npm run test:coverage` above its floors (all files 80.37 / 84.02 / 80.21 lines / branches /
  functions, both specs of the branch; floors raised to 79 / 83 / 79; `navigation.ts` 97.43 / 94.24 / 97.78,
  `navigationurl.ts` 98.51 / 86.96; the uncovered lines are a failed read-back after a rollback
  and after the moves).
- `npm run test:integration -- Navigation` against the reserved site, three times (the later runs
  after tightening the assertions): reachability 6 of 6, Navigation 6 tests, 6 pass, 0 `todo`,
  0 fail, about 46 s; no leftover scratch web after the last run. The first test's request log
  confirmed that the nested `$select`/`$expand` read returns three levels in one request per
  collection (no fallback read); the owner test confirmed a `MoveAfter` on a child collection
  (`quicklaunch(<id>)/children/MoveAfter`) next to one on the Quick Launch.
- PP365 (throwaway script on a scratch web, removed afterwards; SiteFields, Lists, ClientSidePages
  and Navigation): Standardmal then EnkelVenstremeny through one provisioner, then both again with
  a new provisioner. Run 1: template step `16 added, 0 URL(s) updated, 2 deleted, 12 moved`
  (SharePoint's Hjem, Notatblokk, Dokumenter and Områdeinnhold kept by title; Notatblokk kept its
  stored OneNote URL, not the template's `#`), extension step `0 added, 0 URL(s) updated,
  7 deleted, 2 moved`. Run 2: template step `12 added, 0 URL(s) updated, 0 deleted, 8 moved`,
  extension step `0 added, 0 URL(s) updated, 7 deleted, 2 moved`; the final Quick Launch equals
  run 1's (the same 8 ids). No Navigation Error or Warning in either run; 1.4.0 deleted and
  re-added every node in both steps.
- Review (2026-10-29), five reviewers and an adversarial check of each finding, fixed with a unit
  test that fails without the fix: a title the template declares twice, stored once, was rewritten
  on the second unchanged run (duplicates now count the template's nodes too); the normaliser left
  reserved escapes (`%2C`, `%26`) encoded, so an `IgnoreExisting` node with one was MERGEd on every
  run (now decoded per path segment). Afterwards `npm test` 265 tests, all pass;
  `test:coverage` 80.50 / 84.07 / 80.41.
- Full integration run after the fixes (`npm run test:integration`): reachability 6 of 6; the
  suites 52 tests, 51 pass, 0 `todo`, 1 skipped (custom script blocked), 0 fail; nothing left
  behind.
- Assertions replaced (behaviour changed by the decisions), none deleted:
  - `tests/integration/navigation.test.ts` (at `cdf7886`, 123-130) "the second run deleted and
    re-added every node ... new ids": now the second run keeps every id and sends no write, and
    reads each collection once.
  - the same file, 154-163 (B: "A child never keeps its URL"): the child `SPJS Itest Master pages`
    now keeps A's URL; `navigation.b.json` gains an `IgnoreExisting` child (`SPJS Itest Contoso`)
    whose template URL wins, and B removes the grandchild A gave it.
  - the same file, 175-209: the `{ todo }` is a plain test; it compares the whole tree with ids
    (it compared titles and URLs) and the message's format.
  - `tests/unit/navigation.test.ts` (at `cdf7886`, 88-103) "replaces the existing nodes": same
    outcome, now asserted on the call log (the add before the delete).
- New integration tests: an owner's changes (audiences, a child URL, an added node, a move), a
  refusal in the top navigation bar undoing the Quick Launch, and a recycled target.
  `navigation.json` gains a third level (`SPJS Itest Contoso grandchild`).

## Open Questions

- Keep-existing mode (`Navigation.KeepExistingNodes`): wanted when PP365 should keep owners' links
  on a re-run, or for additive extensions. Template-level flag or run mode?
- What an English-UI account reads for SharePoint's own nodes (inferred from `TitleResource`, not
  observed: no such account on the tenant).
- Should a failure after the deletes started (a delete or move) also undo the adds of the run? Today
  the run fails with the template's nodes added and the old ones partly deleted; a re-run
  reconciles from there.
- Should the engine refuse to replace a hub site's top navigation bar (the hub navigation of every
  associated site)? No template has a `TopNavigationBar` today.

## Rollout Notes

- **1.5.0, Changed**: re-applying a template reconciles the navigation in place instead of deleting
  and re-adding every node: unchanged nodes keep their ids and audiences and nothing is written for
  them; a kept child keeps its stored URL like a top-level node (a template change to an existing
  child's URL now needs `IgnoreExisting`). **Fixed**: a refused node no longer leaves half a menu,
  and its error names the node and the URL; the two collections are written one after the other,
  so a failure in one no longer lets the other carry on unreported; the template object is no
  longer changed.
- **PP365 CreateParent** applies the parent template and then appends the menu it saved
  (`projectSetup/index.ts:271-285`). That relies on the template deleting every node it does not
  declare, which replace mode keeps; a keep-existing default would duplicate every node.
- **Title matching depends on the UI language of the account that runs the setup** (N5): a PP365
  template's `Hjem` or `Dokumenter` matches SharePoint's node only for a Norwegian-UI account, so
  that node keeps its URL only then. Otherwise it is added with the template's URL and the stored
  one deleted: the end result of 1.4.0. Nodes the engine adds have one title in every language and
  match whoever runs the next setup.
- **`IsVisible` cannot be stored through REST** (N3d): a node hidden in the browser cannot be read
  or kept as hidden by this engine; the template has no `IsVisible`.
- A page, file or list deleted since the last run takes its node with it (N3b); an unchanged
  re-apply then fails on that node, as in 1.4.0, but leaves the menu as it was and names the node.
- PP365's own follow-ups are unchanged by this spec: the menu "Run project setup" saves is still
  not restored, and an extension with `Navigation` still replaces the main template's menu.
- **A template plus an extension that both declare `Navigation`** (PP365 applies the main
  template's whole schema and then each extension, on every run): each `applyTemplate` reconciles
  the menu to its own template, so a re-run adds the main template's extra nodes back and the
  extension deletes them again. The nodes both declare keep their ids, and the end result is
  stable, but such a re-run is not write-free, and it validates the targets of the main template's
  extra nodes again: if an owner deleted one of them, the template step fails on a node the final
  menu does not have, and the extension does not run. Follow-up in PP365: leave out the main
  template's `Navigation` when a selected extension declares one (or a later engine option).
