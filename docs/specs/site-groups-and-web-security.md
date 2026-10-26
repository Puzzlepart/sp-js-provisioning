# Spec: Site Groups And Web Security

Status: Verified (2026-10-21, `tests/integration/security.test.ts`, including a list granted to a group of the same template, and the list `Security` of `tests/integration/lists.test.ts`)

Owner: Remi Blom-Ohlsen

Related:

- Builds on list `Security` (1.3.15, `src/handlers/lists.ts` `processListSecurity`)
- PnP provisioning schema `Security` / `SiteGroups` element

## Problem

Since 1.3.15 a template can break role inheritance on a list and add role assignments, but the
principals have to exist already: the associated owner/member/visitor groups, an existing site
group or a user. A template cannot create its own SharePoint groups, put members in them or give
them permissions on the site itself.

Project templates therefore need a manual step (or a separate script) to create groups such as
"Styringsgruppe" or "Prosjektledere" before the list permissions in the template can refer to
them, and on a fresh site the list role assignments for those groups fail and are only logged.

## Goals

- A top-level `Security` section creates SharePoint site groups (`SiteGroups`) with title,
  description, owner, membership settings and members.
- Re-applying a template is safe: an existing group (matched by title) is updated, not
  duplicated.
- `Security` can break role inheritance on the web and add role assignments on the web, with
  the same principal and role definition rules as list `Security`.
- Groups are created before `Lists` runs, so list `Security` can grant permissions to groups
  created by the same template.
- Group titles, owners, members and role assignment principals support tokens, so groups can
  be named per site (`{sitename} Styringsgruppe`) or per parameter (`{parameter:Name}`).

## Non-Goals

- Custom role definitions (PnP `RoleDefinitions` with permission masks).
- Adding users to the associated groups (`AdditionalOwners`/`AdditionalMembers`/`AdditionalVisitors`).
- Changing which groups are the associated owner/member/visitor groups.
- Removing groups or role assignments that are not in the template.
- Folder- and item-level permissions.
- A `{groupid:...}` token.

## Constraints

- PnPjs v4 (`@pnp/sp` 4.21): `siteGroups.add`, `siteGroup.update`, `siteGroup.setUserAsOwner`,
  `siteGroup.users.add/removeById`, `web.ensureUser`, `web.breakRoleInheritance`,
  `web.roleAssignments.add`. No JSOM needed.
- Site groups are scoped to the site collection; a title is unique within it.
- SharePoint rejects role assignments on a securable object that still inherits permissions.
  The root web of a site collection always has unique permissions.
- `DefaultHandlerSort` is a contract. Consumers that pass their own `handlerSort` without a
  `Security` key get the default sort value 99 for it, which runs it after `Lists`.
- Consumers on an older version ignore the new top-level key (only sections with a handler run),
  so templates can adopt `Security` without breaking older clients.

## Proposed Approach

### Template contract (`src/schema.ts`)

```json
"Security": {
  "SiteGroups": [
    {
      "Title": "{sitename} Styringsgruppe",
      "Description": "Styringsgruppen for prosjektet",
      "Owner": "{associatedownergroupid}",
      "AllowMembersEditMembership": false,
      "OnlyAllowMembersViewMembership": false,
      "Members": ["ola.nordmann@contoso.no"],
      "ClearExistingMembers": false
    }
  ],
  "BreakRoleInheritance": false,
  "RoleAssignments": [
    { "Principal": "{sitename} Styringsgruppe", "RoleDefinition": "Lese" }
  ]
}
```

- `ISiteGroup`: `Title` (required), `Description`, `Owner`, `AllowMembersEditMembership`,
  `AllowRequestToJoinLeave`, `AutoAcceptRequestToJoinLeave`, `OnlyAllowMembersViewMembership`,
  `RequestToJoinLeaveEmailSetting`, `Members`, `ClearExistingMembers`.
- `IWebSecurity extends IListSecurity` adds `SiteGroups`. `BreakRoleInheritance`,
  `CopyRoleAssignments`, `ClearSubscopes` and `RoleAssignments` mean the same as for lists, but
  apply to the web.
- `Schema.Security?: IWebSecurity`; `schema.json` gets the matching definitions.

### Handler (`src/handlers/security.ts`)

New `Security` handler, registered in `Handler`, `DefaultHandlerMap` and `DefaultHandlerSort`
with sort value `2.5` (after `Features`, right before `Lists`; the other values are unchanged).

1. For each site group, in template order:
   1. Replace tokens in `Title`/`Description`.
   2. Look the group up by title (`siteGroups.filter(Title eq '...')`). Create it with
      `siteGroups.add` when missing; otherwise update the configured properties.
   3. Set the owner (`setUserAsOwner`) when `Owner` is set. The owner is resolved like a role
      assignment principal, so it can be an associated group token, a group title (including the
      group itself) or a user.
   4. Ensure each member with `web.ensureUser` and add it with `users.add(LoginName)`.
   5. With `ClearExistingMembers`, remove members whose login name is not in `Members`.
2. Web security: when `BreakRoleInheritance` is set and the web does not already have unique
   permissions, break inheritance (so a re-run with `ClearSubscopes` does not wipe list
   permissions again). Then add each role assignment.

Failures are logged per group, per member and per role assignment and never abort the template,
matching list `Security`. An unexpected failure outside those steps is thrown and wrapped in
`ProvisioningError` like other handlers.

### Shared principal resolution (`src/util/security.ts`)

The principal and role definition resolution in `Lists` moves to shared functions used by both
handlers, unchanged in behaviour. Callers replace tokens in the principal before resolving it,
so list `Security` also accepts `{sitename}` and `{parameter:...}` in principals.

### Token

`{sitename}` resolves to the title of the web being provisioned (PnP semantics), from
`context.web.Title`.

## Acceptance Criteria

- A template with `Security.SiteGroups` creates each missing group with the configured
  properties, owner and members.
- Re-applying the template updates the group and adds missing members without creating a
  duplicate group or failing.
- `ClearExistingMembers: true` leaves exactly the configured members in the group.
- `Security.RoleAssignments` add role assignments on the web; `BreakRoleInheritance` only breaks
  inheritance when the web still inherits.
- A list `Security` role assignment can reference a group created in the same template.
- One failing group, member or role assignment is logged and the rest is still applied.
- `{sitename}` resolves in group titles and principals.
- `npm run build`, `npm test` and `npm run typecheck:tests` pass.

## Verification Plan

- Unit tests (`tests/unit/security.test.ts`, plus `lists.test.ts` and token tests) with fake
  webs: create versus update, owner, members, `ClearExistingMembers`, web role assignments,
  inheritance check, per-item failure isolation, handler order and `{sitename}`.
- `npm run build`, `npm test`, `npm run typecheck:tests`.
- On a test tenant, `npm run test:integration -- Security Lists`
  ([integration testing](./integration-testing.md)): a scratch subweb with unique permissions and
  its own associated groups (D2 there), in place of a fresh group-connected site:
  - groups created, owner set, members added, list permissions for the new groups applied;
  - second run on the same web: no duplicates, no errors, the read-back unchanged;
  - group owner set to another group (`{associatedownergroupid}`) and to the group itself;
  - a nested web that inherits permissions: inheritance broken once, assignments copied.
- The 0a smoke run (`npm run provision -- --package tests/fixtures/templates/smoke.json --twice
  --strict`) is its first tenant evidence.

## Verification Result

- 2026-10-14: `npm run build`, `npm test` (30/30, of which 13 new) and `npm run typecheck:tests`
  pass. Prettier is clean; eslint reports only the `"''"` OData escape, the same pattern
  as the existing `_escapeOData` in `lists.ts`, where prettier and eslint `quotes` disagree.
- 2026-10-21, first tenant run (integration testing spec, Phase 0a): `npm run provision --
  --package tests/fixtures/templates/smoke.json --twice --strict` against a reserved team site,
  signed in with cookies as PP365's e2e user. Both runs exited 0 with 0 errors and 0 warnings
  logged. Read back through REST afterwards:
  - Site groups `SPJS Itest Readers` (12) and `SPJS Itest Owners` (13) created once; the second
    run updated them (same ids, no duplicates). `{sitename}` resolved in the description;
    `AllowMembersEditMembership: false` applied.
  - Web role assignment `SPJS Itest Readers` -> `Lese` (the English `Read` resolved by role
    type on a Norwegian site).
  - The list got unique permissions with the inherited assignments copied, plus
    `SPJS Itest Owners` -> `Full kontroll` and `SPJS Itest Readers` -> `Lese`.
  - **Defect: the group owner is not applied.** Both groups are still owned by the signed-in
    user (principal 7, `PrincipalType` 1), although the log says "Owner of site group ... set to
    {associatedownergroupid} (3)" and "... set to SPJS Itest Owners (13)". `SetUserAsOwner`
    answers success for a group id and changes nothing. A user owner was not tested.
- 2026-10-21, after the integration testing spec's Phase 0b fix (a group owner through CSOM
  `Group.Owner` over `ProcessQuery`, then read back): the same run, twice, exit 0, no errors or
  warnings, each line printed once. Read back: `SPJS Itest Readers` (12) is owned by the
  associated owner group (3, `PrincipalType` 8) and `SPJS Itest Owners` (13) by itself. The unit
  tests cover a user owner (`SetUserAsOwner`), a bare id that is a group, and an owner SharePoint
  ignores (logged as an error).
- 2026-10-21, Verified by `tests/integration/security.test.ts` on a scratch subweb, twice in a
  row in the full run of the integration suites: three site groups named through `{sitename}`
  (owners: the associated owner group, and a group that owns itself), the join/leave settings,
  `ClearExistingMembers` leaving exactly the template's member, web role assignments with `Read` /
  `Contribute` resolved by role type on a Norwegian site, the "already unique" branch, and on a
  nested web that inherits: inheritance broken once with the assignments copied. A second apply
  changes nothing. List `Security` (break, copy, assignments to the scratch web's groups) is in
  `tests/integration/lists.test.ts`; that `breakRoleInheritance` is sent on every run is pinned
  there (SharePoint ignores it today). The suite also found that a `$select` left on the shared
  web made a later run lose `{sitename}`, fixed in the integration testing spec (D4).

## Open Questions

- Answered and fixed 2026-10-21: `SetUserAsOwner` with a group id does **not** make the group the owner; it
  returns success and leaves the owner unchanged, so the handler logs a success that did not
  happen. A group owner needs CSOM (`Group.Owner = <group>` over `_vti_bin/client.svc/
  ProcessQuery`), as PnP PowerShell's `Set-PnPGroup -Owner` does. Tracked as an engine defect in
  [integration testing](./integration-testing.md) (D4, Phase 0b): set the owner through
  ProcessQuery when the principal is a group, read the owner back, and log an error when it
  differs.

## Rollout Notes

- Additive: no change to existing template sections.
- Consumers with a custom `handlerSort` must add `Security` before `Lists` for list permissions
  to reference groups from the same template.
- SharePoint makes the provisioning user the owner of a new group. Set `Owner` (for example to
  `{associatedownergroupid}`) when the group should not depend on that user.
