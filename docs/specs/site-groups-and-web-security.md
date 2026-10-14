# Spec: Site Groups And Web Security

Status: Implemented (unit tested; tenant verification pending)

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
- Manual, on a test tenant with `npm run provision` (see `debug/README.md`):
  - fresh group-connected site: groups created, owner set, members added, list permissions for
    the new groups applied;
  - second run on the same site: no duplicates, no errors;
  - group owner set to another group (`{associatedownergroupid}`) and to the group itself.

## Verification Result

- 2026-10-14: `npm run build`, `npm test` (30/30, of which 13 new) and `npm run typecheck:tests`
  pass. Prettier is clean; eslint reports only the `"''"` OData escape, the same pattern
  as the existing `_escapeOData` in `lists.ts`, where prettier and eslint `quotes` disagree.
- Tenant verification with `npm run provision` is still pending.

## Open Questions

- `SetUserAsOwner` with a group id (a group owning a group) is expected to work through REST but
  has to be confirmed on a tenant. Until then a failure is logged and the group keeps the
  provisioning user as owner.

## Rollout Notes

- Additive: no change to existing template sections.
- Consumers with a custom `handlerSort` must add `Security` before `Lists` for list permissions
  to reference groups from the same template.
- SharePoint makes the provisioning user the owner of a new group. Set `Owner` (for example to
  `{associatedownergroupid}`) when the group should not depend on that user.
