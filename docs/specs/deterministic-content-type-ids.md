# Spec: Deterministic Content Type ID Provisioning

Status: Implemented

Owner: okms

Related:

- PR: https://github.com/Puzzlepart/sp-js-provisioning/pull/6
- Implementation plan: ./deterministic-content-type-ids-implementation.md
- Decision record: ../adr/jsom-explicit-content-type-ids.md

## Problem

Provisioning content types with deterministic IDs fails on fresh SharePoint sites when content type creation supplies both an explicit content type ID and an explicit parent content type.

SharePoint rejects that combination with:

```text
parameters.Id, parameters.ParentContentType cannot be used together. Please only use one of them.
```

This prevents a provisioning template from creating site content types with stable IDs, which in turn affects later list content type bindings that depend on those IDs.

## Goals

- Create content types with deterministic IDs on fresh sites.
- Preserve update behavior for content types that already exist.
- Keep list content type bindings working after site content types are provisioned.
- Keep the TypeScript build passing.

## Non-Goals

- Redesign the full content type provisioning handler.
- Change the public provisioning schema.
- Replace JSOM or PnP usage outside the content type creation path.

## Current Behavior

On a fresh site, content type creation fails before any configured content types are created if SharePoint receives both:

- the explicit content type ID from the provisioning template
- the explicit parent content type object derived from that ID

On existing sites, the handler must still identify an existing content type and update it instead of attempting to create a duplicate.

## Desired Behavior

When a provisioning template contains a content type with an explicit `ID`:

1. The handler creates the content type with that exact ID if it does not already exist.
2. The handler does not also set `ParentContentType` during creation.
3. SharePoint infers the parent from the content type ID structure.
4. Later provisioning runs update the existing content type instead of recreating it.
5. List bindings can attach the provisioned site content type and receive an inherited list content type ID.

When a provisioning template references an existing content type by `Name` and omits `ID`:

1. The handler resolves the ID from the initialized provisioning context.
2. The handler updates the existing content type.
3. The handler does not create a new content type.

## Requirements

- Context initialization must complete before create-versus-update decisions are made.
- The in-memory content type context must remain valid according to the project `IContentType` type.
- The implementation must keep deterministic IDs intact across site content type creation and list binding.
- Field references must still be added, updated, and reordered after the site content type exists.

## Platform Constraints

- SharePoint does not allow content type creation parameters to include both an explicit ID and `ParentContentType`.
- The installed `@types/sharepoint` package does not expose `set_id(...)` on `SP.ContentTypeCreationInformation`.
- Existing provisioning templates may refer to content types by explicit ID or by name.

## Acceptance Criteria

- TypeScript build passes.
- Diff whitespace check passes.
- Fresh-site content type creation succeeds with deterministic IDs.
- Existing-site provisioning updates existing content types without duplicate creation attempts.
- List content type bindings still attach inherited content types correctly.
- Unrelated `package-lock.json` changes are either intentionally handled or removed from the PR.

## Verification Plan

- `npm run build`
- `git diff --check origin/main...HEAD`
- Public-safe smoke fixture: `tests/smoke/deterministic-content-type-ids/`
- Fresh SharePoint site smoke test with configured content types using explicit IDs.
- Existing SharePoint site smoke test with a provisioning template that contains explicit-ID content types.

## Resolved Questions

- The `set_id(...)` type gap is handled with a local cast scoped to the content type creation path.

## Rollout Notes

Do not merge until both fresh-site and existing-site behavior are verified. The existing-site path is the most likely regression because the PR introduced a create-versus-update branch based on initialized context state.
