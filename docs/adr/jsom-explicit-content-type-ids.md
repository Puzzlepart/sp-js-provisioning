# ADR: Use JSOM Creation Information for Explicit Content Type IDs

Status: Proposed

Date: 2026-05-11

Related:

- Spec: ../specs/deterministic-content-type-ids.md
- Implementation plan: ../specs/deterministic-content-type-ids-implementation.md
- PR: https://github.com/Puzzlepart/sp-js-provisioning/pull/6

## Context

The provisioning handler needs to create SharePoint site content types with deterministic IDs from the provisioning template.

SharePoint rejects creation requests that set both an explicit content type ID and `ParentContentType`:

```text
parameters.Id, parameters.ParentContentType cannot be used together. Please only use one of them.
```

The parent content type is already encoded in a valid SharePoint content type ID, so setting both values is redundant and invalid.

## Decision

Create new site content types through `SP.ContentTypeCreationInformation` with the explicit ID set and without setting `ParentContentType`.

Existing content types should continue to be loaded by ID from the web content type collection and updated in place.

## Consequences

- Fresh-site provisioning can create content types with deterministic IDs.
- The parent relationship is inferred by SharePoint from the ID structure.
- The code depends on a JSOM method for setting the content type ID that is missing from the installed `@types/sharepoint` package.
- The implementation must keep any typing workaround narrow and well documented.
- Context initialization becomes more important because create-versus-update behavior depends on knowing which content types already exist.

## Alternatives Considered

- Continue using PnP `web.contentTypes.add(...)`.
  - Rejected because the observed behavior sets both explicit ID and parent content type, which SharePoint rejects on fresh sites.

- Set `ParentContentType` and omit explicit ID.
  - Rejected because deterministic content type IDs are required by the provisioning template and downstream list bindings.

- Generate IDs after creation.
  - Rejected because SharePoint content type IDs are identity values and must be stable before bindings depend on them.
