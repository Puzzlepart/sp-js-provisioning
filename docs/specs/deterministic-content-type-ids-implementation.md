# Implementation Plan: Deterministic Content Type IDs

Status: Draft

Related:

- Spec: ./deterministic-content-type-ids.md
- PR: https://github.com/Puzzlepart/sp-js-provisioning/pull/6
- Decision record: ../adr/jsom-explicit-content-type-ids.md

## Review Findings

1. TypeScript build fails.
   - `src/handlers/contenttypes.ts` caches a partial `IContentType` without required `Description` and `Group` fields.
   - `SP.ContentTypeCreationInformation` in the installed `@types/sharepoint` package does not expose `set_id(...)`.

2. Context initialization is not awaited.
   - `ProvisionObjects` calls `this._initContext(web)` without `await`.
   - The PR relies on `this.context.contentTypes` to decide whether to create or update, so this can turn existing content types into attempted duplicate creations.

3. Existing-site behavior is not yet verified.
   - The PR body confirms the fresh-site scenario.
   - The existing-template smoke test is still unchecked.

4. Local working tree has unrelated package lock noise.
   - `package-lock.json` is modified locally.
   - It is not part of the PR diff and should be kept separate unless intentionally handled.

## Task Plan

### 1. Fix Context Initialization

- Change `this._initContext(web)` to `await this._initContext(web)`.
- Confirm `context.contentTypes` is populated before processing the first configured content type.

### 2. Fix Context Cache Shape

- Ensure cached content type entries satisfy `IContentType`.
- Preserve the configured values where available.
- Preserve discovered SharePoint values when reading existing content types in `_initContext`.

### 3. Resolve `set_id(...)` Typing

Choose one approach:

- Local cast: narrow the workaround to the `ContentTypeCreationInformation` instance.
- Module augmentation: add the missing method to the SharePoint type surface used by the project.
- Dependency/type update: only if it does not create broader churn.

Preferred initial approach: local cast, because it keeps the fix scoped to the platform typing gap.

### 4. Re-run Local Checks

- `npm run build`
- `git diff --check origin/main...HEAD`

### 5. Smoke Test SharePoint Behavior

- Fresh site with explicit content type IDs.
- Existing site with the same template re-run.
- Existing site template that references content types by name without explicit IDs, if the project supports that path.
- List binding verification for content types attached to document libraries or lists.

### 6. Update PR Metadata

- Mark the existing-template smoke test complete if it passes.
- Note how the `set_id(...)` type gap was handled.
- Keep `package-lock.json` out of the PR unless intentionally changed.

## Suggested Code Shape

The implementation should keep the behavior easy to audit:

- `ProvisionObjects` awaits context initialization.
- `processContentType` resolves the intended content type ID.
- `processContentType` branches clearly between existing content type update and new content type creation.
- `createContentType` is the only place that sets the explicit ID.
- `createContentType` documents why `ParentContentType` is not set.

## Done Checklist

- [ ] Build passes.
- [ ] Whitespace diff check passes.
- [ ] Fresh-site smoke test passes.
- [ ] Existing-site smoke test passes.
- [ ] PR body verification section is updated.
- [ ] Unrelated `package-lock.json` change is resolved or explicitly excluded.
