# ADR: Reach SharePoint Through CSOM Over HTTP Where REST Cannot

Status: Accepted

Date: 2026-11-10

Related:

- Spec: ../specs/jsom-exit.md (J1, J2 and the spike)
- Supersedes: ./jsom-explicit-content-type-ids.md

## Context

Four handlers used SharePoint's JSOM (`SP.ClientContext`, loaded into the page by `spfx-jsom`):
ContentTypes, Taxonomy, PropertyBagEntries and the classic web parts of Files. JSOM cannot load in
Node, so none of them had an integration test, and the global `SP.ClientContext` that
PropertyBagEntries and Files used does not exist on an SPFx page, so both refused to run in PP365.

What REST cannot do, the spike of 2026-11-10 showed on the reserved site:

- `POST _api/web/contenttypes` ignores the id it is given and makes a child of Item with an id of
  its own; list bindings and child content types need the template's id.
- Field links, their `Required`/`Hidden` and their order, and `Update(true)` to push a change to
  what inherits from a content type, have no REST endpoint that does all of it.
- The term store's REST API answers 403 to a browser session on tenants that enforce "OAuth only".

`_vti_bin/client.svc/ProcessQuery`, the endpoint JSOM itself posts to, did all of it: a content type
with the asked id (also a child), field links, flags, `Reorder`, `Update(true)`, and reads of the
term store.

## Decision

Writes that REST cannot do are sent as CSOM requests over HTTP: XML posted to the web's
`ProcessQuery` through the web's own PnPjs pipeline. One small builder in `src/util/csom.ts`
(`CsomRequest`: object paths, actions, typed parameters; `executeQuery`: the results by action id,
SharePoint's error as a `CsomError`) writes every request; no dependency is added. Reads use REST
where it has them.

A content type is created with `ContentTypeCreationInformation` holding the explicit `Id` and no
`ParentContentType`, as the superseded ADR decided for JSOM.

## Consequences

- The handlers run in Node and on an SPFx page alike, signed in as the web is, so they get
  integration suites, and `spfx-jsom` can leave the package once all four are moved.
- Permissions and behaviour are what JSOM had: the same endpoint and object model.
- The request XML is ours to keep right: unit tests read it back (`tests/helpers/csom.ts`,
  `parseCsom`, `FakeCsom`) and the suites run it against SharePoint.
- CSOM answers 200 with the error in the body and stops at the first failing action without
  rolling back what came before it; a handler groups a content type's changes in one request and
  reads the state first, so a re-run picks up where a failed one stopped.
- Type ids (`{168f3091-…}` for `ContentTypeCreationInformation`, …) are SharePoint's and fixed; they
  are named in the code where they are used.

## Alternatives Considered

- REST only: refused for content types (the id) and field link order; the term store's REST API is
  refused on "OAuth only" tenants.
- Keep JSOM and stub it in Node: no integration tests, and PropertyBagEntries and Files still could
  not run under SPFx.
- A CSOM library from npm: none fits ES5 output and PnPjs's pipeline; the requests the engine needs
  are few and regular.
