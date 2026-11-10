---
name: spjs-handlers
description: 'How to change or add a provisioning handler in sp-js-provisioning, and what a change must come with. Use when asked to "fix the Lists handler", "add a handler", "change what a handler does", "support a new template property", "why is the token not replaced", "should this throw or log", "make the re-run safe", "a handler stops the whole template", "add to schema.ts". Encodes the handler contract (HandlerBase, ProvisionObjects, the context, DefaultHandlerMap and DefaultHandlerSort), throw versus log_error, the token reach per member, the PnPjs 4 and ES5 pitfalls, the generated schema.json, and the tests a change needs (unit, contract, integration, verify). Pair with spjs-testing for the tests and spjs-release to ship it.'
---

# Changing a handler

Authoritative description: `docs/ARCHITECTURE.md` (handlers, tokens, context) and `README.md`
(the template format). This skill is the operational summary; where it and the code disagree, the
code wins and the doc is fixed in the same change.

## Read first

- `docs/ARCHITECTURE.md`, "Handler Architecture" and the handler's own section
- The handler's row in `docs/specs/integration-testing.md`, "Per-handler matrix" (token reach,
  read-back, pins), and its re-apply spec when it has one (`docs/specs/lists-re-apply.md`,
  `sitefields-re-apply.md`, `navigation-re-apply.md`)
- `src/handlers/<handler>.ts`, `tests/unit/<handler>.test.ts`, `tests/integration/<handler>.test.ts`

## The contract

| Piece | Where | Rule |
| --- | --- | --- |
| Handler | `src/handlers/<name>.ts` | `extends HandlerBase`; `ProvisionObjects(web, section, context)`; `scope_started()` first, `scope_ended(error?)` last |
| Registration | `src/handlers/exports.ts` | `Handler` union, `DefaultHandlerMap`, `DefaultHandlerSort` (the run order: Taxonomy -1, SiteFields 0, ContentTypes 1, Features 2, Security 2.5, Lists 3, Files 4, CustomActions 5, ComposedLook 6, ClientSidePages 7, PropertyBagEntries 8, Navigation 9, WebSettings 10, Hooks 11) |
| Template | `src/schema.ts` | TSDoc on every property; `npm run build:schema` regenerates `schema.json` (CI fails when it is stale) |
| Context | `src/provisioningcontext.ts` | Shared between handlers (lists, views, site fields, content types, what this setup created); key formats are a contract |
| Errors | `WebProvisioner` | A thrown error becomes `ProvisioningError(handler)` and stops the template; later handlers do not run |

## Throw or log

- **Throw** when nothing else in the section can go on (a failure that leaves the rest invalid). PP365
  shows it as a failed step and the setup stops.
- **`log_error` and carry on** for a failing item (a site column, a list field, a view field, a row,
  a page control, a role assignment): the rest of the section is still applied. PP365 1.15 shows
  these in the advanced log and ends with "completed with warnings". Never `log_info` a failure: it
  is invisible to `--strict`, the suites and PP365.
- `log_warn` for what the template asked and the site already has differently on purpose (a
  content type removal refused as in use). `log_info` for progress.
- Whatever you throw or log is read by people: `readableErrorMessage` (`src/provisioningerror.ts`)
  turns a JSOM `{ sender, args }`, a REST `message.value` or a string into words; name the item
  (`Navigation node '<Title>' (<Url>) was refused by SharePoint: ...`).

## Tokens

- `TokenHelper.replaceTokens` (`{listid:}`, `{webid}`, `{sitename}`, `{parameter:}`, ...) and
  `replaceUrlTokens` (`{site}`, `{sitecollection}`, `{hosturl}`, ...) are applied per member, not
  per section. The matrix in the integration spec says which member gets which; `verify`
  (`debug/lib/verify.ts`, `expectations()`) mirrors it. Change one, change the other, the matrix and
  `docs/ARCHITECTURE.md` ("Where tokens are applied").
- An unresolved known token is an error to log, naming the item (`unresolvedTokens()`), not a
  string to send to SharePoint.

## Typical task: change what a handler does

1. Unit test first (`tests/unit/<handler>.test.ts`): a fake web that records the calls, or the
   contract transport (`tests/helpers/transport.ts`, `createWireWeb`) to pin the requests on the
   wire (`tests/contract/wire.test.ts`). Assert the log level of a failure.
2. The change, in TypeScript 4.0 syntax (ES5 output): no `satisfies`, no `Awaited`; an `Error`
   subclass restores its prototype.
3. `src/schema.ts` TSDoc and `npm run build:schema` when the template changes; a fixture or a
   `NOT_COVERED` reason for a new property (`tests/unit/fixture-coverage.test.ts`).
4. The integration suite and fixture: read back what SharePoint stored; `applyTwice` proves the
   re-run and runs `verify`; teach `verify` the rule when the change alters what a re-apply makes
   true (with a unit test in `tests/unit/verify.test.ts`).
5. `README.md` (template format), `docs/ARCHITECTURE.md`, a CHANGELOG line under `[Unreleased]`
   (what a consumer notices: PP365 runs this in every project setup and package import).
6. `npm run build`, `npm test`, `npm run typecheck:tests`, `npm run lint:check`,
   `npm run check:schema`; then the suite live with the user's go-ahead.

## Typical task: add a handler

As above, plus: the `Handler` union, `DefaultHandlerMap`, a `DefaultHandlerSort` number (what must
exist before it runs?), the `Schema` property, `tests/contract/orchestration.test.ts` (its handler
list), `debug/lib/cli.ts` `OPT_IN_HANDLERS` if it writes beyond the site (as Taxonomy does), a reader in
`debug/lib/readback/`, a fixture, a suite, `verify` expectations, and the matrix row.

## Pitfalls

- PnPjs 4's `select()` sets the query on the queryable itself (2026-10-21): a handler that reads
  through the web it was given leaves `$select` for the next handler. `WebProvisioner` hands each
  handler a fresh `Web(this.web)`; inside a handler, build a new queryable for each read.
- A re-apply on a site in use must not lose data (1.5.0): update in place, never delete and
  recreate a field, a view or a node; act on what this setup created (`context.created*`) when the
  template asks to remove.
- What REST cannot do goes through CSOM over HTTP (`src/util/csom.ts`: `CsomRequest`,
  `executeQuery`, `CsomError`), not JSOM: ContentTypes, Taxonomy, PropertyBagEntries and the
  Security group owner do; no handler uses JSOM (`docs/specs/jsom-exit.md`). Taxonomy writes to the
  tenant-wide term store: tests write only to the group `SPJS Itest` and never delete from the
  store. Unit-test a
  CSOM request with `FakeCsom` and `parseCsom` (`tests/helpers/csom.ts`): assert the actions as
  lines, not the XML.
- CSOM answers 200 with the error in the body, and stops at the first failing action without
  undoing the ones before it: read first, group one object's changes in one request, so a re-run
  picks up where a failed one stopped.
- Handlers mutate their input (WebSettings writes values back): a caller that applies twice passes
  a fresh copy.
- `@pnp/*` is a peer dependency and one physical copy (`tests/unit/package.test.ts`).

## Verification checklist

- [ ] A unit or contract test fails without the change and passes with it
- [ ] Failures are `log_error` (or thrown), named in words; no failure at Info
- [ ] Tokens per member as the matrix says; `verify` agrees
- [ ] `schema.json` regenerated; README, ARCHITECTURE and CHANGELOG updated
- [ ] `npm run build` (TypeScript 4.0.7, ES5) and the tsx-based tests both green

## Do not

- Run `provision`, `cleanup` or the integration suites against a site the user did not name, or
  without the user's go-ahead; never print a cookie, token or certificate, or commit `.env`.
- Delete an assertion, skip a test without a reason, or lower a coverage floor to get green.
- Add `@pnp/*` to `dependencies`, or syntax newer than TypeScript 4.0 to `src/`.
- Swallow a failure silently, or change a handler's re-apply behaviour without its spec.
