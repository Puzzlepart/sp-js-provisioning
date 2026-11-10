---
name: spjs-testing
description: 'How sp-js-provisioning is tested and how to write, run and fix its tests. Use when asked to "add a test", "run the tests", "test the handler", "test against the site", "test against the tenant", "integration test", "check:tenant", "why does the suite fail", "the second run changed", "coverage", "flaky test", "test the Lists handler", "add a fixture", "verify found differences". Encodes the layers (unit, contract, integration against scratch webs, browser-only JSOM handlers), the commands, the harness (withScratchWeb, applyTwice, apply, read-backs, verify, loadFixture), the fixture rules, the site guard, NoScript, and the failure triage. Pair with spjs-provision to apply a template by hand.'
---

# Testing sp-js-provisioning

Authoritative description: `docs/TESTING.md`. This skill is the operational summary; where it and
the code disagree, the code wins and the doc is fixed. The plan: `docs/specs/integration-testing.md`.

## Read first

- `docs/TESTING.md` (layers, commands, tenant, harness, rules, triage)
- `tests/integration/security.test.ts` with `debug/lib/readback/security.ts` and
  `tests/fixtures/templates/security.json`: the reference suite
- The handler's row in the spec's "Per-handler matrix" (token reach, read-back, pins)

## Layers

| Layer | Location | Runs | Tenant |
| --- | --- | --- | --- |
| Unit | `tests/unit/*.test.ts` | `npm test`, CI | no |
| Contract | `tests/contract/*.test.ts` (real PnPjs, `createWireWeb`: the requests on the wire) | `npm test`, CI | no |
| Integration | `tests/integration/<handler>.test.ts` | `npm run test:integration -- <Handler>` | yes: scratch webs below `INTEGRATION_SITE_URL` |
| Everything | `tests/integration/everything.test.ts` (`all.node.json`, `verify --exact`) | `npm run test:integration -- all` | yes |
| Packages | `tests/integration/packages.test.ts` (prosjektportalen-hosting, an exact list of what Node cannot do) | `npm run test:integration -- packages` | yes |
| Browser only | Taxonomy, ContentTypes, Files, PropertyBagEntries | suites skip until the JSOM exit | — |

## Running

- Offline: `npm test`, `npm run typecheck:tests`, `npm run lint:check`, `npm run test:coverage`;
  `npm run test:tasks` after a change to `.tasks/publish-test-report.js` (copy it to
  prosjektportalen365 too: the file is shared byte for byte).
- Tenant ready? `npm run check:tenant`. Then `npm run test:integration -- Lists` (reachability runs
  first). `--keep` keeps the scratch web; `npm run cleanup -- --webs` lists leftovers older than
  2 h (`--older-than 0m` for younger ones, only while no other run is active; `--apply` removes
  them).
- One suite file directly, without the runner (no reachability first; `INTEGRATION_KEEP=true`
  instead of `--keep`):
  `npx tsx --tsconfig tests/tsconfig.json --test --test-concurrency=1 tests/integration/<name>.test.ts`.
  Tenant runs are serial: one run per reserved site at a time.
- Lists and Lists (DataRows) tests build on the first test's lists: run those files whole, not
  with a `--test-name-pattern` that skips the first test.

## Typical task: add or extend a handler suite

1. Fixture `tests/fixtures/templates/<name>.json`: `SPJS Itest ...` titles, `SPJSItest...` internal
   names, fixed lower-case GUIDs, engine tokens only where the handler replaces them, `@@user@@` /
   `@@hookUrl@@` for run-time values, site group titles that start with `{sitename}`.
2. Reader `debug/lib/readback/<name>.ts`: what the handler provisioned as plain data, without ids,
   timestamps or other volatile members.
3. Suite: `describe('<Handler>', { skip: skipReason() }, ...)`, `withScratchWeb('<name>')`, then
   `applyTwice(conn, () => loadFixture('<name>'), read)` and assertions on the read-back. Tenant
   values (role names, group titles) are read in the test. `applyTwice` also runs `verify` on the
   sections that ran and fails on a finding (`assertVerified`; call it yourself after a B variant).
4. A defect: pin it with `it(name, { todo: 'why' }, ...)` in the suite. A unit test that exposes it
   is marked `{ todo: 'why' }` too, or lands with the fix (a plain failing test turns the required
   CI check red). Fix it in its own change with a CHANGELOG line, removing the `todo`s there.
5. `npx prettier --write`, `npm run typecheck:tests`, `npm test`, then the suite live.

## Pitfalls

- PnPjs 4's `select()` mutates the queryable (found 2026-10-21): a `$select` left on a shared web
  made the next run read only that property, so `{sitename}` stayed literal and groups named
  "{sitename} Readers" were created. Readers use `conn.web` (fresh on every access).
- A web that inherits permissions has no associated groups: `{associated...groupid}` fails there.
- NoScript is on by default on modern sites: no ScriptLink, property bag or master page writes.
- `createDefaultAssociatedGroups` answers 403 under NoScript although it creates the groups; the
  harness checks the groups instead.
- Site groups belong to the site collection and outlive a deleted web: they are removed with it
  when their title starts with its title (`{sitename} ...`), never by the stamp alone (two suites
  started in the same second share one).
- A web with a subweb cannot be deleted: delete nested webs in `finally`.
- Several full runs in a short time get the account throttled (2026-10-21, fifth run in three
  hours): `# SharePoint: ... answered 429` lines, a suite many times slower, odd 404s. Rest, then
  rerun the failed suites; do not loosen an assertion for it.
- The cookie session lasts about two days although the cookie says otherwise; a 403 on the first
  read means renew it (`npx playwright test --project=setup` in `prosjektportalen365/e2e`).
- `verify found differences` with a green read-back: `verify`'s rule for that member does not
  match what the handler does (fix `debug/lib/verify.ts` with a unit test in
  `tests/unit/verify.test.ts`), or the web really differs (a defect the read-back missed).
- SharePoint refuses a request whose URL passes its length limit (about 260 characters of path,
  2026-11-10): a scratch web's path is in every request and `getFileByServerRelativePath` names it
  twice, so keep suite names short (`pkg forskningsprosjekt project`, not the whole case label)
  and reach files through their list where you can.
- SharePoint stores CAML normalised (`<FieldRef Name= "X"/>` becomes `<FieldRef Name="X" />`):
  compare a view query through `verify`, not as text.
- The packages suite holds each case to an exact list (`tests/integration/packages.expected.ts`,
  what each list lacks). A new logged error or finding fails as `unexpected`, a fixed one as
  `missing` (remove its entry). One case: `npm run test:integration -- packages
  --test-name-pattern='pp-enkel|pp-forskning'` (the pattern matches the case's name). It reads the
  prosjektportalen-hosting checkout next to this one, uncommitted changes included.

## Verification checklist

- [ ] The suite asserts the read-back, not only a clean run
- [ ] `applyTwice` passes (second read-back equal to the first, and no `verify` finding)
- [ ] A changed handler fixture: `npm run build:fixtures` (else `tests/unit/fixtures.test.ts` fails)
- [ ] Known defects are `todo`, with a reason
- [ ] No tenant URL, cookie, token or certificate in any committed or new file
      (`npm run check:secrets` scans them all, docs and specs too); no user name either (it is
      not scanned: use `@@user@@`, and `<tenant>` when recording a run)
- [ ] `npm test`, `npm run typecheck:tests`, `npm run lint:check` green
- [ ] Nothing left behind: reachability's "reports leftover scratch webs" lists none of yours, or
      `npm run cleanup -- --webs --older-than 0m` (while no other run is active) plans nothing

## Do not

- Run the integration suites, `provision` or `cleanup --apply` without the user's go-ahead, or
  against a site other than `INTEGRATION_SITE_URL`.
- Use `--allow-site` or `cleanup --apply --all` while another run may be active.
- Delete an assertion, skip without a reason, add retries or lower a coverage floor to get green.
- Print or commit a cookie, token, certificate or `.env`.
