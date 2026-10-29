# Testing

This is the authoritative description of how sp-js-provisioning is tested. `AGENTS.md` and the
`spjs-testing` skill summarise it; where they and this page disagree, fix them. Where this page and
the code disagree, the code wins and this page is fixed in the same change. The plan behind it is
[the integration testing spec](./specs/integration-testing.md).

## Layers

| Layer | Where | Needs a tenant | Runs |
| --- | --- | --- | --- |
| Unit | `tests/unit/*.test.ts` (node:test through tsx; real PnPjs with fake webs or a recording web) | no | `npm test`, and in CI (`ci.yml`, job "Build and offline tests") with JUnit and coverage floors |
| Contract | `tests/contract/*.test.ts` (Phase 2: real PnPjs, a fake transport, the wire shape of each handler) | no | `npm test`, CI |
| Integration | `tests/integration/*.test.ts`, one suite per handler, against scratch webs below the reserved site | yes | `npm run test:integration` locally; `integration.yml` (Phase 3, dispatched by hand behind a protected environment) |
| Packages | `tests/integration/packages.test.ts` (Phase 2: the prosjektportalen-hosting packages) | yes | `npm run test:integration -- packages` |
| Browser only | Taxonomy, ContentTypes, Files and PropertyBagEntries run through JSOM | yes, in a browser | their suites skip until the JSOM exit (spec, D3) |

## Which test do I write?

| You changed | Write |
| --- | --- |
| Logic with no SharePoint call (tokens, parsing, plans) | a unit test |
| What a handler sends (a URL, a body, the order of requests) | a unit test with a recording web (`tests/unit/webprovisioner.test.ts`, `csom.test.ts`); in Phase 2 a contract test |
| How a handler reacts to a failure (log level, `ProvisioningError`) | a unit test with a throwing fake and `captureLogs` (`tests/helpers/logs.ts`) |
| What ends up on a site | the handler's integration suite: fixture, read-back, `applyTwice` |
| A defect found on a site | first the test that fails because of it (unit if possible, else integration), then the fix |

## Commands

| Command | Does |
| --- | --- |
| `npm test` | Unit tests, offline, seconds |
| `npm run typecheck:tests` | Type-checks the tests with `src/` and `debug/` |
| `npm run test:ci` / `npm run test:coverage` | Unit tests with spec + JUnit to `test-results/unit.xml` / with source-mapped coverage floors |
| `npm run lint:check` | eslint on `src/` and prettier on all TypeScript, without rewriting |
| `npm run check:pack`, `npm run check:secrets`, `npm run check-skills` | The package holds `lib/`; no tenant URL, cookie, token or certificate committed; the agent skills are valid |
| `npm run check:tenant` | Is the tenant ready: sign-in, site collection admin, term store, custom script, a scratch web created and deleted, leftovers. Strict: a missing site or sign-in fails instead of skipping |
| `npm run test:integration` | Reachability in a process of its own, then every handler suite, one file at a time; the suites do not run when reachability fails |
| `npm run test:integration -- Lists Security` | Those suites (handler names, any case) |
| `npm run test:integration -- Lists --keep` | Keep the scratch webs to look at them |
| `npm run test:integration:ci` | Strict (a missing site or sign-in fails), spec + JUnit to `test-results/reachability.xml` and `test-results/integration.xml`; both, and the log, pass through `redact()` |
| `npm run cleanup -- --webs` / `--prefix 'SPJS Itest'` | What a crashed run left behind (scratch webs and their groups older than 2 h / prefixed artefacts); a plan unless `--apply` |
| `npm run cleanup -- --webs --older-than 0m` | The same right after a run or `--keep` (webs younger than 2 h are left for a run that may be active); only while no other run is |
| `npm run provision -- --package <path> --twice --strict` | Apply a template or package by hand (`debug/README.md`) |

## The tenant

- **A reserved site**, `INTEGRATION_SITE_URL` in `.env`: a site nothing else writes to (a
  group-connected team site works). The tests write only to scratch webs below it; `provision` and
  `cleanup` refuse any other site unless `--allow-site` repeats it.
- **Subsites enabled** in the tenant (SharePoint admin center, Settings, Classic settings page,
  Subsite creation: Enable subsite creation for all sites). Each suite creates a scratch web.
- **The account** is a site collection administrator of the reserved site: cleanup deletes site
  groups, including one that owns itself. Locally that is PP365's e2e user, whose session
  (`prosjektportalen365/e2e/.auth/user.json`) lasts about two days: renew it with
  `npx playwright test --project=setup` in `prosjektportalen365/e2e`. In CI it is an app with a
  certificate (`CI_TENANT`, `CI_CLIENT_ID`, `CI_CERT_BASE64`; `Sites.FullControl.All` and
  `TermStore.Read.All` with admin consent) acting for `INTEGRATION_USER`.
- **Custom script** (NoScript, `DenyAddAndCustomizePages`) is blocked on modern sites by default.
  `check:tenant` reports it; the suites that need it (ScriptLink custom actions, property bag) skip
  with that reason. An SPO admin can allow it per site, but SharePoint blocks it again after a day.

## The harness

| Piece | File | Does |
| --- | --- | --- |
| Sign-in | `debug/lib/auth.ts`, `debug/lib/env.ts` | Cookies (local) or certificate (CI), from `.env`; each attempt at a request bounded at 120 s (`REQUEST_TIMEOUT`), a throttled one (429, 503) waiting `Retry-After` and tried again, one not answered in time not sent again |
| Connection | `debug/lib/connect.ts` | `connect(url)`: a client, what `_api/site` and `_api/web` say, and the SPFx context stand-in the engine is set up with; `assertWritable` (the site guard). `conn.web` is a fresh queryable on every use. |
| Scratch webs | `debug/lib/scratch.ts` | `SPJS Itest <suite> <stamp>` at `spjs-itest-<suite>-<stamp>` (UTC `yyMMddHHmmss`), unique permissions, its own associated groups; deleted with the site groups whose title starts with its title (its associated groups and the fixtures' `{sitename} ...` groups) |
| Applying | `debug/lib/apply.ts` | `apply(conn, template, options)` returns `{ order, skipped, errors, warnings, error }` and never throws |
| Reading back | `debug/lib/readback/<handler>.ts` | What a handler provisioned, as plain data without ids or timestamps |
| Suites | `tests/integration/helpers.ts` | `withScratchWeb`, `applyTwice`, `assertClean`, `skipReason` |
| Fixtures | `tests/fixtures/templates/*.json`, `tests/fixtures/fixtures.ts` | `loadFixture(name, { user, hookUrl })` |
| Runner | `tests/integration/run.ts` | Maps names to suite files; runs reachability first in its own process (`node --test` sorts the files it is given), then the suites, one file at a time |

A suite looks like this (`tests/integration/security.test.ts` is the reference):

```ts
describe('Security', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('security')
  it('creates groups ... and changes nothing on a second apply', async () => {
    const { conn } = scratch()
    const { afterFirst } = await applyTwice(conn, () => loadFixture('security', { user }), (c) =>
      readSecurity(c, titles)
    )
    assert.deepEqual(afterFirst.groups.map((group) => [group.Title, group.Owner]), [...])
  })
})
```

## Rules

1. **The read-back is the proof.** Assert what SharePoint stored, read back through REST, against
   the fixture. A green apply with an empty error log is a second signal only: SharePoint can accept
   a request and change nothing (a group owner through REST's `SetUserAsOwner`).
2. **Apply twice.** `applyTwice` applies, reads back, applies again and reads back again: both runs
   without a `ProvisioningError` or a logged error, and the second read-back equal to the first.
3. **Fresh fixtures.** Handlers mutate their input; `loadFixture` returns a new copy every call.
4. **Fixture conventions.** Titles start with `SPJS Itest`, internal names with `SPJSItest`, site
   column groups are `SPJS Itest`, GUIDs are fixed and lower case. No tenant URL or user in a
   fixture: `@@user@@` and `@@hookUrl@@` are filled at run time. Engine tokens only where the
   handler replaces them (the spec's matrix, "Token reach"). Site group titles **start with**
   `{sitename}` (`{sitename} Readers`): they then start with the scratch web's title and are
   deleted with it. A group with `{sitename}` further in (`SPJS Itest Readers of {sitename}`) is
   left behind on every run.
5. **Tenant values are read in the test**: role definition names, group titles, web part ids.
6. **Known defects are pinned** with `it(name, { todo: 'why' }, ...)`, never hidden. No retries: a
   flaky call becomes a `todo` and an issue.
7. **Skip, or fail under strict.** A missing site or sign-in skips the suite with the reason
   (`skipReason()`); `INTEGRATION_STRICT=true` (CI) turns that into a failure, so a suite skipped
   whole cannot pass. A missing permission, user or custom script skips the test and names it.
8. **Write only where the tests may.** Suites write to their scratch web; nothing writes to the
   reserved site itself or another site.

## When a test fails

| Symptom | Likely cause | Do |
| --- | --- | --- |
| `SharePoint refused the cookies sign-in to ... (403)` in `before`, or `check:tenant` | The session expired, or the account has no access | Renew the session; check the account is site collection admin |
| `[403] Forbidden` from PnPjs inside a test, after sign-in | A write custom script blocks (ScriptLink, property bag, master page), or a permission the account lacks | `check:tenant` names custom script and site admin; skip that test with the reason, as the ScriptLink test does |
| `Nye sekundære områder er ikke tilgjengelige` / "New subsites are not available" | Subsite creation is disabled in the tenant | Enable it (see The tenant) |
| A suite skips "custom script is blocked" | NoScript on the reserved site | Expected; or have an SPO admin allow custom script |
| `the second run changed what the first provisioned` | The handler is not idempotent for that member | Compare the two read-backs in the diff; a defect, or a volatile member the reader must drop |
| `the first run logged errors` | A handler logged and carried on | The message names the handler and SharePoint's answer |
| `The operation was aborted due to timeout` | SharePoint did not answer one attempt within 120 s (`REQUEST_TIMEOUT`); seen once while creating a scratch web's groups, the next suite fine | Run the suite again; `withScratchWeb` deletes a half-made web in `after`, else `cleanup` |
| `# SharePoint: GET answered 429; waiting 120 s, trying again` (or 503), a suite many times slower than usual | SharePoint throttles the account: several full runs in a short time (seen 2026-10-21 on the fifth run in three hours: one test took 15 minutes and timed out, another got a 404 from a node it had just listed; both green when run again) | Let the tenant rest, then run the failed suites again; a failure without these lines is not throttling |
| `# scratch web ... may be left behind` | `after` could not delete the scratch web (node:test reports nothing else when `before` failed) | `npm run cleanup -- --webs --older-than 0m`, then `--apply` |
| Leftover `spjs-itest-*` webs | A run crashed before `after`, or `--keep` | reachability's "reports leftover scratch webs" lists them; `npm run cleanup -- --webs` (`--older-than 0m` for a web younger than 2 h, while no other run is active), then `--apply` |

Never delete an assertion or lower a coverage floor to make a run green.

## Per-handler status

Measured on 2026-10-27 with `npm run test:integration` against the reserved site, after the
[Lists re-apply spec](./specs/lists-re-apply.md): 51 tests (6 reachability, 45 in the suites), 48
pass, 2 `todo` (pinned defects), 1 skipped (custom script blocked), 0 fail; about 9 minutes; nothing
left behind. (Phase 1, 2026-10-21: 45 tests, 34 pass, 10 `todo`.)

Since then the [SiteFields](./specs/sitefields-re-apply.md) and
[Navigation](./specs/navigation-re-apply.md) re-apply specs (2026-10-29) turned both pins into
plain tests and added tests: SiteFields 7 (was 3), Navigation 6 (was 3). Each passed whole on
2026-10-29 (`npm run test:integration -- SiteFields Lists`: 25 tests, all pass; `-- Navigation`: 6
of 6, three times). Full run on 2026-10-29, after the review fixes: reachability 6 of 6; the suites
52 tests, 51 pass, 0 `todo`, 1 skipped (custom script blocked), 0 fail; about 9 minutes; nothing
left behind.

| Handler | Suite | Fixture | Status | Pinned (`todo`) |
| --- | --- | --- | --- | --- |
| SiteFields | `sitefields.test.ts` | `sitefields.json`, `.b.json`, `failures.json` | green (7 tests: an unchanged second apply writes nothing, B and back, an admin's description kept, a narrowing type change, a lookup to a list of the same template, a taxonomy column, a refused column) | — (re-apply and same-template `{listid:}` fixed, [spec](./specs/sitefields-re-apply.md)) |
| Features | `features.test.ts` | `features.json`, `.b.json` | green | — (re-apply fixed in Phase 1) |
| Security | `security.test.ts` | `security.json` | green (also a list granted to a group of the same template, and `ClearExistingMembers` removing a member added since) | — |
| Lists | `lists.test.ts` | `lists.json`, `.b.json` | green; the tests after the first build on its lists, so run the file whole | — (re-apply without data loss, [spec](./specs/lists-re-apply.md)) |
| Lists (DataRows) | `lists-datarows.test.ts` | `lists-datarows.json`, `.b.json` | green; the tests after the first build on its lists, so run the file whole | — (key, URL and lookup defects fixed by the same spec) |
| CustomActions | `customactions.test.ts` | `customactions.json`, `.b.json`, `customactions-script.json` | green; classic actions skip under NoScript | — (add-only behaviour asserted) |
| ComposedLook | `composedlook.test.ts` | `composedlook.json` | green | — (palette-only look fixed in Phase 1) |
| ClientSidePages | `clientsidepages.test.ts` | `clientsidepages.json`, `.b.json` | green | — |
| Navigation | `navigation.test.ts` | `navigation.json`, `.b.json` | green (6 tests: a second apply keeps ids and writes nothing, an owner's changes, B, refusals in either collection, a recycled target) | — (re-apply reconciles in place, [spec](./specs/navigation-re-apply.md)) |
| WebSettings | `websettings.test.ts` | `websettings.json`, `.b.json` | green | — (applying without a welcome page fixed in Phase 1) |
| Hooks | `hooks.test.ts` | `hooks.json` | green | — (the 202 poll, its `Location` and failures fixed in Phase 1) |
| Taxonomy, ContentTypes, Files, PropertyBagEntries | `<name>.test.ts` | — | skipped: JSOM, until the JSOM exit | — |
