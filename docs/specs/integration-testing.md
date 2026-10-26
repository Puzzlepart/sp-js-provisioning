# Spec: Integration Testing Against A Site

Status: Ready (Phases 0a, 0b and 1 done 2026-10-21; the first `ci.yml` run on GitHub pending; Phase 2 next)

Owner: Remi Blom-Ohlsen

Related:

- PP365 `releases/1.15`: `integration/` (`src/env.ts`, `src/auth.ts`, `src/sites.ts`),
  `e2e/tests/auth.setup.ts`, `.claude/skills/pp365-testing`, `.tasks/check-skills.js`
- [Site groups and web security](./site-groups-and-web-security.md) (Verified in Phase 1)
- [Deterministic content type ids](./deterministic-content-type-ids.md),
  [ADR](../adr/jsom-explicit-content-type-ids.md), `tests/smoke/deterministic-content-type-ids/`
- `debug/README.md`, [Architecture](../ARCHITECTURE.md)
- Follow-ups: `docs/specs/jsom-exit.md` (D3), a Lists idempotency spec (D4)

## Problem

The engine has 14 handlers and no test that runs against SharePoint. At 4ddfa93:

- 30 offline unit tests in 7 files (`node:test` through `tsx`, fakes per file). No CI, no
  coverage, no `AGENTS.md`, no skills.
- The only live path, `npm run provision` (`debug/provision.ts`), asserts nothing, cleans nothing,
  cannot pass `parameters`, loads only `manifest.provisioning.hubTemplate` and reads credentials
  from the gitignored `debug/provision.settings.ts`. Its `HANDLERS` list omits `Security` and
  `WebSettings`, so both are dropped unless `--handlers` names them. `package-lock.json` resolves
  every `@pnp/*` to 3.17.0 against 4.21.0 in `package.json`, so `npm ci` fails. No successful run
  is recorded since the PnPjs 4 upgrade.
- "No handler threw" proves little. Security, Lists, ClientSidePages and Files log failures and
  carry on. Security logs at Error; Lists logs content type binding and removal, list fields, view
  fields, the root folder, folders and data rows at Info (`lists.ts` 210, 268, 349, 613, 646, 719,
  788); ClientSidePages logs a failed `addControl` at Info and a missing web part at Warning. An
  unresolved lookup value is dropped from the item payload without a log line (`JSON.stringify`
  omits `undefined`): the column stays empty on add and keeps its old value on update; in a
  LookupMulti it is sent as `null`.
- Four handlers cannot run in Node. Taxonomy and ContentTypes use JSOM (`spfx-jsom`, stubbed by
  `debug/stubs/spfx-jsom.ts`). Files throws a string before its `try` whenever
  `config.spfxContext` is set, which every consumer does, so any template with `Files` aborts
  `applyTemplate` with `ProvisioningError('Files')` and an empty message; PP365 1.15's
  `PackageInstaller` strips only `Taxonomy`, so a `pp-testprosjekt` hub import aborts there.
  PropertyBagEntries calls `reject()` in Node and under SPFx, and
  `new ProvisioningError(handler, undefined)` throws a `TypeError`, losing `.handler`.
- Defects only a site shows: WebSettings calls `web.update` only inside `if (WelcomePage)`; Hooks
  never resolves when a 202 poll ends in 200; `onSetup` subscribes a `ConsoleListener` on every
  `applyTemplate`; `applyTemplate` reads `this.config.logging.prefix` unguarded; `{hosturl}` is
  `${protocol}//${host}:${port}`, ending in `:` on default ports. ComposedLook, ClientSidePages,
  Navigation and WebSettings (and Files) reach `window`/`document` through `src/util/index.ts` and
  run in Node only behind the runner's shims.
- Idempotency is claimed and never asserted. Lists is partly destructive: `processField` deletes
  and recreates list fields on every run, `RemoveExistingViews` deletes every view, list
  `breakRoleInheritance` is unconditional.
- eslint (`quotes: single`) and prettier (`singleQuote`) disagree on `"''"` (3 errors on main:
  `lists.ts:1015`, `security.ts:125`, `util/security.ts:57`). `schema.json` lacks `DataRows`,
  `Folders` and `Taxonomy`, leaks DOM types from `IHooks` and allows unknown keys.

PP365 1.15 built this layer for its own library (`integration/`: placeholder-aware configuration,
certificate or cookie sign-in, a reserved site for writes, read-back assertions, strict mode,
JUnit, skills gated by `check-skills`). This spec ports that shape to the engine.

## Goals

- One `.env` and short commands to try a template against a site (`provision`), prove it is there
  (`verify`), see what a site holds (`inspect`), remove leftovers (`cleanup`) and test every
  handler (`test:integration`).
- Every handler has a fixture, a REST read-back, an apply-twice check, an offline test and a
  matrix row; a handler that cannot be tested yet skips and says why.
- Security moves to Verified; the defects the tests expose are fixed (D4).
- Offline CI on every pull request and push to `main`; tenant runs only by an approved
  `workflow_dispatch` from `main` (D5).
- Skills and `AGENTS.md` that let an agent pick the layer, run it and classify a failure.

## Non-Goals

- Porting the JSOM handlers (D3) or changing Lists' destructive behaviours (D4, pinned only).
- A browser suite for the engine; rendering is PP365 `e2e/`'s concern.
- PP365's wizard (PreTask, SitePermissions, SetTaxonomyFields, CopyListData, Planner);
  `--mode project` mirrors only its `ApplyTemplate` and `Hooks` steps.
- Validating `.pppkg` manifests; device-code or Azure CLI sign-in (D1).
- Upgrading TypeScript 4.0.7, `@types/node` 8 or eslint 6; publishing from CI.

## Constraints

- `src/` compiles with TypeScript 4.0.7 to ES5 and uses parameter properties (`webprovisioner.ts`,
  `provisioningerror.ts`, `tokenhelper.ts`) and a `paths`-mapped `spfx-jsom`, so Node's type
  stripping cannot run it. Everything runs through `tsx --tsconfig tests/tsconfig.json`
  or `debug/tsconfig.json`, whose `paths` map `spfx-jsom` to the stub.
- `tsx --test` forwards `--test-reporter`, `--test-reporter-destination` and
  `--experimental-test-coverage` (verified on Node 22.22.2). A JUnit destination in a missing
  folder crashes with `ENOENT`, so scripts create `test-results/` first.
- Harness code parses under TypeScript 4.0 (namespace import of `node-forge`, extensionless
  imports). `node:test` typings are hand-declared in `tests/types/node-builtins.d.ts` and grow
  (`before`, `after`, `describe` options, `t.skip`, `t.todo`, `t.diagnostic`, `assert.throws`).
- One physical `@pnp/sp` (`tests/unit/package.test.ts`). `@azure/msal-node` 7.0.1 needs Node 20+;
  it is a devDependency and `.nvmrc` is 22.22.2. `.gitignore` whitelists `debug/*` entries.
- `_api/v2.1/termstore` needs `TermStore.Read.All` (admin consent) app-only and a `Mozilla/5.0`
  User-Agent with cookies. App-only has no user and is not permission-trimmed. SharePoint Online
  refuses app-only tokens from a client secret.
- Modern sites block custom script (`DenyAddAndCustomizePages`) by default and subwebs inherit
  it; `ScriptLink` custom actions and property-bag writes need it allowed.
- No token, cookie, certificate or tenant URL in a log, artifact or committed file; errors name
  the variable, not the value. Tenant runs are serial, one run per reserved site at a time.

## Decisions

| | Decision | Rationale |
|---|---|---|
| D1 | Local sign-in is cookie mode with PP365's Playwright storage state (default `INTEGRATION_COOKIES_FILE=../prosjektportalen365/e2e/.auth/user.json`, renewed with `npx playwright test --project=setup` in PP365 `e2e/`). CI signs in with the certificate (`CI_TENANT`, `CI_CLIENT_ID`, `CI_CERT_BASE64`, as in PP365). No client secret, no settings file. Device code and Azure CLI are a possible later addition, not planned work. | No app registration or certificate on a laptop; PP365 already renews the session; secret-based app-only is refused. |
| D2 | A reserved site created once by hand (e.g. `/sites/spjs-itest`, subsites enabled). A scratch subweb per suite or run, created with unique permissions and its own associated groups (`web.createDefaultAssociatedGroups(...)`), deleted afterwards unless kept; `cleanup` removes leftovers older than 2 h, so it never deletes the web of a run in progress (`--older-than`, `--all` when no other run is active). Commands that write (tests, `cleanup`, `provision`) refuse a site that is not `INTEGRATION_SITE_URL` or below it unless an explicit override flag is given. NoScript is probed in reachability; the ScriptLink and PropertyBag suites skip naming the setting. | Provisioning cannot be undone; a fresh subweb gives "first" and "second apply" meaning; its own groups make `{associated*groupid}` resolve without a group-connected site; the guard keeps agents off customer sites. |
| D3 | The JSOM handlers (ContentTypes, Taxonomy, PropertyBagEntries, Files web parts) get `docs/specs/jsom-exit.md`: a half-day live spike (REST `contentTypes.add` with an explicit id; one `ProcessQuery` term set), then a port to CSOM over HTTP (`_vti_bin/client.svc/ProcessQuery`) and removal of `spfx-jsom`. Until then their suites skip with the documented reason. | CSOM keeps caller-supplied ids and runs in Node and the browser; the spike first checks whether REST suffices. |
| D4 | Small defects the tests find are fixed with the test that exposes them, one CHANGELOG line each, shipped in 1.5.0 with Security. Lists' field recreation, unconditional list `breakRoleInheritance` and `RemoveExistingViews` on every run are only pinned (`{ todo }`) and get their own spec. | Each fix is small and tested; the Lists behaviours change data on consumer sites. |
| D5 | Tenant CI: `workflow_dispatch` only, from `main`, behind a protected GitHub environment with required reviewers, a concurrency group, `INTEGRATION_STRICT=true`. Offline CI (build, typecheck, unit and contract tests, coverage floors, lint once fixed, `check-skills`, `npm pack` check) on every pull request and push to `main`. No push-triggered tenant runs. | The repository is public and the certificate likely grants tenant-wide `Sites.FullControl.All`. |
| D6 | Scope now: this spec plus Phase 0a. | The runner must work first. Its first live run is the first tenant evidence for Security, recorded in that spec's Verification Result; Security becomes Verified with `security.test.ts` in Phase 1, and 1.5.0 is cut then with D4's fixes. |

## Proposed Approach

### Layers

| Layer | Where | Tenant | Phase |
|---|---|---|---|
| Unit: handlers and runner logic with fakes | `tests/unit/` | no | exists; 0a adds `runner.test.ts` |
| Contract: real PnPjs 4, fake transport, assert requests | `tests/contract/` | no | 2 |
| Runner: try a template against a site | `npm run provision` | yes | 0a |
| Integration: apply, read back, apply again on a scratch subweb | `tests/integration/` | yes | 1 |
| Hosting packages | `tests/integration/packages.test.ts` | yes | 2 |
| JSOM handlers | the smoke README, by hand | yes | until D3 |

Ported from PP365 1.15: `configured()`, `authMode()`, both sign-in pipes, `connect(url)` with a
stand-in SPFx context, `skipReason()` with `INTEGRATION_STRICT`, the reserved-site and read-back
rules, `test:ci`, `check-skills.js`, the skill format.
Not ported: Jest/Heft, `pnpStub.js`, loader hooks, type stripping, a separate package, the
Norwegian guide. Only `CI_*`, `INTEGRATION_AUTH`, `INTEGRATION_COOKIES_FILE`, `INTEGRATION_USER`
and `INTEGRATION_STRICT` mean what they mean in PP365; PP365's `INTEGRATION_LOCALE` (the SPFx
bundle locale) has no counterpart, the template locale is `--lang`.

### Phase 0a: a runner you can trust (done)

No `src/` change:

- `package-lock.json` regenerated so every `@pnp/*` resolves to 4.21.0; `@pnp/nodejs` removed from
  devDependencies; `@azure/msal-node` 7.0.1, `node-forge` 1.4.0, `@types/node-forge` 1.3.14 and
  `dotenv` 17.2.3 added as devDependencies; `npm ci` green.
- `debug/lib/env.ts`: the root `.env` through dotenv; `configured()` treats `<placeholder>` as
  unset; `env = { siteUrl: INTEGRATION_SITE_URL, auth, certificate, cookiesFile }` (`user` and
  `strict` arrive with `skipReason()` in Phase 1, where they are first read);
  `authMode()` = `INTEGRATION_AUTH`, else `certificate` when `CI_*` are set, else `cookies` when the
  file exists.
- `debug/lib/auth.ts`: `signedIn(mode, baseUrl)`, the cookie and certificate behaviours ported from
  PP365 `integration/src/auth.ts` (namespace import of `node-forge`, TypeScript 4.0 syntax).
- `debug/lib/connect.ts`: `connect(url, mode?, configuration?)` returns
  `{ url, sp, web, siteInfo, webInfo, context, mode }`, the SPFx context stand-in built from
  `_api/site` and `_api/web`.
- `debug/lib/cli.ts`: `parseArgs`, `planHandlers`, `cloudHubSubset`, `resolveTemplate`,
  `transformSchema`; pure, unit-tested.
- `debug/lib/logs.ts`: a `FunctionListener` collecting entries by `(Handler):` prefix and level;
  `startRun()` clears the subscribers and subscribes a fresh collector, on the engine's `Logger`
  (a dynamic `import('@pnp/logging')` loads the ES module build, a second `Logger`).
- `debug/provision.ts` on the harness: `--site` defaults to `INTEGRATION_SITE_URL`; handlers come
  from `DefaultHandlerSort` (Security and WebSettings now run); `--handlers` validated (exit 2 with
  the valid list); `--parameters k=v,...` into `IProvisioningConfig.parameters`; `--strict` exits 1
  on a `ProvisioningError` or any Error-level entry, documented as weaker than a read-back until
  D4's log-level fixes land; `--twice` applies twice and reports what the second run logged;
  `--mode cloud` mirrors `PackageInstaller.provisionCloudTemplateHubDependencies` (Lists kept when
  `(bindings && !DataRows) || DataRows`, minus `manifest.provisioning.listContent[].sourceList`;
  Taxonomy unless `--skip-taxonomy`). `debug/provision.settings(.example).ts` removed. Browser shims
  kept for now; the `src/util` guards come with D4.
- `.env.example` at the root; `.gitignore` adds `.env` and `test-results/` and whitelists
  `debug/lib`.
- `tests/fixtures/templates/smoke.json`: a small Node-capable, prefixed template (a site field, a
  site group with an owner and a list granted to it, list `Security`) for the first live run:
  `npm run provision -- --package tests/fixtures/templates/smoke.json --twice --strict`.
- `tests/unit/runner.test.ts` (offline): `parseArgs`; `planHandlers` including Security, WebSettings
  and `NODE_UNSUPPORTED`; `--handlers` validation; `--parameters`; `cloudHubSubset` versus the
  PackageInstaller rules; `configured()` placeholders; cookie session parsing and expiry (with a
  temp file).
- `tests/unit/connect.test.ts` (offline, fake `fetch`): `connect()` sends the session cookie to
  the target web's `_api` and builds `pageContext.site` from the site collection and
  `pageContext.web` from the web.
- `debug/README.md` rewritten around `.env` and cookie sign-in; CHANGELOG `[Unreleased]`
  "Development" lines.

Exit: `npm ci`, `npm run build`, `npm test` and `npm run typecheck:tests` green; a dry run of
`smoke.json` lists Security, with no site or `.env` needed; the user performs the
first live run against the reserved site and records it in the Security spec's Verification
Result.

### Commands (target)

| Command | Phase | Does |
|---|---|---|
| `npm test`, `npm run typecheck:tests` | now | Offline unit tests (contract tests join in 2); type check. |
| `npm run provision -- --package <path> [--site <url>] [--lang <locale>] [--mode hub\|cloud] [--handlers <csv>] [--parameters k=v,...] [--skip-taxonomy] [--strict] [--twice] [--dry-run]` | 0a | The runner. Phase 1 added `--allow-site <url>`; Phase 2 adds `--verify` and `--mode project`. |
| `npm run test:ci`, `npm run test:coverage`, `npm run lint:check` | 0b | JUnit to `test-results/unit.xml`; coverage floors; lint without rewriting. |
| `npm run check:tenant` | 1 | Reachability only, strict: mode, site, scratch web, associated groups, term store, custom script, each missing permission named. |
| `npm run test:integration -- [<Handler ...> \| all \| packages] [--keep]` | 1 | Reachability in its own process (the suites do not run when it fails), then serial suites. No argument: every handler suite; `<Handler ...>`: those suites; `all` (2): the handler suites plus `everything.test.ts`; `packages` (2): `packages.test.ts`. `--keep` keeps scratch webs. |
| `npm run test:integration:ci` | 1 | Strict, spec + JUnit to `test-results/reachability.xml` and `test-results/integration.xml`. |
| `npm run verify -- --package <path> [--site <url>] [--exact] [--json]` | 2 | Read-only template-versus-site diff; exit 1 on a finding. |
| `npm run inspect -- [--site <url>] [--prefix <text>] [--out <file>]` | 2 | Read-only JSON dump of everything the handlers write. |
| `npm run cleanup -- [--site <url>] [--webs] [--prefix <text>] [--older-than <age> \| --all] [--apply]` | 1 | Dry run unless `--apply`. |
| `npm run check-skills` | 1 | Validates `.claude/skills`. |

```json
"test:ci": "node -e \"require('node:fs').mkdirSync('test-results',{recursive:true})\" && tsx --tsconfig tests/tsconfig.json --test --test-reporter=spec --test-reporter-destination=stdout --test-reporter=junit --test-reporter-destination=test-results/unit.xml \"tests/unit/*.test.ts\"",
"test:coverage": "tsx --enable-source-maps --tsconfig tests/tsconfig.json --test --experimental-test-coverage --test-coverage-include=\"src/**\" --test-coverage-lines=66 --test-coverage-branches=74 --test-coverage-functions=65 \"tests/unit/*.test.ts\"",
"lint:check": "eslint --ext .ts ./src && prettier \"**/*.ts*\" --check",
"check:tenant": "tsx --tsconfig tests/tsconfig.json tests/integration/run.ts reachability --strict",
"test:integration": "tsx --tsconfig tests/tsconfig.json tests/integration/run.ts",
"test:integration:ci": "tsx --tsconfig tests/tsconfig.json tests/integration/run.ts --ci",
"cleanup": "tsx --tsconfig debug/tsconfig.json debug/cleanup.ts",
"check-skills": "node ./.tasks/check-skills.js"
```

From Phase 2, `test`, `test:ci` and `test:coverage` run
`"tests/unit/*.test.ts" "tests/contract/*.test.ts"`, and `verify` and `inspect` join as
`tsx --tsconfig debug/tsconfig.json debug/<name>.ts`.

Later phases add to `.env.example`: `INTEGRATION_KEEP`, `INTEGRATION_SCRATCH=suite|run` (1),
`INTEGRATION_HOSTING_DIR`, `INTEGRATION_PACKAGES=testprosjekt|all` (2).

### Sign-in (D1)

- Cookies (local default): `FedAuth`/`rtFa` from the storage state, expiry checked (the error
  names the renew command), `RequestDigest()`, a `Mozilla/5.0` User-Agent. The account is PP365's
  MFA-excluded e2e user (`E2E_USERNAME`/`E2E_PASSWORD` in PP365 `e2e/.env`; about two days per
  session) and must be an owner (Full Control) of the reserved site for 0a, and its site
  collection administrator from Phase 1 (cleanup deletes site groups, including one that owns
  itself, and may purge the recycle bin).
- Certificate (CI): `ConfidentialClientApplication`, the PFX opened in memory by `node-forge`
  (empty and no password, SHA-256 thumbprint), `acquireTokenByClientCredential` for
  `<origin>/.default`. The app needs `Sites.FullControl.All` and `TermStore.Read.All` with admin
  consent; `INTEGRATION_USER` supplies the user for members, owners and User values.

### Target site, scratch webs and the site guard (D2)

- The reserved site is a modern team site with subsite creation enabled. Its URL lives in `.env`
  and a GitHub environment variable, never in the repository. It need not be group-connected (the
  one in use is, `GROUP#0`). Subsite creation was disabled in the tenant ("Disable subsite
  creation for all sites" also refuses `webs.add` through the API: "Nye sekundære områder er ikke
  tilgjengelige for organisasjonen"); it was enabled on 2026-10-21 for D2. Custom script is
  blocked on the reserved site (NoScript); `check:tenant` reports it.
- Scratch web (`debug/lib/scratch.ts`, Phase 1): `webs.add(title, url, '', 'STS#3', lcid, false)`
  with title `SPJS Itest <suite> <stamp>`, url `spjs-itest-<suite>-<stamp>` and the reserved web's
  language, so it has unique permissions; then `createDefaultAssociatedGroups(title, owner)`.
  PnPjs 4.21 calls `breakRoleInheritance(false, true)` first, then the REST method. A web created
  with unique permissions has no associated groups until then. Under NoScript the call answers 403
  although it creates and associates the groups (observed 2026-10-21), so a 403 is accepted when all
  three groups are there afterwards. The owner is the signed-in user (cookies) or
  `INTEGRATION_USER` (certificate). `<stamp>` is UTC `yyMMddHHmmss`, so age is readable from names.
  One web per suite file by default; `INTEGRATION_SCRATCH=run` shares one per invocation. The web
  and the site groups whose title starts with the web's title (its associated groups, and fixture
  groups named through `{sitename}`) are deleted in `after` unless kept; site groups outlive the
  web. Groups are matched by the whole title, not the stamp: two suites started in the same second
  share a stamp, and a stamp-only match deleted another suite's groups during Phase 1.
- Site guard (`assertWritable(url)` in `connect.ts`, Phase 1): `provision`, `cleanup` and every
  suite refuse a target that is not `INTEGRATION_SITE_URL` or a web below it (case-insensitive
  path prefix). `--allow-site <url>` (equal to `--site`) is the override for `provision` and
  `cleanup`. Suites and `check:tenant` take their target only from `INTEGRATION_SITE_URL`
  (scratch webs below it), so they cannot address another site and need no override. Writes
  start after printing target, sign-in mode and handlers. Until Phase 1, `provision` writes
  wherever `--site` points (default: the reserved site).
- Custom script: `customScriptAllowed(web)` checks `AddAndCustomizePages` in the web's effective
  permissions and confirms with a throwaway `ScriptLink` custom action it deletes again.
  Reachability prints it; `customactions-script` and `propertybag` skip with "custom script is
  blocked (DenyAddAndCustomizePages); an SPO admin must allow it".
- Cleanup (dry run unless `--apply`): `--webs` removes `spjs-itest-*` subwebs older than
  `--older-than` (default 2 h; `--all` ignores age) with their groups, and stamped `SPJS Itest ...`
  groups older than that, whether or not their web is still there; `--prefix` removes prefixed artefacts from one web in dependency order
  (lists, content types, site fields, custom actions, navigation nodes, pages, site groups). The
  age threshold keeps a run from deleting another's webs; reachability reports younger scratch
  webs as "another run may be active". The 0a smoke run's artefacts on the reserved root go with
  `npm run cleanup -- --prefix 'SPJS Itest' --apply`.

### Harness (Phase 1)

- `debug/lib/apply.ts`: `apply(conn, schema, { handlers?, parameters?, prefix?, onHandler? })`
  returns `{ order, skipped, entries, errors, warnings, error? }` instead of throwing; `handlers`
  defaults to `planHandlers()`. Quiet without `prefix` (no console listener).
- `debug/lib/redact.ts`: `redact()` and `describeError()`. Failures are reported by message only:
  a `ProvisioningError`'s `cause.response` holds the response headers, `Set-Cookie` included, and a
  failing `assert.equal(result.error, undefined)` printed them (found in Phase 1). Cookie values,
  bearer tokens and non-example SharePoint hosts are replaced.
- `debug/lib/readback/<handler>.ts`: REST readers returning normalised plain objects.
- `debug/lib/verify.ts`: `expectations(schema, conn)` applies the engine's per-field token rules
  (matrix), not a blanket `TokenHelper` pass; `diff()` returns `{ handler, path, expected, actual }`
  findings. Lists compare as a superset unless `RemoveExisting*`; `Skip` rows by key; pages with
  `Overwrite: false` and custom actions by existence; navigation URLs only with `IgnoreExisting`.
  `--exact` compares everything on a fresh scratch web.
- `debug/lib/cleanup.ts`; `debug/{verify,inspect,cleanup}.ts` as thin CLIs. Phase 1 delivered
  `cleanup`; `verify` and `inspect` moved to Phase 2 (the suites' readers are their basis).
- `tests/integration/run.ts` maps handler names to files and spawns
  `tsx --tsconfig tests/tsconfig.json --test --test-concurrency=1`, reachability in a spawn of its
  own first (`node --test` sorts the files it is given, so one spawn ran it in alphabetical place,
  after ComposedLook); `--ci` creates `test-results/`, sets `INTEGRATION_STRICT=true` and adds
  spec and JUnit reporters.
- `tests/integration/helpers.ts`: `skipReason`, `withScratchWeb`, `applyTwice`, `assertClean`;
  `tests/integration/hookserver.ts`: the local listener for Hooks. `tests/fixtures/fixtures.ts`:
  `loadFixture(name, { user, hookUrl })`. A handler name selects its suites
  (`Lists` runs `lists` and `lists-datarows`).
- `.gitignore` whitelists `debug/cleanup.ts` and `debug/lib/` (Phase 1), then `debug/verify.ts`
  and `debug/inspect.ts` (Phase 2).
  `tests/tsconfig.json` `include` adds `integration/**/*.ts`, `fixtures/**/*.ts` and
  `../debug/*.ts` (Phase 1), then `contract/**/*.ts` and `helpers/**/*.ts` (Phase 2).

### Test shape and assertion rules

```ts
describe('Lists', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('lists')
  it('applies, reads back and changes nothing on a second apply', async () => {
    const { conn } = scratch()
    const { afterFirst } = await applyTwice(conn, () => loadFixture('lists'), (c) =>
      readLists(c, [TASKS, DOCS])
    )
    // assertions on afterFirst, with tenant values read in the test
  })
})
```

As built in Phase 1 (`docs/TESTING.md` has the reference suite).

1. Read back through REST and compare with the fixture. Tenant-dependent values (role definition
   ids, group ids, web part component ids) are read in the same test.
2. `applyTwice`: apply and read back, twice. Each apply ends without a `ProvisioningError` and with
   `logs.errors()` empty; the read-backs are equal after stripping volatile members (`Modified`,
   ETags, navigation node ids, page `UniqueId` with `Overwrite: true`); handler checks prove
   "updated, not duplicated".
3. Each apply and expectation gets a fresh copy from `loadFixture()`: handlers mutate their input
   (Navigation sets `node.Url`, WebSettings writes values back into `settings`, ContentTypes sorts
   in place).
4. The read-back is the proof; logs are a second signal (until 0b, `errors()` misses Lists and
   ClientSidePages failures).
5. Fixtures use engine tokens only where the engine replaces them (matrix). `loadFixture`
   substitutes two test-only placeholders in fields the engine never token-replaces: `@@user@@`
   (User values, members in certificate mode) and `@@hookUrl@@` (Hooks `Url`). No token in a
   `KeyColumn` value.
6. A missing site or sign-in skips the `describe` (`skipReason`), and throws under
   `INTEGRATION_STRICT=true`. A refused sign-in fails `before` with what to do (`accessDenied()`
   in `connectToSite()`). A missing permission (a 403 in certificate mode, checked in the test, as
   reachability does for the term store), user or custom script skips the test naming it. No retries: a flaky call becomes `{ todo }` and an
   issue. Known defects are pinned with `{ todo }`.

### Per-handler matrix

Token reach: T = `TokenHelper` (`{listid:}`, `{listviewid:}`, `{webid}`/`{siteid}`, `{sitename}`,
`{sitecollectionid}`, `{sitecollectiontermstoreid}`, `{parameter:}`); U = URL tokens (`{site}`,
`{sitecollection}`, `{wpgallery}`, `{hosturl}`, `{themegallery}`); G = associated-group tokens.
"Node today" is the engine at 4ddfa93 under the 0a runner.

| Handler (sort) | Node today | Token reach | Read-back (REST) | Second apply | Pins and notes |
|---|---|---|---|---|---|
| Taxonomy (-1) | no: JSOM | none | `_api/v2.1/termstore` groups, sets, terms (ids, labels, properties, sort order) | same ids and counts | skips (D3) |
| SiteFields (0) | yes | T in `SchemaXml`; `{listid:}` only for lists existing when the run starts | `web.fields.getByInternalNameOrTitle(n)`: `Id`, `TypeAsString`, `Title`, `Group`, `Required`, `Hidden`, `SchemaXml` | same `Id`, update path; B variant changes `Title`/`Required` | a lookup to a list from the same template is unresolved on the first apply: pre-create the source or use list `Fields` |
| ContentTypes (1) | no: JSOM | none | `web.contentTypes.getById(id)` with field links | no duplicate, links unchanged | skips (D3) |
| Features (2) | yes | none | `web.features.select('DefinitionId')()` membership (inactive is not a 404) | unchanged | feature chosen in Phase 1 |
| Security (2.5) | yes (old `HANDLERS` omitted it) | T in group `Title`, string properties, `Members`, `Owner`, `Principal`; G in `Owner`, `Principal` | `siteGroups.getByName(t)` with owner and join/leave settings, `.users()`; `HasUniqueRoleAssignments`; `roleAssignments.expand('Member','RoleDefinitionBindings')` | no duplicate group or assignment; `ClearExistingMembers` exact | failures at Error; owner `{associatedownergroupid}` and the group itself; the "inherits, break" branch on a nested inheriting web |
| Lists (3) | yes (`initSpfxJsom` on the stub, JSOM unused) | T in `Fields` XML, DataRows text and URL values; T + G in `Security` principals; not titles, views, `FieldRefs`, bindings, folders, User/Lookup values or the `KeyColumn` match | list properties, `contentTypes()`, fields, views with fields, folders, items with `FieldValuesAsText`, `roleAssignments` | row count by key; `Skip` keeps an edited value, `Overwrite` restores it; no duplicate views, folders, bindings | D4 levels (0b); `{ todo }`: field recreation (a `Fields` column value is lost), list `breakRoleInheritance`, `RemoveExistingViews`, tokenised key duplicates rows, silent unresolved lookups; until D3 only OOTB content types valid for the template (not `0x0101` on a 100 list) |
| Files (4) | no: throws under `spfxContext` before `try`; JSOM web parts | T + U in `Src`, `FileSrc`; T in web part XML | file, `listItemAllFields`, web parts | `Overwrite` | skips (D3); D4 `Error` (0b) |
| CustomActions (5) | yes | none | `userCustomActions` by title prefix: `Location`, `Sequence`, component id/properties, `ScriptBlock` | no duplicate by title; changed `Sequence` not applied (add-only) | `ScriptLink` in `customactions-script.json`, skipped under NoScript |
| ComposedLook (6) | yes, shims | U | best effort: `_api/web/ThemeInfo` changed | no error | util guards (D4, 1) |
| ClientSidePages (7) | yes, shims | T + U in `Properties`; T in `ServerProcessedContent`; not `Name`, `Title`, control `Id` | `web.loadClientsidePage(path)`: title, layout, comments, sections, controls | `Overwrite: false` skipped; `true` recreated | literal OOTB component GUIDs, asserted in `getClientsideWebParts()` first; D4 level (0b); no Warning from the handler asserted |
| PropertyBagEntries (8) | no: `reject()` with `undefined` | none | `web.allProperties()`, `vti_indexedpropertykeys` | same | skips (D3, NoScript); D4 `Error` (0b); writes only `Overwrite: true` entries (pinned) |
| Navigation (9) | yes, shims | U in `Url` only | `quicklaunch()`, `topNavigationBar()`, children: titles, URLs, order | same count and order (new ids) | deletes all nodes first (scratch defaults vanish: asserted); an existing title keeps its URL unless `IgnoreExisting` (B) |
| WebSettings (10) | yes, shims (old `HANDLERS` omitted it) | U in string values | `web.select(...)`, `rootFolder.select('WelcomePage')` | same | D4: apply without `WelcomePage` (1); master page and CSS URLs out of the fixture (NoScript) |
| Hooks (11) | yes (global `fetch`) | none | a local `node:http` listener (reachable: hooks run in the calling process) records method, headers, body with merged `allProperties()` | called again by design | D4: resolve after 202 then 2xx, a relative `Location`, a 202 without one, a failed poll (1, fixed) |
| WebProvisioner | | | order from `progressCallback` = `DefaultHandlerSort`; `handlers` filter; `Parameters`/`Version` ignored; tokens equal site values | | D4 (0b) |

### Engine defects (D4)

| Defect | Fix | Exposing test | Phase |
|---|---|---|---|
| Lists and ClientSidePages failures at Info | `log_error` for binding, list field add, view fields, root folder, folder, row, `addControl`; `log_warn` for a content type removal refused as in use; "field does not exist" stays Info | unit tests with throwing fakes asserting the level | 0b, done |
| `ProvisioningError` reads `error.message` from any cause | readable message for non-`Error` and `undefined` causes, `.handler` kept, prototype restored (the ES5 build returns a plain `Error`, so `instanceof ProvisioningError` is false in `lib/`) | orchestration unit tests | 0b, done |
| Files throws a string; PropertyBagEntries rejects with `undefined` and logs unprefixed | an `Error` naming the reason, logged through `log_error` | unit tests; `--handlers Files` | 0b, done |
| `ConsoleListener` subscribed on every `applyTemplate` | once per process; the runner's `startRun()` then subscribes its collector once and clears it per run, since its `Logger.clearSubscribers()` would remove that listener before run 2 | `webprovisioner.test.ts`: `applyTemplate` twice (same and a new provisioner, `logging` set) leaves one `ConsoleListener`; then `provision --twice` printing each line once is also evidence | 0b, done |
| `this.config.logging.prefix` read unguarded | optional chaining, as `logWarning` does | `setup({})` unit test | 0b, done |
| Security: a group as `Owner` is not applied (found by the 0a tenant run, 2026-10-21): `SetUserAsOwner` answers success for a group id and changes nothing, and the handler logs success | CSOM `Group.Owner` over `_vti_bin/client.svc/ProcessQuery` when the principal is a group (`SetUserAsOwner` kept for users); read the owner back and `log_error` on a mismatch | unit test with a fake web that ignores the call (asserts the logged error); `provision --twice --strict` on `smoke.json` plus a read-back of `Owner` | 0b, done |
| WebSettings updates the web only with `WelcomePage` | update whenever other keys are set; no empty update | `tests/unit/websettings.test.ts` (fake web: `update` without `WelcomePage`, no empty update); `tests/integration/websettings.test.ts` | 1, done |
| Hooks hangs after 202 then 200 | resolve on a non-202 2xx | fake-`fetch` unit test, `hooks.test.ts` | 1, done |
| `src/util` reads `window`/`document` unconditionally | origin from `spfxContext.pageContext.web.absoluteUrl`, `_spPageContextInfo` only with `window`, `new URL` in `makeUrlRelative`, `Buffer` when `btoa` is missing; runner shims removed | `tests/unit/util.test.ts` without `window`, and with a fake `window` for browser precedence | 1, done |
| `{hosturl}` ends in `:` | the guard yields the origin; CHANGELOG calls it out under Changed as consumer-visible | `tests/unit/util.test.ts` pins the string | 1, done |
| Each handler and `applyTemplate` ran on the caller's web instance; PnPjs 4's `select()` mutates it, so a `$select` left by a handler (ClientSidePages and Files read `ServerRelativeUrl`, web `Security` `HasUniqueRoleAssignments`) made the next run load only that property: `{sitename}`, `{webid}`, `{siteid}` unresolved (found by the Security suite: groups named "{sitename} Readers") | `onSetup` reads `Web(this.web)()`; each handler gets `Web(this.web)` | `webprovisioner.test.ts`: a stale `$select` is not sent; no `$select` left after a handler read the web | 1, done |
| `{associated...groupid}` on a web without that group (one that inherits) sent `principalid=undefined` | an error naming the missing group | `security.test.ts` (unit) | 1, done |
| Features is not idempotent: `add` on an active feature answers 500, `remove` on an inactive one 400, so a second apply stopped the template | read the active features once; skip what is already in the requested state (Info); `force` still re-activates | `tests/unit/features.test.ts`; `features.test.ts` (second apply) | 1, done |
| ComposedLook without `FontSchemeUrl` threw a `TypeError` in `replaceUrlTokens` | send `null`, as for the background image; both optional in `IComposedLook` | `tests/unit/composedlook.test.ts`; `composedlook.test.ts` | 1, done |
| WebSettings with `WelcomePage` failed under Node: `import * as omit from 'object.omit'` is not callable as a namespace | object rest instead of `object.omit` | `tests/unit/websettings.test.ts` | 1, done |
| Hooks: a failed poll request escaped as an unhandled rejection; a 202 without `Location` polled `null` (in a browser a GET to `<page folder>/null` with the hook's headers); a relative `Location` failed in Node and went to the SharePoint page in a browser | the poll rejects; no `Location` rejects naming the hook; the `Location` is resolved against the hook's URL | `tests/unit/hooks.test.ts` | 1, done |
| Hooks: a request that could not be sent (network, CORS) failed as a bare "Failed to fetch", naming no hook; the poll `Location` was resolved against the hook's URL, which a relative hook URL (valid in a browser) made throw (found by the closing Phase 1 reviews) | the error names the hook, as an error status does; the `Location` is resolved against the URL that answered | `tests/unit/hooks.test.ts` | 1, done |
| ClientSidePages: tokens were replaced in the JSON text of web part properties, so a `{sitename}` or `{parameter:}` value with a quote or a backslash made `JSON.parse` throw and the page lost the part (pre-existing; found by the final Phase 1 review) | replace in each string (`mapStrings`) | `tests/unit/clientsidepages.test.ts`, `util.test.ts` | 1, done |

Phase 2's contract tests pin the same behaviour on the wire.

Pinned only (`{ todo }` with the evidence), each for its own change:

- Lists, for the Lists spec: field delete-and-recreate (a value in a list `Fields` column is lost
  on every re-apply; the hosting packages' Fasesjekkliste lists use `Fields` with `Skip`),
  unconditional list `breakRoleInheritance` (SharePoint ignores it today), `RemoveExistingViews`
  (it deletes every view on every run, the template's own included, and adds them again with new
  ids), the raw `KeyColumn` match (a token in a key value adds the row on every run), silent
  unresolved lookups (a single Lookup is dropped; a LookupMulti entry fails the whole row with 400
  and a message that does not name it), and a URL value's default `Description` keeping the raw
  token.
- SiteFields: `{listid:}` for a list from the same template stays literal on the first run, and
  SharePoint accepts it without an error or warning (SiteFields runs before Lists).
- Navigation: the handler deletes every node before adding the template's; a node SharePoint
  refuses stops the run and leaves the menu with only the nodes added before it.
- Not defects, recorded: classic custom actions (ScriptLink, EditControlBlock, StandardMenu) are
  refused under NoScript, only `ClientSideExtension.*` actions get through; `{site}` on a subweb is
  the site collection's URL (documented); PnPjs's `loadClientsidePage().commentsDisabled` reads
  `false` for a page whose comments are disabled (the reader uses the file's item).

### Fixtures and hosting packages

- `tests/fixtures/templates/`: tenant-free JSON, titles `SPJS Itest ...`, internal names
  `SPJSItest...`, `Group: "SPJS Itest"`, fixed GUIDs listed in `tests/fixtures/README.md`; one per
  matrix row (B variants as `<name>.b.json`) plus `lists-datarows.json`,
  `customactions-script.json` and, in Phase 2, `parameters.json` and `all.node.json`. Members no
  hosting package uses get a line each: `AdditionalProperties`, `ServerProcessedContent`, an
  `EditControlBlock` action with `RegistrationType`, join/leave group settings,
  `PersonalView`/`Hidden` views, `{wpgallery}`, `{themegallery}`, `IgnoreExisting`, `UserMulti`,
  `{lookupId}`. They also mirror three members the hosting packages do use: `{siteid}` in list
  field XML (`WebId`, `SourceID`; anlegg, bygg, vei), a view `Scope` and a Number `KeyColumn`
  (`GtSortOrder`, 6 row sets).
- `all.node.json` composes the handler fixtures in sort order without the JSOM-dependent links
  (fixture content types, term sets, taxonomy rows, Files, PropertyBagEntries);
  `everything.test.ts` applies it twice and runs `verify()`. The full composition is D3's exit.
- `prosjektportalen-hosting`: 12 packages, 26 provisioning JSON files, 8 of them hub templates in
  6 packages. ContentTypes is in all 8 hub templates (14 of 26 files), Taxonomy in 7 (7 of 26);
  Lists 16, SiteFields 11, Navigation 5, ClientSidePages 5; Files and PropertyBagEntries 4 each,
  all in `pp-testprosjekt`; WebSettings once (`pp-enkel-prosjektmal/.../EnkeltProsjekt.json`,
  `WelcomePage` only); `Parameters` once (`pp-enkel-prosjektmal` `template.json`). Navigation
  `Children` is in `pp-forskningsprosjekt/.../Forskningsmal.json`, list `Security` in 6 lists of the
  anlegg, bygg and vei hub templates. Web `Security`, CustomActions, ComposedLook, Features, Hooks
  and `TopNavigationBar` appear nowhere.
- `packages.test.ts` (Phase 2, when `INTEGRATION_HOSTING_DIR` exists): `pp-testprosjekt` in
  `--mode hub`, `cloud` and `project`, nb-NO and en-US, then `verify()`; `EnkeltProsjekt.json`,
  `Forskningsmal.json` and the vei list `Security`. Until D3 it asserts an explicit
  expected-failure list (bindings to content types not created in Node, taxonomy rows).
  `INTEGRATION_PACKAGES=all` adds the rest.
- `schema.json` is regenerated from `src/schema.ts` in Phase 2 (`DataRows`, `Folders`, `Taxonomy`
  and definitions; `IHooks` without DOM types; `additionalProperties: false` where unknown keys are
  ignored). Only then do `tests/unit/schema.test.ts` (fixtures and the 26 hosting files with `ajv`)
  and `fixture-coverage.test.ts` (every property in a fixture or in `NOT_COVERED` with a reason)
  count as evidence.

### Offline layers and CI (D5)

- 0b: eslint `quotes: ['error', 'single', { avoidEscape: true }]` and a `.prettierignore` (`lib/`,
  `coverage/`, `test-results/`) end the lint conflict. Coverage is measured with
  `--enable-source-maps`: without it Node maps V8 offsets in tsx's compiled JS onto the `.ts` text,
  so the figures were wrong per file and changed with the length of the checkout path (the earlier
  68.02 / 77.65 / 77.48 and 73.94 / 76.91 / 79.89 were such figures). The floors start at the
  source-mapped measurement when they were first enforced, at the end of 0b (2026-10-21):
  61.03 / 72.18 / 59.18 (lines, branches, functions; `src/**`), set one point below as 60/71/58
  so small differences between machines do not fail the check. They only rise: raised at the end
  of Phase 1 to 66/74/65 (measured 67.09 / 75.05 / 65.91).
- Phase 2: `tests/contract/` turns the recording web of `webprovisioner.test.ts` into a fake
  transport (`on.send.replace`, canned `Response` bodies, recorded `{ url, method, body }`, an
  SPFx-style relative mode) pinning wire shapes (`createfieldasxml`, `$top=5000`, OData escaping,
  `addroleassignment`, `setuserasowner`, the custom-action batch, navigation delete-then-add, the
  page save). `tests/helpers/` consolidates the fakes; orchestration tests cover sort order, the
  `handlers` filter, a throwing handler and a custom `handlerSort` without `Security` (runs last).
- `ci.yml` (0b): `pull_request` and `push` to `main`; `permissions: contents: read`;
  `concurrency` per ref, cancel in progress; `timeout-minutes: 15`; one job named
  `Build and offline tests` (the required check; no path filter: a workflow skipped by a path
  filter leaves a required check Pending and blocks the merge): `actions/checkout@v7`,
  `actions/setup-node@v7` (`.nvmrc`, npm cache), `npm ci`, `lint:check`, `build`,
  `typecheck:tests`, `test:ci`, `test:coverage`, `npm pack --dry-run --json` asserted to list `lib/index.js` and `lib/handlers/*.js`, `check-skills` (from Phase 1), a grep
  failing on a non-`contoso` `*.sharepoint.com` host, `FedAuth=`, `Bearer ey` or `MII` in
  fixtures, `debug/` and `.env.example`, and `actions/upload-artifact@v7` of `test-results/`
  (`if: always()`, 7 days).
- `integration.yml` (Phase 3): `workflow_dispatch` only (inputs `handlers`, `packages`
  `none|testprosjekt|all`, `keep`), a job condition requiring `refs/heads/main` and the
  `Puzzlepart/sp-js-provisioning` repository, `environment: tenant` (required reviewers, branch
  `main`), `concurrency: { group: spjs-integration, cancel-in-progress: false }`, `timeout-minutes: 30`.
  Secrets (`CI_*`, `INTEGRATION_USER`), the `INTEGRATION_SITE_URL` variable and the inputs reach
  the steps through `env:`, never interpolated into `run:`. Steps: `npm ci`, the hosting checkout
  when `packages` is not `none`, `check:tenant`, `test:integration:ci`, upload `test-results/`.

### Skills, AGENTS.md and docs

- `.claude/skills/<name>/SKILL.md` only; `.tasks/check-skills.js` copied verbatim (root
  `path.resolve(__dirname, '..')`; `.github/skills` and `.agents/skills` fail). PP365's format:
  `name` = folder; `description` of quoted triggers, "Encodes ...", "Pair with ...", at most 1024
  code points; body opening with the authoritative doc (`Authoritative description: <doc>`; where
  a doc and the code disagree, the code wins), tables,
  "Typical task", dated "Pitfalls", "Verification checklist", "Do not".
- Phase 1: `spjs-testing` (layers, commands, fixture rules, matrix, failure triage) and
  `spjs-provision` (sign-in, flags, `hub`/`cloud`/`project` versus PackageInstaller, reading logs,
  what `--strict`/`--twice` prove; `--verify` in Phase 2). Phase 2: `spjs-handlers` (adding a
  handler, token reach, throw versus log, required tests) and `spjs-release` (`postversion`
  publishes; PP365 pin).
  Every "Do not" includes: never run `provision` or `cleanup` against a site the user did not
  name; tenant suites only with the user's go-ahead; never print a secret or commit `.env`; never
  lower a floor or delete an assertion to go green.
- `AGENTS.md` (about 100 lines) and `CLAUDE.md` (`@AGENTS.md`); `docs/TESTING.md` is the
  authoritative testing doc (layers, "which test do I write", the matrix, tenant setup);
  `docs/ARCHITECTURE.md` points to it.

### JSOM handlers (D3)

Their suites exist with fixtures and readers and skip with the reason "JSOM handler, see
`docs/specs/jsom-exit.md`". The ADR rejected PnPjs `contentTypes.add` because it "sets both" `Id`
and `ParentContentType`; PnPjs 4.21 posts `Id: { StringValue }` and no `ParentContentType`, so the
spike re-tests it.

### Phases

| Phase | Delivers | Size | Status |
|---|---|---|---|
| 0a. A runner you can trust | As above | S | Done 2026-10-21: first tenant run green twice, recorded in the Security spec; it found the group-owner defect (D4) |
| 0b. Offline CI and log levels | `ci.yml`, lint fix, `test:ci` (mkdir), coverage floors, `npm pack` check; D4's log-level and `ProvisioningError` fixes (Lists/ClientSidePages levels, non-`Error` and `undefined` causes, Files/PropertyBagEntries `Error`s) plus two small logging guards (`ConsoleListener` once, `logging.prefix`); `tests/fixtures/templates/failures.json` (a list whose `Fields` holds a well-formed Lookup to a missing list, a `Files` entry, a `PropertyBagEntries` entry) | S-M | Done 2026-10-21: the CI sequence green locally (77 unit tests; source-mapped coverage 61.03/72.18/59.18, floors 60/71/58); the three `failures.json` runs exit 1 with the handler and a readable message; `smoke.json --twice --strict` green with each line printed once and both group owners read back as set (also the Security owner defect). Pending: the first `ci.yml` run on GitHub and branch protection. |
| 1. Every Node-capable handler against a site | Site guard; scratch webs with associated groups; reachability with the NoScript probe; `apply`, readers, `verify`, `inspect`, `cleanup`; suites and fixtures for the ten Node-capable handlers, skipping suites for the four JSOM ones; D4's Phase 1 fixes; `AGENTS.md`, `docs/TESTING.md`, two skills, `check-skills` (in `ci.yml`); the gulp-era `.vscode/launch.json`, `.vscode/tasks.json` and `settings.example.js` removed; the Security spec's Verification Plan amended for D2 and Security Verified | M-L | Done 2026-10-21 (`verify` and `inspect` moved to Phase 2): `check:tenant` green; `npm run test:integration` 45 tests, 34 pass, 10 `todo`, 1 skipped, 0 fail, in about 9 minutes, nothing left behind (green on the final code and twice in a row before its last transport change; one run in between was throttled and green again on rerun); 15 suites (10 Node handlers in 11 suites, plus the JSOM four skipping) and reachability; the Security spec Verified. D4's four Phase 1 fixes done; the suites also found 10 engine defects the plan did not list, 6 fixed with tests and 4 pinned (D4); the closing reviews found 2 more, fixed with unit tests. The runs also found two harness defects, fixed: `node --test` sorts its files (reachability ran in alphabetical place), and a request that hung for 3 x 300 s left a scratch web behind. |
| 2. Depth, packages, schema | DataRows value matrix (DateTime under `TZ` Europe/Oslo, UTC, Asia/Tokyo); `all.node.json`; contract tests; `--mode project`; `packages.test.ts`; `schema.json`; floors raised; two skills | M | |
| 3. Tenant CI | `integration.yml`, the `tenant` environment | S | |
| 4. JSOM exit | Own spec (D3) | S + L | |

## Acceptance Criteria

- 0a: `npm ci` succeeds on a fresh clone; `npm ls @pnp/sp @pnp/core @pnp/queryable @pnp/logging`
  shows one 4.21.0 copy each and `@pnp/nodejs` is gone; `build`, `test` and `typecheck:tests` pass;
  `npm run provision -- --package tests/fixtures/templates/smoke.json --dry-run` lists Security
  without signing in, without `.env` and without `--site` (a dry run needs no site); the same
  command with `--handlers Bogus` in place of `--dry-run` exits 2 with the valid names, and
  without `--dry-run` and without a site it exits 2 with "Missing --site". The user performs the first live run
  (`npm run provision -- --package tests/fixtures/templates/smoke.json --twice --strict` against
  the reserved site) and records its outcome in the Security spec's Verification Result: exit
  code, errors and warnings per run, and a look at groups, owners and list permissions. A defect
  it exposes is fixed under D4 with its unit test in 0b, not in 0a.
- 0b: `ci.yml` is green on a pull request without secrets, and `main`'s branch protection requires
  `Build and offline tests`; `lint:check` passes; `test:ci` writes `test-results/unit.xml` on a
  fresh clone; `test:coverage` fails below a floor.
  `npm run provision -- --package tests/fixtures/templates/failures.json --handlers Lists --strict`
  exits 1 (today the failed field is logged at Info); `--handlers Files` and
  `--handlers PropertyBagEntries` each exit 1 naming the handler and a readable message; the
  orchestration unit tests assert the same offline.
- 1: `check:tenant` names mode, site, scratch-web capability, associated groups, term store,
  custom script and each missing permission. `npm run test:integration` applies every
  Node-capable fixture twice on scratch webs, reads back, asserts stability and zero logged
  errors, deletes the webs and their groups, and passes twice in a row; without `.env` every
  suite skips naming the variable, and with `INTEGRATION_STRICT=true` it fails.
  `npm run test:integration -- Lists --keep` plus `npm run verify` (Phase 2) on the kept web
  reports zero findings; `npm run cleanup -- --webs --older-than 0m` (no other run active) lists
  the kept web and deletes nothing, and the same with `--apply` deletes it and its groups. `provision`, `cleanup` and the suites refuse a site outside `INTEGRATION_SITE_URL`. The
  Security spec's Verification Plan reads "a scratch subweb with unique permissions and its own
  associated groups (D2)" in place of "fresh group-connected site" (the second-run and
  group-owner scenarios stay; the 0a smoke run is its first tenant evidence), and the spec is
  Verified by `security.test.ts` and the Lists security fixture, including a group owning a
  group. `check-skills` passes.
- 2: `npm run test:integration -- all` is green twice; `schema.test.ts` validates the 26 hosting
  files; floors are raised.
- 3: an approved dispatch on `main` is green with `integration.xml` uploaded and no suite skipped
  for missing configuration.
- No committed file, log or artifact contains a tenant URL, app id, token, cookie or certificate.

## Verification Plan

- 0a: fresh clone, `nvm use && npm ci && npm run build && npm test && npm run typecheck:tests`;
  `npm ls`; the dry run above on a fresh clone with no `.env`; then the user's
  live run, recorded as in the Acceptance Criteria.
- 0b: a pull request shows the check green and `test-results/` uploaded; the new unit tests (log
  levels, `ProvisioningError` causes, `ConsoleListener` once); the three `failures.json` runs;
  `--twice` on `smoke.json` prints each line once.
- 1: `npm run check:tenant` in cookie mode (certificate mode once the app is granted);
  `npm run test:integration` twice; `npm run test:integration -- Lists --keep`, then
  `npm run inspect` and `npm run verify -- --package tests/fixtures/templates/lists.json` on the
  kept web (Phase 2); `npm run cleanup -- --webs --older-than 0m`, then the same with `--apply`
  (webs younger than 2 h are otherwise left for a run that may be active);
  `INTEGRATION_STRICT=true npm run test:integration` without `INTEGRATION_SITE_URL` exits 1;
  `npm run provision -- --site <other site> --package tests/fixtures/templates/smoke.json` is
  refused; an agent given "test the Lists handler against my site" reaches
  `npm run test:integration -- Lists` from the skill.
- 2: `npm run test:integration -- all`;
  `npm run provision -- --package tests/fixtures/templates/all.node.json --twice --strict --verify`;
  `packages.test.ts` on `pp-testprosjekt`.
- 3: `gh workflow run integration.yml`, approval, `gh run download <id>`.
- Each phase updates this spec's status, CHANGELOG `[Unreleased]` and the specs it verifies.

## Open Questions

- Does `createDefaultAssociatedGroups` work app-only, and with which owner (`INTEGRATION_USER` or
  empty)? Do the associated-group tokens then resolve on the subweb? Reachability, Phase 1.
- Which custom-script signal is reliable in both modes (effective permissions or the throwaway
  `ScriptLink`), and does the tenant re-apply NoScript after an admin allows custom script?
- Which harmless web-scoped feature activates on an STS#3 subweb (the MDS feature from
  `sample-schemas` is not assumed to).
- Does a fresh STS#3 subweb have modern `SitePages`, and does `_api/web/ThemeInfo` reflect
  `applyTheme`? If not, ComposedLook asserts zero errors only.
- What a scratch web costs per suite, and whether `INTEGRATION_SCRATCH=run` should be the default.
- Whether PP365's CI app (the `CI_*` secret D1 reuses) holds `Sites.FullControl.All` and
  `TermStore.Read.All` with admin consent; reachability names what is missing. A dedicated
  `Sites.Selected` app is a possible later hardening, not planned work.
- Whether deleted scratch webs must be purged from the recycle bin on frequent runs.

## Rollout Notes

- 0a changes no `src/`, `lib/` or `schema.json`; of the packed files only `package.json`
  (devDependencies), `README.md` and `CHANGELOG.md` change, with no runtime effect. `@pnp/nodejs`
  leaves devDependencies;
  `debug/provision.settings.ts` is gone, so runner users move to `.env` (`debug/README.md`).
- 1.5.0 ships Security with D4's fixes and is cut when Phase 1 exits with Security Verified.
  Consumer-visible: Lists and ClientSidePages failures appear at Error in install logs; `{hosturl}`
  loses its trailing colon; WebSettings applies without `WelcomePage`; Files and
  PropertyBagEntries errors become readable.
- Then PP365's exact pin (1.4.0 in `ProjectExtensions` and `PortfolioExtensions`) is bumped with
  `npm run rush:update` and one hand project setup. A `pp-testprosjekt` hub import in PP365 still
  stops at Files, now with a readable message, until PackageInstaller strips Files and
  PropertyBagEntries or D3 lands.
- Prerequisites: before 0a can exit, the reserved site (created by hand, subsites enabled, e.g.
  `/sites/spjs-itest`), `INTEGRATION_SITE_URL` in `.env`, and PP365's e2e user as an owner of it
  with a fresh session (`npx playwright test --project=setup` in PP365 `e2e/`); before Phase 1
  ends, that user as site collection administrator; before Phase 3, the certificate app's
  permissions, `INTEGRATION_USER`, and the `tenant` environment with reviewers, secrets and
  variable.
- An interrupted run leaves `spjs-itest-*` webs and `SPJS Itest` groups;
  `npm run cleanup -- --webs --apply` removes those older than 2 h; add `--all` when no other run
  is active.
- `tests/smoke/deterministic-content-type-ids/` stays the manual check until D3; that spec is
  already Verified and gains the automated suite in its Verification Result then.
