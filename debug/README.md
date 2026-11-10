# Local provisioning runner

Apply a template or a Prosjektportalen package to a real SharePoint site straight from Node, with
no SPFx build/deploy loop. The runner reproduces what PP365's `PackageInstaller` does in the
browser (`WebProvisioner(web).setup({ spfxContext, logging }).applyTemplate(schema)`), signs in
with the configuration in `.env`, and prints every log line, including the SharePoint error
payloads. It runs the engine **source** (`src/`) through `tsx`, so a handler change takes effect
on the next run without a build.

The plan this runner is part of is `docs/specs/integration-testing.md`.

## Setup (once)

1. **Node** from `.nvmrc` (`nvm use`), then `npm ci`.
2. **A site reserved for these runs.** Create a site that nothing else writes to, for example
   `/sites/spjs-itest`, and make the account the runner signs in as (step 4) a site collection
   administrator of it (the integration tests and `cleanup` need it; `provision` alone needs an
   owner). A template run cannot be undone, so `provision` and `cleanup` write only to that site
   and webs below it (the site guard): any other site needs `--allow-site` repeating it.
3. **Settings:**
   ```sh
   cp .env.example .env
   ```
   Set `INTEGRATION_SITE_URL` to the reserved site. `.env` is gitignored.
4. **Sign-in.** Locally the runner signs in as PP365's e2e user (`E2E_USERNAME` in
   `prosjektportalen365/e2e/.env`, an MFA-excluded test account), with the SharePoint session
   cookies its Playwright sign-in left (`prosjektportalen365/e2e/.auth/user.json`, the default
   `INTEGRATION_COOKIES_FILE`). Everything the runner writes is attributed to that account.
   Create or renew the session, about every two days, in a PP365 checkout:
   ```sh
   cd ../prosjektportalen365/e2e && npx playwright test --project=setup
   ```
   In CI it signs in as an app with a certificate (`CI_TENANT`, `CI_CLIENT_ID`, `CI_CERT_BASE64`:
   the PFX, base64, without a password, read in memory). SharePoint Online refuses app-only
   tokens from a client secret, so there is no secret option.

## Run

```sh
# The smoke template: a site column, two site groups (one owns itself), a list granted to them.
# First run against a new site: apply twice, fail on any logged error.
npm run provision -- --package tests/fixtures/templates/smoke.json --twice --strict

# A package from prosjektportalen-hosting (folder, .pppkg or .zip) or a raw template JSON
npm run provision -- --package ../prosjektportalen-hosting/packages/pp-testprosjekt
npm run provision -- --package ../prosjektportalen-hosting/dist/pp-testprosjekt-1.0.0.pppkg

# A web below INTEGRATION_SITE_URL, only some handlers, values for {parameter:...} tokens
npm run provision -- --site <reserved-site>/<web> --package <path> --handlers SiteFields,Security,Lists

# Another site than INTEGRATION_SITE_URL: --allow-site must repeat it
npm run provision -- --site <url> --allow-site <url> --package <path>
npm run provision -- --package <path> --parameters stamp=1,owner=kari@contoso.com

# The cloud-template subset, the English hub template, no taxonomy, or only the plan
npm run provision -- --package <path> --mode cloud
npm run provision -- --package <path> --lang en-US --skip-taxonomy
npm run provision -- --package <path> --dry-run

# The project template and its extensions, as the project setup applies them, on a web below
# INTEGRATION_SITE_URL; then compare the result with the templates
npm run provision -- --site <reserved-site>/<web> --package <path> --mode project --verify
npm run provision -- --site <reserved-site>/<web> --package <path> --mode project --extensions all

npm run provision -- --help
```

| Flag | Does |
| --- | --- |
| `--site <url>` | The target site. Default: `INTEGRATION_SITE_URL`. Must be that site or a web below it, unless `--allow-site` names it. |
| `--allow-site <url>` | Write to a site outside `INTEGRATION_SITE_URL`; the same URL as `--site`, to confirm it (exit 2 otherwise). |
| `--package <path>` | A package folder (with `manifest.json`), a `.pppkg`/`.zip` (extracted with `unzip`), or a raw template JSON. A package resolves to its hub template, or with `--mode project` to its project template. |
| `--lang <locale>` | The package's localized template and list content, e.g. `en-US`. |
| `--mode hub\|cloud\|project` | `hub`: the whole hub template (default). `cloud`: the subset PP365 applies when a package is published as a cloud template. `project`: the project template and the selected extensions, as the project setup's `ApplyTemplate` and `Hooks` steps apply them (below). |
| `--extensions <ids\|all\|none>` | `--mode project`: the extensions to apply, by id. Default: the ones the setup selects to begin with (`defaultSelected` or `locked`). |
| `--handlers <csv>` | Only these handlers, checked against the engine (exit 2 with the valid names). Also forces handlers that are skipped in Node. |
| `--parameters <k=v,...>` | Values for `{parameter:k}` tokens. |
| `--skip-taxonomy` | Drop the Taxonomy section. |
| `--strict` | Exit 1 when a handler logged an error and carried on, not only when one threw. |
| `--twice` | Apply twice; the second run shows what re-applying does. |
| `--verify` | Then compare the template with the site, as `npm run verify` does (below), for the sections that ran; exit 1 on a finding. |
| `--exact` | With `--verify`: also compare what a re-apply leaves as it is. For a web the template was just applied to. |
| `--dry-run` | Print the plan and exit, without signing in (no site needed). |

Exit codes: 0 applied, 1 failed (or, with `--strict`, an error was logged; with `--verify`, a
finding), 2 usage error (also a target the site guard refuses).

`--strict` sees what the handlers log as errors, including the failures they log and carry on
after (a site column SharePoint refuses, an unresolved token in a site column, a list field, a
FieldRef to a column that is not found, a view's fields, a folder, a row or an unresolved
lookup or user value in it, a page control). Since 1.5.0 a failing site column is one of these
and no longer stops the run, so only `--strict` makes it fail. It does not read the result
back from the site, so a green strict run is not proof that everything was applied: the first
tenant run found a group owner that SharePoint silently did not change. `--verify` reads it back,
and so do the integration suites (`npm run test:integration`, `docs/TESTING.md`).

## Verify and inspect

```sh
# What differs between a template (or a package) and a web; exit 1 on a finding
npm run verify -- --package tests/fixtures/templates/lists.json --site <reserved-site>/<web>
npm run verify -- --package <path> --mode project --json

# What a web holds of what the handlers write, as JSON
npm run inspect -- --site <reserved-site>/<web> --prefix 'SPJS Itest' --out inspect.json
```

Both only read, so they need no site guard; they take `--site` (default `INTEGRATION_SITE_URL`)
and sign in as `provision` does.

`verify` takes `provision`'s template options (`--package`, `--lang`, `--mode`, `--extensions`,
`--handlers`, `--parameters`, `--skip-taxonomy`) plus `--exact` and `--json`. It replaces the
tokens the way the engine does for each handler and member (the token matrix in
`docs/specs/integration-testing.md`), reads what the template names and reports each difference
as `FINDING [handler] path: expected ..., found ...`, and what it does not compare as `not checked`.
It compares what a re-apply makes true: lists, views, bindings, folders, group members and role
assignments as a superset (exact with `RemoveExistingViews`, `RemoveExistingContentTypes` or
`ClearExistingMembers`); data rows by their `KeyColumn`, the values of `Overwrite` rows only; pages
with `Overwrite: false` and custom actions by existence; menus by their titles, in order, at every
level, and a node's URL only with `IgnoreExisting`. `--exact` also compares the rest, for a web the
template was just applied to. Unlike `provision`, it compares every section by default, browser-only
ones included, so it also checks a site PP365 provisioned.

`inspect` reads the web's own site columns and content types, its site groups, permissions and
associated groups, lists (with views, fields, folders and permissions), custom actions, theme,
pages, navigation, settings and property bag. `--prefix` keeps the named items to those whose
name starts with it. Its JSON holds the tenant's URLs and user names: do not commit it.

## Clean up

```sh
npm run cleanup -- --webs                    # scratch webs and their groups older than 2 h: the plan
npm run cleanup -- --webs --older-than 0m    # also younger ones (only while no run is active)
npm run cleanup -- --prefix 'SPJS Itest'     # prefixed lists, pages, fields, actions, nodes, groups
npm run cleanup -- --webs --apply            # remove what the plan lists
```

`cleanup` prints what it would remove and removes it only with `--apply`. It follows the same site
guard (`--site`, `--allow-site`), deletes a web's subwebs before the web, and never removes the
web's own associated groups (a site titled `SPJS Itest` has `SPJS Itest Owners`).

## What runs in Node

The engine uses three transports: `@pnp/sp` REST and CSOM over HTTP (`src/util/csom.ts`, posted to
`_vti_bin/client.svc/ProcessQuery` through the same pipeline), which work in Node, and JSOM
(`spfx-jsom` -> `SP.ClientContext`), which is browser-only and stubbed out here
(`debug/stubs/spfx-jsom.ts`) until the JSOM exit (`docs/specs/jsom-exit.md`).

| Handler | Runner | Why |
| --- | --- | --- |
| SiteFields, Features, Security, Lists, CustomActions, ComposedLook, ClientSidePages, Navigation, WebSettings, Hooks | runs | REST |
| ContentTypes | skipped by default | runs in Node (CSOM over HTTP); left out until the packages list counts it (JSOM exit, step 5) |
| Taxonomy | skipped | JSOM |
| Files | skipped | refuses whenever an SPFx context is set, which the runner always sets |
| PropertyBagEntries | skipped | rejects in Node |

Skipped handlers are left out unless `--handlers` names them; a forced JSOM handler fails with
"JSOM (spfx-jsom) is not available". Unless ContentTypes is forced, list content type bindings
assume the content types already exist on the site.

## How it maps to PP365

| PP365 (`PackageInstaller`) | Runner |
| --- | --- |
| `_provisionHub` (import) | `--mode hub` (default) |
| `provisionCloudTemplateHubDependencies` (publish as cloud template) | `--mode cloud` |
| The project setup's `ApplyTemplate` and `Hooks` steps (ProjectExtensions) | `--mode project` |
| Taxonomy feature flag off | `--skip-taxonomy` |
| Locale from the hub's language (`_applyLocale`) | `--lang` |
| Sign-in through SPFx | `.env`: cookies locally, certificate in CI (`debug/lib/auth.ts`) |
| The page's SPFx context | a stand-in built from `_api/site` and `_api/web` (`debug/lib/connect.ts`) |

`--mode project` applies, in the setup's order: the project template without `Hooks` and without
each menu (`QuickLaunch`, `TopNavigationBar`) a selected extension declares too, then each extension
without `Hooks`, then the template's hooks and each extension's. The rest of the setup (PreTask,
SitePermissions, SetTaxonomyFields, CopyListData, Planner) is not mirrored, and the template's
`Parameters` are not passed to the engine (the setup does not either).

The sign-in and `.env` variables are ported from PP365's `integration/` harness (releases/1.15),
so the certificate secret and the cookie file serve both repositories.

## Files

| File | Holds |
| --- | --- |
| `provision.ts` | The command |
| `cleanup.ts` | Removes leftovers of runs and tests (a plan unless `--apply`) |
| `verify.ts` | Compares a template with a web (read-only) |
| `inspect.ts` | Prints what a web holds, as JSON (read-only) |
| `lib/env.ts` | `.env`, the sign-in mode, the skip reasons, site URL normalisation |
| `lib/auth.ts` | Cookie and certificate sign-in for PnPjs |
| `lib/connect.ts` | A signed-in client, the SPFx context stand-in, the site guard (`assertWritable`) |
| `lib/apply.ts` | `apply(conn, template, options)`: one template run, with its handlers, errors and warnings |
| `lib/cli.ts` | Arguments, handler plan, package, cloud-subset and project-step resolution (unit-tested in `tests/unit/runner.test.ts`) |
| `lib/verify.ts` | `verify(conn, template, options)`: expectations with the engine's token rules, the diff, the result lines (unit-tested in `tests/unit/verify.test.ts`) |
| `lib/inspect.ts` | `inspectWeb(conn, { prefix })` |
| `lib/logs.ts` | Collects log lines by handler for `--strict` and the tests |
| `lib/scratch.ts` | Scratch webs for the integration tests: names, create, delete with their groups |
| `lib/cleanup.ts` | What `cleanup` finds: scratch leftovers by age, prefixed artefacts |
| `lib/redact.ts` | The one rule for what a run prints or writes where others can read it: without secrets, the tenant, its site, users or ids (`redact`, `redactXml`, `describeError`); `unredacted()` finds what a report still holds (`npm run check:reports`) |
| `lib/placeholders.ts` | The fixtures' `@@user@@` and `@@hookUrl@@`: who the user is, and filling them in |
| `lib/readback/<handler>.ts` | What a handler provisioned, read back for the integration tests |
| `lib/readback/site.ts` | The readers `verify` and `inspect` share: any template's lists, fields, content types, groups and the rest |
| `stubs/spfx-jsom.ts` | Node stand-in for `spfx-jsom` |
