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

npm run provision -- --help
```

| Flag | Does |
| --- | --- |
| `--site <url>` | The target site. Default: `INTEGRATION_SITE_URL`. Must be that site or a web below it, unless `--allow-site` names it. |
| `--allow-site <url>` | Write to a site outside `INTEGRATION_SITE_URL`; the same URL as `--site`, to confirm it (exit 2 otherwise). |
| `--package <path>` | A package folder (with `manifest.json`), a `.pppkg`/`.zip` (extracted with `unzip`), or a raw template JSON. A package resolves to its hub template. |
| `--lang <locale>` | The package's localized hub template and list content, e.g. `en-US`. |
| `--mode hub\|cloud` | `hub`: the whole hub template (default). `cloud`: the subset PP365 applies when a package is published as a cloud template. |
| `--handlers <csv>` | Only these handlers, checked against the engine (exit 2 with the valid names). Also forces handlers that are skipped in Node. |
| `--parameters <k=v,...>` | Values for `{parameter:k}` tokens. |
| `--skip-taxonomy` | Drop the Taxonomy section. |
| `--strict` | Exit 1 when a handler logged an error and carried on, not only when one threw. |
| `--twice` | Apply twice; the second run shows what re-applying does. |
| `--dry-run` | Print the plan and exit, without signing in (no site needed). |

Exit codes: 0 applied, 1 failed (or, with `--strict`, an error was logged), 2 usage error (also a
target the site guard refuses).

`--strict` sees what the handlers log as errors, including the failures they log and carry on
after (a list field, a view's fields, a folder, a row, a page control). It does not read the result
back from the site, so a green strict run is not proof that everything was applied: the first
tenant run found a group owner that SharePoint silently did not change. The integration suites
read back (`npm run test:integration`, `docs/TESTING.md`).

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

The engine uses two transports: `@pnp/sp` REST, which works in Node, and JSOM (`spfx-jsom` ->
`SP.ClientContext`), which is browser-only and stubbed out here (`debug/stubs/spfx-jsom.ts`).

| Handler | Runner | Why |
| --- | --- | --- |
| SiteFields, Features, Security, Lists, CustomActions, ComposedLook, ClientSidePages, Navigation, WebSettings, Hooks | runs | REST |
| ContentTypes, Taxonomy | skipped | JSOM |
| Files | skipped | refuses whenever an SPFx context is set, which the runner always sets |
| PropertyBagEntries | skipped | rejects in Node |

Skipped handlers are left out unless `--handlers` names them; a forced JSOM handler fails with
"JSOM (spfx-jsom) is not available". Because ContentTypes is skipped, list content type bindings
assume the content types already exist on the site.

## How it maps to PP365

| PP365 (`PackageInstaller`) | Runner |
| --- | --- |
| `_provisionHub` (import) | `--mode hub` (default) |
| `provisionCloudTemplateHubDependencies` (publish as cloud template) | `--mode cloud` |
| Taxonomy feature flag off | `--skip-taxonomy` |
| Locale from the hub's language (`_applyLocale`) | `--lang` |
| Sign-in through SPFx | `.env`: cookies locally, certificate in CI (`debug/lib/auth.ts`) |
| The page's SPFx context | a stand-in built from `_api/site` and `_api/web` (`debug/lib/connect.ts`) |

The sign-in and `.env` variables are ported from PP365's `integration/` harness (releases/1.15),
so the certificate secret and the cookie file serve both repositories.

## Files

| File | Holds |
| --- | --- |
| `provision.ts` | The command |
| `cleanup.ts` | Removes leftovers of runs and tests (a plan unless `--apply`) |
| `lib/env.ts` | `.env`, the sign-in mode, the skip reasons, site URL normalisation |
| `lib/auth.ts` | Cookie and certificate sign-in for PnPjs |
| `lib/connect.ts` | A signed-in client, the SPFx context stand-in, the site guard (`assertWritable`) |
| `lib/apply.ts` | `apply(conn, template, options)`: one template run, with its handlers, errors and warnings |
| `lib/cli.ts` | Arguments, handler plan, package and cloud-subset resolution (unit-tested in `tests/unit/runner.test.ts`) |
| `lib/logs.ts` | Collects log lines by handler for `--strict` and the tests |
| `lib/scratch.ts` | Scratch webs for the integration tests: names, create, delete with their groups |
| `lib/cleanup.ts` | What `cleanup` finds: scratch leftovers by age, prefixed artefacts |
| `lib/redact.ts` | Removes cookies, tokens and tenant hosts from what is printed |
| `lib/readback/<handler>.ts` | What a handler provisioned, read back for the integration tests |
| `stubs/spfx-jsom.ts` | Node stand-in for `spfx-jsom` |
