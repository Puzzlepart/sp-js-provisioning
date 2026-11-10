---
name: spjs-provision
description: 'How to apply a template or a Prosjektportalen package to a SharePoint site with the local runner, and how to read the result. Use when asked to "provision", "run the template", "apply the package", "test the template on a site", "npm run provision", "dry run", "why did provisioning fail", "which handlers run", "cloud mode", "project mode", "--strict", "--twice", "--verify", "npm run verify", "what differs from the template", "npm run inspect", "what is on the site", "clean up the site". Encodes the .env sign-in (PP365 e2e cookies locally, certificate in CI), the site guard, the flags, the handlers that run in Node, how the modes map to PP365 PackageInstaller and the project setup, verify and inspect, the exit codes and the log lines. Pair with spjs-testing for automated suites.'
---

# Provisioning a site by hand

Authoritative description: `debug/README.md` (and `docs/ARCHITECTURE.md` for the handlers). This
skill is the operational summary; where it and the code disagree, the code wins and the doc is
fixed.

## Setup

- `.env` (from `.env.example`): `INTEGRATION_SITE_URL` is the reserved site; cookie sign-in reads
  `../prosjektportalen365/e2e/.auth/user.json` (PP365's e2e user, about two days per session:
  `npx playwright test --project=setup` in `prosjektportalen365/e2e`). CI signs in with
  `CI_TENANT`/`CI_CLIENT_ID`/`CI_CERT_BASE64`.
- The account must have access to the site (owner; site collection admin for cleanup).

## Commands

| Command | Does |
| --- | --- |
| `npm run provision -- --package <path> --dry-run` | The plan only: sections, handlers, skips; no sign-in, no site needed |
| `npm run provision -- --package <path>` | Apply to `INTEGRATION_SITE_URL` |
| `... --twice --strict` | Apply twice; exit 1 on any logged error |
| `... --handlers SiteFields,Lists` | Only those (checked against the engine, exit 2 otherwise); also forces browser-only ones |
| `... --mode cloud` | The subset PP365 applies when publishing a cloud template |
| `... --mode project [--extensions <ids\|all\|none>]` | The project template and extensions, as the project setup's `ApplyTemplate` and `Hooks` steps apply them |
| `... --verify [--exact]` | Then compare the template with the site (exit 1 on a finding) |
| `npm run verify -- --package <path> [--exact] [--json]` | Compare a template with a site, read-only (every section, browser-only ones included) |
| `npm run inspect -- [--prefix <text>] [--out <file>]` | What a web holds of what the handlers write, as JSON, read-only |
| `... --lang en-US`, `--skip-taxonomy`, `--parameters k=v,...` | Locale, no Taxonomy, `{parameter:k}` values |
| `... --site <url> --allow-site <url>` | A site outside `INTEGRATION_SITE_URL`: refused unless `--allow-site` repeats it |
| `npm run cleanup -- --prefix 'SPJS Itest'` / `--webs` | What to remove; `--apply` removes it |

`<path>` is a package folder (with `manifest.json`), a `.pppkg`/`.zip`, or a template JSON.
Exit codes: 0 applied, 1 failed (or a logged error with `--strict`, a finding with `--verify`), 2
usage error or refused site. `verify`: 0 no finding, 1 a finding.

## What runs in Node

REST handlers run: SiteFields, Features, Security, Lists, CustomActions, ComposedLook,
ClientSidePages, Navigation, WebSettings, Hooks. Skipped unless forced (`--handlers`): Taxonomy
(JSOM), Files (refuses with an SPFx context), PropertyBagEntries, and ContentTypes, which runs in
Node through CSOM over HTTP but is left out by default until the packages list counts it (JSOM exit,
step 5); list content type bindings then assume the content types already exist.

## Reading the result

- Lines are `(provision) (<Handler>): (<scope>): <message>`; errors carry SharePoint's answer.
- `Failing handler: X` means a `ProvisioningError` stopped the run there; later handlers did not
  run.
- Since 1.5.0 a site column that fails is an error line, not `Failing handler: SiteFields`: the
  run goes on, and only `--strict` exits 1. A navigation node SharePoint refuses still fails the
  run, as `Navigation node '<Title>' (<Url>) was refused by SharePoint: ...`, and both menus are
  left as they were.
- `--strict` sees logged errors, not what SharePoint silently ignored: `--verify` (or
  `npm run verify`) reads the result back and compares it with the template.
- `FINDING [Lists] L/DataRows/0 (Title=A)/Num: expected 1, found 2`: handler, path in the
  template, the value the template gives (tokens replaced as the engine does) and the site's.
  `not checked [...]` lines say what verify does not compare, and why (a `{parameter:}` without a
  value, a DateTime without a time zone, Hooks).
- verify compares what a re-apply makes true: a re-applied template does not change a custom
  action, a page with `Overwrite: false`, a `Skip` row's values or a kept menu node's URL (without
  `IgnoreExisting`), so verify only checks those exist; `--exact` compares them too, for a web the
  template was just applied to.

## Pitfalls

- A 403 on `_api/site` is the sign-in: an expired session, or no access for the e2e account.
- Files and PropertyBagEntries refuse under an SPFx context, and PP365 stops at them too (its
  PackageInstaller strips only Taxonomy). The runner skips them by default; naming them in
  `--handlers` forces them, and the run then stops at that handler.
- NoScript blocks ScriptLink custom actions, property bag and master page writes on modern sites.
- `--mode project` on a package without `provisioning.template` exits 2. The template's menu is
  left out where a selected extension declares the same menu (as the setup does), so verify checks
  that menu against the extension.
- `inspect` output holds tenant URLs and user names: never commit it or paste it into an issue.

## Do not

- Provision a site the user did not name, or use `--allow-site` on your own initiative.
- Run `cleanup --apply` without showing the plan first.
