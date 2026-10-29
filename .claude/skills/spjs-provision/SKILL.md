---
name: spjs-provision
description: 'How to apply a template or a Prosjektportalen package to a SharePoint site with the local runner, and how to read the result. Use when asked to "provision", "run the template", "apply the package", "test the template on a site", "npm run provision", "dry run", "why did provisioning fail", "which handlers run", "cloud mode", "--strict", "--twice", "clean up the site". Encodes the .env sign-in (PP365 e2e cookies locally, certificate in CI), the site guard, the flags, the handlers that run in Node, how the modes map to PP365 PackageInstaller, the exit codes and the log lines. Pair with spjs-testing for automated suites.'
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
| `... --lang en-US`, `--skip-taxonomy`, `--parameters k=v,...` | Locale, no Taxonomy, `{parameter:k}` values |
| `... --site <url> --allow-site <url>` | A site outside `INTEGRATION_SITE_URL`: refused unless `--allow-site` repeats it |
| `npm run cleanup -- --prefix 'SPJS Itest'` / `--webs` | What to remove; `--apply` removes it |

`<path>` is a package folder (with `manifest.json`), a `.pppkg`/`.zip`, or a template JSON.
Exit codes: 0 applied, 1 failed (or a logged error with `--strict`), 2 usage error or refused site.

## What runs in Node

REST handlers run: SiteFields, Features, Security, Lists, CustomActions, ComposedLook,
ClientSidePages, Navigation, WebSettings, Hooks. Taxonomy and ContentTypes (JSOM), Files (refuses
with an SPFx context) and PropertyBagEntries are skipped unless forced; list content type bindings
then assume the content types already exist.

## Reading the result

- Lines are `(provision) (<Handler>): (<scope>): <message>`; errors carry SharePoint's answer.
- `Failing handler: X` means a `ProvisioningError` stopped the run there; later handlers did not
  run.
- `--strict` sees logged errors, not what SharePoint silently ignored: read the result back (the
  suites do; `debug/lib/readback/`), or look at the site.

## Pitfalls

- A 403 on `_api/site` is the sign-in: an expired session, or no access for the e2e account.
- Files and PropertyBagEntries refuse under an SPFx context, and PP365 stops at them too (its
  PackageInstaller strips only Taxonomy). The runner skips them by default; naming them in
  `--handlers` forces them, and the run then stops at that handler.
- NoScript blocks ScriptLink custom actions, property bag and master page writes on modern sites.

## Do not

- Provision a site the user did not name, or use `--allow-site` on your own initiative.
- Run `cleanup --apply` without showing the plan first.
