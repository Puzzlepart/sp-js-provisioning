# Local provisioning runner

Apply a Prosjektportalen template package to a real SharePoint site straight from
Node — no SPFx build/deploy loop. This reproduces what the catalog's
`PackageInstaller` does in the browser
(`WebProvisioner(web).setup({ spfxContext, logging }).applyTemplate(schema)`),
but authenticates with `@pnp/nodejs` (MSAL app-only) so you can iterate on the
engine handlers (Lists, ContentTypes, SiteFields, Taxonomy, …) and see every log
line — including the SharePoint error payloads — in your terminal.

## Setup (once)

1. **App registration** in Entra ID with SharePoint **application** permission
   `Sites.FullControl.All` (admin consent). Add a client secret or (recommended)
   upload a certificate.
2. **Settings:**
   ```sh
   cp debug/provision.settings.example.ts debug/provision.settings.ts
   ```
   Fill in `tenantId`, `clientId`, and either `clientSecret` or
   `certificate.{thumbprint,privateKeyPath}`. This file is gitignored.
3. **Node 18+** (your system default is fine). The Node 16 requirement only
   applies to the SPFx build, not this runner.

## Run

```sh
# Full hub template from an unzipped package folder (prosjektportalen-hosting)
npm run provision -- \
  --site https://contoso.sharepoint.com/sites/prosjektportalen \
  --package ../prosjektportalen-hosting/packages/pp-testprosjekt

# A built .pppkg (auto-extracted)
npm run provision -- --site <url> --package ../prosjektportalen-hosting/dist/pp-testprosjekt-1.0.0.pppkg

# Only the handlers you're debugging
npm run provision -- --site <url> --package <path> --handlers SiteFields,ContentTypes,Lists

# Cloud-template subset (SiteFields + ContentTypes + binding lists only)
npm run provision -- --site <url> --package <path> --mode cloud

# English hub template, skip taxonomy, or just inspect what would run
npm run provision -- --site <url> --package <path> --lang en-US
npm run provision -- --site <url> --package <path> --skip-taxonomy
npm run provision -- --site <url> --package <path> --dry-run
```

`--package` accepts a package folder (with `manifest.json`), a `.pppkg`/`.zip`
(extracted automatically via `unzip`), or a raw hub-template JSON file.

## What it can and can't provision

The engine uses two transports: `@pnp/sp` REST (works in Node) and JSOM
(`spfx-jsom` → `SP.ClientContext`, **browser-only** — stubbed out here).

| Handler                          | Runner | Why                                  |
| -------------------------------- | ------ | ------------------------------------ |
| **SiteFields**                   | ✅ runs | REST                                  |
| **Lists** (folders, views, rows, content-type bindings) | ✅ runs | REST — this is where the folder / `addAvailableContentType` 500s lived |
| ContentTypes                     | ⏭ skipped | creates CTs via JSOM                |
| Taxonomy                         | ⏭ skipped | term store via JSOM                 |
| Files / PropertyBagEntries       | ⏭ skipped | engine refuses these server-side    |

Skipped handlers are auto-excluded unless you force them with `--handlers`
(forcing a JSOM one fails with a clear "JSOM not available" error). Because
ContentTypes is skipped, list **content-type bindings** assume the content types
already exist on the target site — which is exactly the re-import case the
idempotency fixes target.

## How it maps to the catalog

| Catalog (`PackageInstaller`)        | Runner                                  |
| ----------------------------------- | --------------------------------------- |
| `_provisionHub` (Mode A import)     | `--mode hub` (default)                  |
| `provisionCloudTemplateHubDependencies` | `--mode cloud`                      |
| Taxonomy feature-flag off           | `--skip-taxonomy`                       |
| SPFx auth (page context)            | `@pnp/nodejs` MSAL app-only             |
| SPFx context for `{site}` tokens    | a minimal `spfxContext` stand-in        |

The runner runs the engine **source** (`src/`) directly through `tsx`, so changes
to handlers take effect on the next run without a build.
