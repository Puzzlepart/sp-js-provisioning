# Deterministic Content Type ID Smoke Test

This smoke fixture verifies deterministic content type ID provisioning against a
caller-provided SharePoint site. It intentionally contains no tenant URLs,
credentials, app IDs, or user-specific values.

## Runtime

Run this from a consumer SPFx/debug environment where:

- `sp-js-provisioning` is installed from the branch or packed tarball under test.
- A PnP `IWeb` instance points at a fresh test web.
- The current user or app principal can create site fields, content types, and
  document libraries.

The content type handler uses `spfx-jsom`, which loads SharePoint JSOM scripts
with `SPComponentLoader`. Because of that, this is not a plain Node.js console
test.

## Scenario

Use `template.json` with:

1. A fresh test web.
2. `WebProvisioner.applyTemplate(template, ['SiteFields', 'ContentTypes', 'Lists'])`.
3. The same command a second time against the same web.

The first run verifies creation with explicit deterministic IDs. The second run
verifies that existing content types are updated instead of recreated.

## Expected Result

- Site content type `SPJS Smoke Document` exists with ID
  `0x0101008F279D4FF9A5427D8FA8FE8A5F8BD9E`.
- Site content type `SPJS Smoke Child Document` exists with ID
  `0x0101008F279D4FF9A5427D8FA8FE8A5F8BD9E00BDA84E8A7A5F41EFB78D473EC5AF1B86`.
- List `SPJS Smoke Documents` exists and has content type management enabled.
- The list content type IDs for both smoke content types start with the
  configured site content type IDs, proving SharePoint created inherited list
  content type IDs.
- The second run completes without duplicate content type creation errors.

## Console Shape

Adapt this to the local consumer app or SPFx workbench context:

```ts
import { WebProvisioner } from 'sp-js-provisioning'
import template from './template.json'

await new WebProvisioner(web)
  .setup({ logging: { prefix: 'smoke' } })
  .applyTemplate(template, ['SiteFields', 'ContentTypes', 'Lists'], console.log)
```
