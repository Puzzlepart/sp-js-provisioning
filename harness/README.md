# Content type ID test harness

This harness lets you validate deterministic content type ID creation on a **fresh site** without needing a full SPFx build.

## Files
- `contenttype-id-harness.js`: Browser-console harness using SharePoint JSOM.
- `contenttype-id-template.sample.json`: Example payload format.

## Quick start
1. Create a fresh site collection / web.
2. Open any modern page on that site while logged in (same tenant).
3. Open browser developer tools → Console.
4. Paste the content of `contenttype-id-harness.js`.
5. Update the `template` object at the bottom with your site URL + desired content types.
6. Run:
   ```js
   await runContentTypeIdHarness(template)
   ```

If successful, the console prints a table where `idMatches` is `true` for every content type.

## Why this is useful
- Reproduces the deterministic-ID flow (`set_id`, parent inference, create/update, re-read verification).
- Works as a low-friction smoke test before/after changes in `src/handlers/contenttypes.ts`.
