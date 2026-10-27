# Fixtures

Templates the tests apply, in `templates/`. They hold no tenant URL, user or secret: run-time
values come in through `loadFixture(name, { user, hookUrl })` (`@@user@@`, `@@hookUrl@@`) or the
engine's own tokens, used only where the handler replaces them (see the per-handler matrix in
`docs/specs/integration-testing.md`).

Conventions: titles start with `SPJS Itest`, internal names with `SPJSItest`, site column groups
are `SPJS Itest`, and site groups are named `{sitename} ...` so they carry the scratch web's title
and are deleted with it. GUIDs are fixed and lower case; no GUID is shared between fixture
families (`<name>.json` and its `<name>.b.json` variant), except tenant-invariant built-ins.

| Fixture | Suite | Covers |
| --- | --- | --- |
| `smoke.json` | the 0a runner check (`npm run provision -- --package tests/fixtures/templates/smoke.json --twice --strict`) | A site column, two site groups (one owns itself), a list granted to them |
| `failures.json` | the 0b runner check (`--handlers Lists\|Files\|PropertyBagEntries --strict`) | A list field SharePoint refuses, a `Files` and a `PropertyBagEntries` entry: each run must exit 1 |
| `security.json` | `security.test.ts` | Site groups (owners, join/leave settings, `ClearExistingMembers`), web role assignments, a list granted to one of the groups |
| `sitefields.json`, `.b.json` | `sitefields.test.ts` | Eleven column types; B renames one and makes it required |
| `features.json`, `.b.json` | `features.test.ts` | Activate / deactivate the FollowingContent web feature |
| `lists.json`, `.b.json` | `lists.test.ts` | A list and a library: bindings, list fields, field refs, views, folders, list security; B changes two list fields' display names, a description and a `MaxLength` |
| `lists-datarows.json`, `.b.json` | `lists-datarows.test.ts` | Every DataRows value shape, `KeyColumn`, `Skip` / `Overwrite`; B: one list per case the Lists re-apply spec fixed |
| `customactions.json`, `.b.json`, `customactions-script.json` | `customactions.test.ts` | SPFx extension actions; B changes a `Sequence`; classic actions (need custom script) |
| `composedlook.json` | `composedlook.test.ts` | An OOTB look (palette, font scheme, background image) |
| `clientsidepages.json`, `.b.json` | `clientsidepages.test.ts` | Pages with sections, a text control, OOTB web parts, `Overwrite` on and off |
| `navigation.json`, `.b.json` | `navigation.test.ts` | Quick Launch with children and the top bar; B: `IgnoreExisting` |
| `websettings.json`, `.b.json` | `websettings.test.ts` | Welcome page and web properties; B without a welcome page |
| `hooks.json` | `hooks.test.ts` | GET and POST hooks to a local listener, a 202 that is polled |

Built-in ids the fixtures use (the same on every tenant):

| Id | What |
| --- | --- |
| `a7a2793e-67cd-4dc1-9fd0-43f61581207a` | FollowingContent web feature |
| `c70391ea-0b10-4ee9-b2b4-006d3fcad0cd` | Quick links web part |
| `d1d91016-032f-456d-98a4-721247c305e8` | Image web part |

The other GUIDs (site columns, list fields, SPFx component ids) are made up for the fixtures and
listed in each file; SharePoint stores any GUID as a custom action's component id.
