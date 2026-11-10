# Fixtures

Templates the tests apply, in `templates/`. They hold no tenant URL, user or secret: run-time
values come in through `loadFixture(name, { user, hookUrl })` (`@@user@@`, `@@hookUrl@@`) or the
engine's own tokens, used only where the handler replaces them (see the per-handler matrix in
`docs/specs/integration-testing.md`). `npm run provision` and `npm run verify` fill `@@user@@`
with the signed-in user (or `INTEGRATION_USER`) and refuse a template with `@@hookUrl@@`
(`debug/lib/placeholders.ts`).

Conventions: titles start with `SPJS Itest`, internal names with `SPJSItest`, site column groups
are `SPJS Itest`, and site groups are named `{sitename} ...` so they carry the scratch web's title
and are deleted with it. GUIDs are fixed and lower case; no GUID is shared between fixture
families (`<name>.json` and its `<name>.b.json` variant), except tenant-invariant built-ins.

| Fixture | Suite | Covers |
| --- | --- | --- |
| `smoke.json` | the 0a runner check (`npm run provision -- --package tests/fixtures/templates/smoke.json --twice --strict`) | A site column, two site groups (one owns itself), a list granted to them |
| `failures.json` | the 0b runner check (`--handlers Lists\|Files\|PropertyBagEntries --strict`); `sitefields.test.ts` (`SiteFields` only) | A list field SharePoint refuses, a `Files` entry whose source is not there and a `PropertyBagEntries` entry (refused where custom script is blocked): each run must exit 1. A site column SharePoint refuses (a formula naming a missing column) followed by a valid one: one error, and the valid column is still created |
| `contenttypes.json`, `.b.json` | `contenttypes.test.ts` | Five site columns, a content type below Item with four field refs in an order of their own (`Required`, `Hidden`, both given as false), its child (`<id>01`, a field ref by name only), a child of Document (a field ref by id only), a list bound to the first; B renames the first, adds a link before the others and makes a column no longer required |
| `taxonomy.json`, `.b.json` | `taxonomy.test.ts` | The term group `SPJS Itest` with one term set of four terms (custom properties, a sort order, a description, closed for new terms). It lives in the tenant's term store, is created once and is never deleted (docs/specs/jsom-exit.md, J5). B reverses the sort order and, with `UpdateExistingTerms`, gives a term another value for a property it has |
| `files.json` | `files.test.ts` | Three files into `SiteAssets` (a library every web has under that URL, whatever its language): a text file from the scratch web (`{parameter:SPJSItestFilesWeb}`) with a `Title`, a binary file and a text file from a local server (`{parameter:SPJSItestFilesServer}`), the last with `Overwrite: false` |
| `propertybagentries.json` | `propertybagentries.test.ts` | Two entries: an indexed one without `Overwrite`, one with `Overwrite` |
| `security.json` | `security.test.ts` | Site groups (owners, join/leave settings, `ClearExistingMembers`), web role assignments, a list granted to one of the groups |
| `sitefields.json`, `.b.json` | `sitefields.test.ts` | Eleven column types; B renames one and makes it required. The suite also applies inline templates: a narrowing type change, a taxonomy column, a lookup to a list of the same template with a FieldRef copy |
| `features.json`, `.b.json` | `features.test.ts` | Activate / deactivate the FollowingContent web feature |
| `lists.json`, `.b.json` | `lists.test.ts` | A list and a library: bindings, list fields, field refs, views, folders, list security; B changes two list fields' display names, a description and a `MaxLength` |
| `lists-datarows.json`, `.b.json` | `lists-datarows.test.ts` | Every DataRows value shape, `KeyColumn`, `Skip` / `Overwrite`; B: one list per case the Lists re-apply spec fixed |
| `customactions.json`, `.b.json`, `customactions-script.json` | `customactions.test.ts` | SPFx extension actions; B changes a `Sequence`; classic actions (need custom script) |
| `composedlook.json` | `composedlook.test.ts` | An OOTB look (palette, font scheme, background image) |
| `clientsidepages.json`, `.b.json` | `clientsidepages.test.ts` | Pages with sections, a text control, OOTB web parts, `Overwrite` on and off |
| `navigation.json`, `.b.json` | `navigation.test.ts` | Quick Launch three levels deep and the top bar; B: `IgnoreExisting` at the top level and on a child, a kept node without its children |
| `websettings.json`, `.b.json` | `websettings.test.ts` | Welcome page and web properties; B without a welcome page |
| `hooks.json` | `hooks.test.ts` | GET and POST hooks to a local listener, a 202 that is polled |
| `parameters.json` | `parameters.test.ts` | `{parameter:}` in site column and list field XML and in DataRows values, the `KeyColumn` value included (with an apostrophe) |
| `all.json` | `everything.test.ts` (`npm run test:integration -- all`); `npm run provision -- --package tests/fixtures/templates/all.json --twice --strict --verify` | Every handler fixture that needs no setup of its own in one template, in run order, plus the list `sitefields.json`'s Lookup points at. Made from the fixtures by `npm run build:fixtures` (`tests/fixtures/fixtures.ts`, `composeAll`); `tests/unit/fixtures.test.ts` fails when it is out of date. Without `taxonomy.json` (the tenant's term store), `files.json` (a local server and a source in the web), `propertybagentries.json` and `customactions-script.json` (custom script), `hooks.json` (a local listener), `failures.json`, `smoke.json` and the `.b` variants |

Built-in ids the fixtures use (the same on every tenant):

| Id | What |
| --- | --- |
| `a7a2793e-67cd-4dc1-9fd0-43f61581207a` | FollowingContent web feature |
| `fa564e0f-0c70-4ab9-b863-0177e6ddd247` | The Title column |
| `c70391ea-0b10-4ee9-b2b4-006d3fcad0cd` | Quick links web part |
| `d1d91016-032f-456d-98a4-721247c305e8` | Image web part |

The other GUIDs (site columns, list fields, SPFx component ids) are made up for the fixtures and
listed in each file; SharePoint stores any GUID as a custom action's component id.
