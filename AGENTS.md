# AGENTS.md

A thin operational index for coding agents working in sp-js-provisioning. The authoritative docs
are `README.md` (the template format), `docs/ARCHITECTURE.md` (handlers, tokens, context),
`docs/TESTING.md` (how it is tested) and `docs/specs/` (plans and decisions). Where a doc and the
code disagree, the code wins and the doc is fixed in the same change.

## Repository map

| Path | Holds |
| --- | --- |
| `src/webprovisioner.ts` | `WebProvisioner`: loads the context, runs the handlers in `DefaultHandlerSort` order, wraps failures in `ProvisioningError` |
| `src/handlers/` | One handler per template section; `exports.ts` registers them (`Handler`, `DefaultHandlerMap`, `DefaultHandlerSort`) |
| `src/util/` | Tokens (`tokenhelper.ts`, `index.ts` URL tokens), principals (`security.ts`), CSOM over HTTP (`csom.ts`) |
| `src/schema.ts`, `schema.json` | The template contract (TypeScript) and its JSON Schema, generated from it (`npm run build:schema`; CI fails when it is stale) |
| `debug/` | `npm run provision`, `verify`, `inspect` and `cleanup`; `debug/lib/` is the harness they share with the tests |
| `tests/unit/`, `tests/contract/` | Offline tests (`npm test`): logic with fakes, and the requests a handler sends through real PnPjs |
| `tests/reporters/` | The node:test HTML reporter (`npm run test:ci`: `test-results/unit/index.html`) |
| `tests/integration/` | Suites against scratch webs on the reserved site (`npm run test:integration`) |
| `tests/fixtures/templates/` | Fixture templates (`SPJS Itest ...`) |
| `docs/specs/`, `docs/adr/`, `docs/templates/` | Specs (lifecycle Draft → Ready → Implemented → Verified), decision records, templates |
| `.claude/skills/` | Agent skills (the only skill folder; `npm run check-skills`) |
| `.tasks/` | Repository checks run by CI, and the test report publisher (shared with prosjektportalen365, byte for byte) |

## Before you edit

1. PnPjs 4 is a peer dependency and must be a single copy: never add `@pnp/*` to `dependencies`
   (`tests/unit/package.test.ts` guards it).
2. PnPjs 4's `select()` sets the query on the queryable itself. Do not read through a shared web
   instance that something else invokes later; `WebProvisioner` hands each handler a fresh
   `Web(this.web)` for this reason.
3. Many handlers log a failure and carry on. They log it with `log_error`; an apply without a
   thrown error still does not prove everything was applied. Read the result back.
4. `src/` compiles with TypeScript 4.0.7 to ES5 (`lib/`, gitignored, built by `npm run build`):
   no syntax newer than TS 4.0 (`Awaited`, `satisfies`), and `Error` subclasses must restore their
   prototype. Tests and the runner run `src/` through tsx (esbuild), so they can pass while
   `npm run build` fails: run both.
5. Every handler runs in Node and under SPFx alike: there is no JSOM since 1.6.0
   (`docs/specs/jsom-exit.md`). What REST cannot do goes through CSOM over HTTP
   (`src/util/csom.ts`): ContentTypes, Taxonomy, PropertyBagEntries and a group's owner. Taxonomy
   writes to the tenant-wide term store: the runner applies it only with `--taxonomy`, tests write
   only to the group `SPJS Itest`, and nothing is ever deleted from the store.
6. `npm version` publishes immediately (`postversion: tsc && npm publish`). Never run it unless
   asked.
7. `.env` holds the tenant configuration and is gitignored; never commit it or print a cookie,
   token or certificate. `.gitignore` ignores `debug/*` and whitelists files one by one: add new
   `debug/` files to it.

## Commands

| Command | Does |
| --- | --- |
| `npm ci` | Install (Node from `.nvmrc`) |
| `npm run build` | `tsc` → `lib/` |
| `npm test` | Unit and contract tests |
| `npm run typecheck:tests` | Type-check tests, `src/` and `debug/` |
| `npm run lint:check` | eslint (`src/`) and prettier, without rewriting; `npm run lint` rewrites |
| `npm run test:coverage` | Unit tests with coverage floors (they only rise) |
| `npm run check:pack`, `npm run check:secrets`, `npm run check-skills`, `npm run check:schema` | Package contents, committed secrets, skills, `schema.json` up to date (`npm run build:schema` writes it) |
| `npm run test:tasks` | The test report publisher's tests, against local bare repositories |
| `npm run check:tenant` | Is the tenant ready for the integration suites |
| `npm run test:integration:ci`, `npm run check:reports -- test-results` | The suites as CI runs them (strict; JUnit and HTML to `test-results/`), and the check that the reports hold nothing `redact()` would replace |
| `npm run test:integration [-- <Handler ...> \| all \| packages] [--keep]` | Integration suites against scratch webs (`all` adds `everything`; `packages` the hosting packages) |
| `npm run build:fixtures` | Writes `all.json` from the handler fixtures |
| `npm run provision -- --package <path> [--mode hub\|cloud\|project] [--twice] [--strict] [--verify] [--dry-run] [--taxonomy]` | Apply a template or a PP365 package by hand (Taxonomy only with `--taxonomy`) |
| `npm run cleanup -- --webs \| --prefix 'SPJS Itest' [--apply]` | Remove leftovers (a plan unless `--apply`) |
| `npm run verify -- --package <path> [--site <web>] [--exact]` | Compare a template with a web, read-only (exit 1 on a finding) |
| `npm run inspect -- [--site <web>] [--prefix <text>]` | What a web holds, as JSON, read-only |

## Testing

Unit tests run offline in seconds; write one for any logic and for what a handler sends. The
integration suites apply a fixture to a scratch web twice, read the result back, compare and run
`verify`; they need the tenant set up as `docs/TESTING.md` describes (reserved site in `.env`, subsites enabled,
the account a site collection admin). Run them only when the user has asked for tenant work: they
write to the tenant. In GitHub they run only in the `Integration` workflow, which the user starts
by hand from `main` and approves (environment `tenant`); its log and artifact are public, so
everything a run prints or writes goes through `redact()` (`debug/lib/redact.ts`). The
`spjs-testing` skill has the details and the failure triage.

## Skills

Skills live only in `.claude/skills/<name>/SKILL.md` (`npm run check-skills` fails on
`.github/skills` or `.agents/skills`).

| Skill | For |
| --- | --- |
| `spjs-testing` | Writing, running and fixing tests: layers, commands, harness, fixture rules, triage |
| `spjs-provision` | Applying a template or package to a site with `npm run provision`, `verify` and `inspect`, and reading the result |
| `spjs-handlers` | Changing or adding a handler: the contract, throw versus log, tokens, the tests a change needs |
| `spjs-release` | Releasing to npm (the user publishes) and taking a release into PP365 |

## Do not

- Run `provision`, `cleanup --apply` or the integration suites against a site the user did not name,
  or start the `Integration` workflow: the user dispatches and approves it.
- Write a file into `test-results/` that has not passed through `redact()`: the folder becomes a
  public artifact and a page on the report site (`check:reports` stops both when it finds one).
- Delete an assertion, skip a test without a reason, or lower a coverage floor to get green.
- Commit or push unless the user asks; give the user a commit message instead when they prefer that.
