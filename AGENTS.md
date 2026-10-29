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
| `src/schema.ts`, `schema.json` | The template contract (TypeScript) and its JSON Schema (hand-maintained; Phase 2 regenerates it) |
| `debug/` | `npm run provision` and `npm run cleanup`; `debug/lib/` is the harness they share with the tests |
| `tests/unit/` | Offline tests (`npm test`) |
| `tests/integration/` | Suites against scratch webs on the reserved site (`npm run test:integration`) |
| `tests/fixtures/templates/` | Fixture templates (`SPJS Itest ...`) |
| `docs/specs/`, `docs/adr/`, `docs/templates/` | Specs (lifecycle Draft → Ready → Implemented → Verified), decision records, templates |
| `.claude/skills/` | Agent skills (the only skill folder; `npm run check-skills`) |
| `.tasks/` | Repository checks run by CI |

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
5. `spfx-jsom` cannot load in Node; `tests/tsconfig.json` and `debug/tsconfig.json` map it to
   `debug/stubs/spfx-jsom.ts`. ContentTypes, Taxonomy, Files and PropertyBagEntries are therefore
   browser-only until the JSOM exit.
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
| `npm test` | Unit tests |
| `npm run typecheck:tests` | Type-check tests, `src/` and `debug/` |
| `npm run lint:check` | eslint (`src/`) and prettier, without rewriting; `npm run lint` rewrites |
| `npm run test:coverage` | Unit tests with coverage floors (they only rise) |
| `npm run check:pack`, `npm run check:secrets`, `npm run check-skills` | Package contents, committed secrets, skills |
| `npm run check:tenant` | Is the tenant ready for the integration suites |
| `npm run test:integration [-- <Handler ...>] [--keep]` | Integration suites against scratch webs |
| `npm run provision -- --package <path> [--twice] [--strict] [--dry-run]` | Apply a template or a PP365 package by hand |
| `npm run cleanup -- --webs \| --prefix 'SPJS Itest' [--apply]` | Remove leftovers (a plan unless `--apply`) |

## Testing

Unit tests run offline in seconds; write one for any logic and for what a handler sends. The
integration suites apply a fixture to a scratch web twice, read the result back and compare; they
need the tenant set up as `docs/TESTING.md` describes (reserved site in `.env`, subsites enabled,
the account a site collection admin). Run them only when the user has asked for tenant work: they
write to the tenant. The `spjs-testing` skill has the details and the failure triage.

## Skills

Skills live only in `.claude/skills/<name>/SKILL.md` (`npm run check-skills` fails on
`.github/skills` or `.agents/skills`).

| Skill | For |
| --- | --- |
| `spjs-testing` | Writing, running and fixing tests: layers, commands, harness, fixture rules, triage |
| `spjs-provision` | Applying a template or package to a site with `npm run provision`, and reading the result |

## Do not

- Run `provision`, `cleanup --apply` or the integration suites against a site the user did not name.
- Delete an assertion, skip a test without a reason, or lower a coverage floor to get green.
- Commit or push unless the user asks; give the user a commit message instead when they prefer that.
