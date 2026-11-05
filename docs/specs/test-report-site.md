# Spec: CI test reports on the test report site

Status: Implemented (2026-11-04: the HTML reporter, the shared publisher with its tests and the
`publish-report` job, checked locally, and revised the same day for F1 and F2 below. The first
publish waits for the setup by hand below.)

Owner: Remi Blom-Ohlsen

Related: [integration testing](./integration-testing.md) (Phase 3 publishes the integration report
the same way); prosjektportalen365 `releases/1.15`: `.tasks/publish-test-report.js` (blob
`5a0af8e8`, from commit `3f0389819`), its tests, `integration/src/reporters/html.ts`, and the
developer guide's «Testrapportsiden» (`.development-guide/ci/kontinuerlig-integrasjon.md`).

## Problem

CI writes the unit tests' JUnit file to the public artifact `test-results`. A browser shows JUnit as
raw XML, so reading a run means downloading the artifact, and it is gone after 7 days. The
integration reports that Phase 3 brings to CI hold tenant values (a site path, scratch web titles,
user names), and this repository is public: its artifacts and logs can be read by anyone signed in
to GitHub.

prosjektportalen365 (PP365) already publishes its end-to-end and integration reports to a private
GitHub Pages site: the `gh-pages` branch of the private repository
`Puzzlepart/prosjektportalen365-test-reports`, readable by the organization's members only. Its
publisher (`.tasks/publish-test-report.js`) knew only PP365's two suites, and it deletes any run
folder under `runs/` that holds neither of them.

## Goals

- The unit report of every push to `main` as a page in the browser, on the same private site, at
  `<site>/sp-js-provisioning/`, linked from the site's front page once PP365 publishes with the
  shared file.
- One publisher, the same file byte for byte in both repositories, tested in this repository's CI.
- No publisher removes or changes the other repository's reports, whichever copy of the file runs.
  This repository's publish changes nothing outside `sp-js-provisioning/`.
- A failed publish says why and shows in the job summary. A push GitHub refuses, or that cannot
  reach GitHub, is told apart from a lost race, waited for longer and points at GitHub's status
  page, as PP365's publisher has done since `3f0389819`.

## Non-Goals

- The integration report. It comes with Phase 3 (`integration.yml`), under one redaction policy for
  the log, the artifact and the site.
- Coverage as a report (it stays in the job log), and reports of pull requests.
- One table for both repositories (a project column or section on the front page): every new
  project would need a lockstep change in both copies, and any version skew would make the shared
  page flip.
- Renaming the report repository or the front page.

## Decisions (2026-11-04, the user)

| | Decision |
| --- | --- |
| R1 | sp-js-provisioning gets its own page in a top-level folder `sp-js-provisioning/` of the existing site, linked from PP365's front page. It never writes under `runs/`, which PP365's publisher prunes. |
| R2 | One generic publisher, written here first and copied byte for byte to PP365 `releases/1.15` in a later pull request; based on PP365's file at `releases/1.15` (blob `5a0af8e8` since `3f0389819`), which it extends and reverts nothing of. It adds a `--project <name>` mode that writes only `<name>/` (and `project.json`), counts read by input rather than by suite name, a guard before the push that refuses any change outside the publisher's own paths, an "Also on this site" line on the front page built from `*/project.json` (by PP365's publishes: F2), a job-summary line when publishing fails, `project` and `repository` in the metadata, and the commit subject from `git log` when `COMMIT_SUBJECT` is empty. Authentication and network errors: F1. |
| R3 | The unit report is published on push to `main` now; the integration report in Phase 3, under one redaction policy for log, artifact and site. |
| R4 | sp-js-provisioning keeps 7 days, as PP365 does; its runs have a budget of 100 MB. |
| R5 | The deploy key is a repository secret, `TEST_REPORTS_DEPLOY_KEY` (as in PP365), with the repository variables `TEST_REPORTS_REPOSITORY` and `TEST_REPORTS_URL`. The user creates them by hand. The publish job runs only on a push to `main`. |
| R6 | The front page keeps its title, "PP365 test reports", and gets the line "Also on this site: sp-js-provisioning" (from PP365's publishes: F2). |
| R7 | PP365's misleading "another job may have published first" is fixed. PP365 fixed it on `releases/1.15` in `3f0389819`, and the shared file keeps that fix (F1). |
| R8 | PP365 runs `test:tasks` in CI, in the pull request job of `ci-build-debug.yml` (part of the later PP365 pull request). |
| F1 | Changed after the review: the shared file follows PP365's `3f0389819` (blob `5a0af8e8`). A push GitHub refuses, or that cannot reach GitHub, is tried again after 10, 20, 30 and 40 s, with its `classifyFailure` and `failureMessage` messages and the `sleep` option; a lost lease keeps the short retries. The later PP365 pull request is then a pure extension of `5a0af8e8`. |
| F2 | sp-js-provisioning writes only its own folder, `sp-js-provisioning/`, and never the front page or `.nojekyll`. The "Also on this site" line is built by the publisher run without `--project` (PP365's mode) from `*/project.json`, so it appears once PP365 takes the shared file. |

The folder is `sp-js-provisioning/`; the report repository keeps its name.

## Constraints

- Pages publishes at most 1 GB for the whole site. PP365's budget for `runs/` is 800 MB; this
  project's is 100 MB.
- Every publish replaces `gh-pages` with one commit without parents, pushed with
  `--force-with-lease` against the commit it fetched. Two publishers are safe together only if both
  push that way and keep the whole fetched tree. PP365's deployed publisher does both (checked
  locally, and in the tests below): it keeps a top-level folder it does not know byte for byte.
- PP365's deployed publisher rewrites the whole front page from `runs/` on every publish, and
  deletes a `runs/<run>/` folder without its suites' metadata.
- GitHub accepts a public key as a deploy key on one repository only: this repository has a key of
  its own on the report repository. A write deploy key can push to every branch there.
- `assertSitePrivate` refuses unless the report repository is invisible to the job's token (404),
  its public Pages address ends at GitHub's sign-in, and `TEST_REPORTS_URL` answers an anonymous
  request with no content. This repository's job token cannot see the private report repository, so
  the check holds unchanged.
- The publisher is dependency-free CommonJS on Node and git (Node 22 here, Node 26 in PP365).

## Approach (as built)

### The HTML reporter

`tests/reporters/html.ts` is PP365's `integration/src/reporters/html.ts` (`releases/1.15`),
ported: the repository root is the working directory; the title is "sp-js-provisioning unit
tests" (or "integration tests", by the run's files); it compiles under TypeScript 4.0 with the
builtins declared in `tests/types/node-builtins.d.ts`; and every text on the page passes through
`redact()` (`debug/lib/redact.ts`) before it is escaped. The reporter runs in the `node --test`
process, past the integration runner's redacted log, so it redacts what it writes itself; a
`redact()` over the finished page would be unsafe, since the `Set-Cookie` rule runs to the next
quote, which on a page is markup. The page embeds its totals in
`<script type="application/json" id="report-summary">`, which the publisher reads.

`node --test` loads a reporter with Node's own loader, which tsx's hooks do not reach, and Node 22
cannot resolve the reporter's extensionless imports itself. `tests/reporters/html.cjs` therefore
loads `html.ts` through tsx's API, and `test:ci` names that file:

```text
--test-reporter=./tests/reporters/html.cjs --test-reporter-destination=test-results/unit/index.html
```

`test-results/unit.xml` stays where it was. The integration runner (`tests/integration/run.ts`) is
not wired: its report is Phase 3 and comes with that phase's redaction.

### The publisher

`.tasks/publish-test-report.js`, the same file as PP365's (see "Sharing the file"). Without
`--project` it does what PP365's deployed file (blob `5a0af8e8`) does, plus the additions below:
the tests check that it publishes the same tree, apart from the front page's line of projects, and
fails with the same attempts, waits and messages. With `--project <name>`:

| Path | Written by | Holds |
| --- | --- | --- |
| `runs/<run id>-<attempt>/<suite>/`, `<suite>.json` | PP365 (no `--project`) | PP365's e2e and integration reports and metadata, as before |
| `index.html` | PP365 (no `--project`) | "PP365 test reports": PP365's runs, and "Also on this site: sp-js-provisioning" once a PP365 publish with this file sees `sp-js-provisioning/project.json` |
| `.nojekyll` | PP365 (no `--project`) | empty |
| `<name>/runs/<run id>-<attempt>/<suite>/`, `<suite>.json` | the project | the report, and its metadata with `project`, `suite`, `title` (the column) |
| `<name>/index.html` | the project | "<name> test reports": its runs of the last 7 days, newest first, a column per suite it published, a link back to the front page |
| `<name>/project.json` | the project | `{ project, title, repository, updatedAt }`: the front page lists the folders whose `project.json` names them |

- **Names.** A project and its suites are lower-case words with hyphens (`^[a-z0-9][a-z0-9-]*$`);
  `runs` is not a project.
- **Retention and budget.** A project drops its own runs older than 7 days, then its oldest while
  its runs are over 100 MB; PP365 drops its runs under `runs/` by its own rules. Neither touches
  the other's.
- **Counts by input.** With `--results`, Playwright's `results.json`; else the `report-summary` of
  the report's `index.html`, whatever the suite is called.
- **The front page line.** A publish without `--project` (PP365's) renders the front page with the
  line, above the table, from the top-level folders whose `project.json` names them. A project's
  publish never writes the front page or `.nojekyll`, so it cannot fall out of step with the page
  PP365's copy renders. Until PP365 publishes with this file, the front page has no line, and
  `<site>/sp-js-provisioning/` is reached from the job summary's link.
- **Ownership guard.** After staging, the paths the commit would change are checked: without a
  project, only `runs/`, `index.html` and `.nojekyll`; with a project, only `<name>/`. Anything
  else refuses the publish without a push and without a retry: "Refused to publish: the commit
  would change ...".
- **Failures**, as in PP365's `3f0389819`, unchanged. A push that lost its lease to another
  publisher's (`[rejected]`, `(stale info)`, `non-fast-forward`, `(fetch first)`) is tried again
  after 1, 2, 3 and 4 s: "Attempt n of 5 lost to another job's push; trying again." A refused
  deploy key or a GitHub that cannot be reached (`Permission denied (publickey)`,
  `Could not read from remote repository`, `Host key verification failed`, an unresolved host, a
  connection timed out, refused, reset or closed, a git call over its 5 minutes) is tried again
  after 10, 20, 30 and 40 s: "Attempt n of 5 could not reach the site repository (<what git
  printed first>); trying again in n0 s." After five attempts the step fails with "The report could
  not be pushed after 5 attempts: could not reach the site repository (...). Check
  https://www.githubstatus.com first; if GitHub is up, check the deploy key and
  TEST_REPORTS_DEPLOY_KEY («Testrapportsiden» in .development-guide/ci/kontinuerlig-integrasjon.md)."
  That guide is PP365's; for this repository the key and the secret are under "Setup by hand"
  below. On 2026-11-04 PP365's run #1404 logged each `Permission denied (publickey)` as a lost
  race, during GitHub's incident with Git operations; `3f0389819` is PP365's fix for that. Five
  attempts that each time out after the SSH connection's 30 s, with their waits, take about
  4 minutes, within the step's 8.
- **The job summary** gets "Not published to the test report site: <reason> The report is only in
  the artifact." when a publish fails, besides the warning annotation. On success it links the
  report and, for a project, the project's page.
- **Metadata.** `run.repository` (from `GITHUB_REPOSITORY`) in every `<suite>.json`; a project's
  also has `project` and `title`. The commit subject is `COMMIT_SUBJECT`, else
  `git log -1 --format=%s $GITHUB_SHA` in the checkout (a run started by hand has none).

### CI (`.github/workflows/ci.yml`)

- The required job "Build and offline tests" keeps its name. Its unit step has `id: unit` (its
  outcome is the job's output `unit`), writes the HTML page beside the JUnit file, and a new step
  runs `npm run test:tasks` (the publisher against local bare repositories; needs git; about 15 s,
  since the tests inject the waits between attempts).
- A second job, `publish-report` ("Publish the unit report to the test report site"):
  `needs: build`, `if: !cancelled()` and a push to `refs/heads/main` of
  `Puzzlepart/sp-js-provisioning`; `permissions: contents: read`; checkout without persisted
  credentials, Node from `.nvmrc`, the `test-results` artifact (`continue-on-error`: it is missing
  when the build failed before the unit tests), no `npm ci`. The publish step runs
  `node .tasks/publish-test-report.js --project sp-js-provisioning --suite unit --title Unit --report test-results/unit --outcome "$OUTCOME"`
  with `continue-on-error: true` and `timeout-minutes: 8`; `TEST_REPORTS_REPOSITORY` and
  `TEST_REPORTS_URL` from repository variables, `TEST_REPORTS_DEPLOY_KEY` from the repository secret
  and `GITHUB_TOKEN` reach it through `env:`. Without them it prints a notice and exits 0.
- Why a job of its own: the required check never waits on a push to another repository, and the
  deploy key is never in a job that runs `npm ci` with third-party install scripts.

### Sharing the file

The publisher and its tests are the same, byte for byte, in both repositories
(`.tasks/publish-test-report.js`, `.tasks/publish-test-report.test.js`), as
`.tasks/check-skills.js` already is. A change is made in one and copied to the other; compare with
`git hash-object .tasks/publish-test-report.js` in both. Until PP365's pull request lands, this
repository's copy is the newer one: it is PP365's current file (blob `5a0af8e8`) with the additions
above, and removes nothing of it. Because each publisher writes only its own paths, an older or
newer copy in either repository cannot harm the other's reports.

The tests run PP365's publishers as deployed on `releases/1.15`, kept byte for byte (a test checks
both ids): `.tasks/fixtures/publish-test-report.b37ae3bb3.js`, deployed from 2026-10-26 until
`3f0389819`, and `.tasks/fixtures/publish-test-report.5a0af8e8.js`, deployed since. Each runs
against the new one in both orders, and the new one without `--project` is compared with
`5a0af8e8`.

## Setup by hand (the user)

An admin of both repositories does this once, from a shell signed in with `gh`. Nothing in CI
creates it. The key is generated in a temporary folder outside the repository and deleted after.

1. A deploy key with write access on the report repository, and its private half as this
   repository's secret:

   ```bash
   d=$(mktemp -d)
   ssh-keygen -t ed25519 -N "" -C "sp-js-provisioning CI test reports" -f "$d/key"
   gh repo deploy-key add "$d/key.pub" --repo Puzzlepart/prosjektportalen365-test-reports --allow-write --title "sp-js-provisioning CI"
   gh secret set TEST_REPORTS_DEPLOY_KEY --repo Puzzlepart/sp-js-provisioning < "$d/key"
   rm -rf "$d"
   ```

2. The two variables, the address read from the report repository's Pages settings:

   ```bash
   gh variable set TEST_REPORTS_REPOSITORY --repo Puzzlepart/sp-js-provisioning --body Puzzlepart/prosjektportalen365-test-reports
   gh variable set TEST_REPORTS_URL --repo Puzzlepart/sp-js-provisioning --body "$(gh api repos/Puzzlepart/prosjektportalen365-test-reports/pages --jq .html_url)"
   ```

3. Check them (names and dates only):

   ```bash
   gh secret list --repo Puzzlepart/sp-js-provisioning
   gh variable list --repo Puzzlepart/sp-js-provisioning --json name --jq '.[].name'
   gh repo deploy-key list --repo Puzzlepart/prosjektportalen365-test-reports
   ```

No Pages setting changes. The report repository's README gets a section on the two projects (its
new text is prepared with this change).

The secret is at repository level, as in PP365: a workflow on a branch of this repository could
read it, which the organization's members with write access can do in PP365 today too. The job
that uses it runs only on a push to `main`.

## Acceptance Criteria

- `npm run test:ci` writes `test-results/unit.xml` (unchanged) and `test-results/unit/index.html`,
  titled "sp-js-provisioning unit tests", with the run's totals in `report-summary`.
- `npm run test:tasks` passes, 35 tests: PP365's 14 (the counts test adapted to counts by input,
  the race test also checking its wait) and the new ones: a project's layout and `project.json`,
  and nothing outside its folder; a column per suite; each publisher prunes only its own runs, by
  age and by budget, and a project's default budget is 100 MB; the line of projects on the front
  page from PP365's publishes, and that page otherwise byte for byte both deployed publishers'; the
  guard refuses a change outside the publisher's paths (for a project also `index.html` and
  `.nojekyll`), without a push or a retry, also on a site's first publish, before it has a branch,
  where a project's writes only its folder; both deployed publishers and the new one keep
  `sp-js-provisioning/` intact, and a project's publish an hour after any of them leaves everything
  outside its folder byte for byte, in both orders; without a project, the same tree as
  `5a0af8e8` (apart from the line) and the same attempts, waits and messages when the site
  repository cannot be reached; a race between the two kinds of publisher keeps both reports; a
  refused deploy key and a timed-out connection are tried 5 times, 10 to 40 s apart, naming
  githubstatus.com; a failed publish writes the job summary; `main` passes its injected wait on
  to the push (a real wait fails the test); the command line warns and exits 1 on an error; the
  commit subject from the checkout; names refused. The waits are injected, so the file runs in
  seconds.
- On a push to `main` after the setup, the `publish-report` job publishes
  `sp-js-provisioning/runs/<run id>-<attempt>/unit/index.html` and the job summary links it. From
  PP365's first publish with the shared file, the front page shows "Also on this site:
  sp-js-provisioning".
- Before the setup, and on every pull request, the required check is green; `publish-report`
  prints the notice (push) or is skipped (pull request).

## Verification Plan

Done locally on 2026-11-04 (no tenant, no GitHub write): `npm run build`, `npm run typecheck:tests`,
`npm run lint:check`, `npm test`, `npm run test:coverage`, `npm run test:tasks`,
`npm run check:pack`, `npm run check:secrets`, `npm run check-skills`; and the plan's
compatibility run against a local bare repository, in both orders, with the real unit report:
PP365's deployed publisher, then this one with `--project`, then both again; the branch one commit,
`runs/` and `sp-js-provisioning/` both whole.

After F1 and F2 (the same day): `npm run test:tasks` twice; PP365's current 14 tests
(`origin/releases/1.15`, test file blob `50afb15d`) against `5a0af8e8` (14 pass) and against the new
file (13 pass: only the counts test fails, as intended, since counts are read by input; it calls
`readCounts('e2e', dir, file)`); and 14 single-string mutants of the new file, among them a project
writing `.nojekyll`, a project rendering the front page again, a project's default budget of
800 MB, the guard letting a project change the front page, and unreachable errors not retried or
waited for 1 s: each fails at least one test.

After the setup and the next push to `main`:

1. The `publish-report` log says `Published sp-js-provisioning/runs/<run id>-1/unit/index.html`, and
   the job summary has "Unit report: Open the report · all runs".
2. `gh api 'repos/Puzzlepart/prosjektportalen365-test-reports/git/trees/gh-pages?recursive=1' --jq '[.tree[].path|split("/")[0]]|unique'`
   lists `runs` and `sp-js-provisioning` side by side, and
   `gh api repos/Puzzlepart/prosjektportalen365-test-reports/pages/builds/latest --jq .status` is
   `built`.
3. `<site>/sp-js-provisioning/` opens when signed in, and a private window lands on GitHub's
   sign-in. The front page is unchanged by the publish (no "Also on this site" yet).
4. After PP365's next publish: the same tree listing still holds `sp-js-provisioning/`, and PP365's
   new run is there. Once that publish uses the shared file, the front page shows "Also on this
   site: sp-js-provisioning".
5. A pull request: `publish-report` is skipped and "Build and offline tests" is green.

## Open Questions

- None for the unit report. For Phase 3: the stronger `redact()` (configured values, UPNs, GUIDs,
  tenant and app id) and a pre-publish leak check, before the integration report is published.

## Rollout Notes

- Order: this change (the job prints the notice until the setup); the setup by hand; the next push
  to `main` publishes. Then the report repository's README (two projects, this project's page and
  its 7 days), and PP365's pull request to `releases/1.15`: the same two files byte for byte (and
  the two fixtures), an extension of `5a0af8e8` that reverts nothing; the developer guide's
  «Testrapportsiden», the stale test counts in `testoversikt.md` and `npm-skript.md`, and
  `test:tasks` in `ci-build-debug.yml`'s pull request job (R8).
- Until PP365 takes the file, the front page has no line: this project's publishes never write it.
  From PP365's first publish with the file, the front page links `sp-js-provisioning/`.
- A report is about 100 KB, so 7 days of pushes to `main` stay far below the 100 MB budget.
- The integration report (Phase 3) uses the same job step in `integration.yml`, with
  `--suite integration --title Integration --report test-results/integration`, once the reporter is
  wired into `run.ts` and the redaction is extended.
