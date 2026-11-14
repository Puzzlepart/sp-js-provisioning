---
name: spjs-release
description: 'How sp-js-provisioning is released to npm and taken into Prosjektportalen 365. Use when asked to "release", "publish", "cut 1.x.y", "npm version", "prepare the changelog for a release", "bump sp-js-provisioning in PP365", "update the pin", "rush update", "is it safe for 1.14", "npm whoami 401". Encodes who runs what (the user publishes), the checks before a release, the CHANGELOG format, npm sign-in and 2FA, npm version publishing at once (postversion), the tag push, and the PP365 bump (exact pins in two extensions, rush:update with Node 26, a lock kept to the engine on a release branch, SBOM, release notes, the CI tag), plus which PP365 versions can take which release. Pair with spjs-handlers for the change itself.'
---

# Releasing the engine

Authoritative description: `package.json` (`postversion`), `CHANGELOG.md`, and in prosjektportalen365
the `pp365-toolchain` skill (the pin and the lock). This skill is the operational summary; where it
and the code disagree, the code wins and the doc is fixed.

## Who does what

| Step | Who |
| --- | --- |
| CHANGELOG, checks, commits | the agent (commits only when the user asks; never pushes) |
| `npm login`, `npm version`, `git push --follow-tags` | the user |
| PP365 bump (pins, lock, SBOM, notes), its tests and commit | the agent, after the release is on npm |
| Push to PP365 | the user |

`npm version` publishes at once: `postversion` is `tsc && npm publish`. Never run it unless the
user asks for that command.

## Before a release

1. CHANGELOG: move `[Unreleased]` into `## [x.y.z] - YYYY-MM-DD` with a short intro (what a
   consumer notices), `### Fixed` / `### Changed` / `### Development`, and the compare links at the
   bottom (`[Unreleased]: .../compare/vx.y.z...HEAD`, `[x.y.z]: .../compare/v<previous>...vx.y.z`).
   The date is the day it is published.
2. Green: `npm run build`, `npm test`, `npm run typecheck:tests`, `npm run lint:check`,
   `npm run test:coverage`, `npm run check:schema`, `npm run check:pack`, `npm run check:secrets`,
   `npm run check-skills`, `npm run test:tasks`.
3. `git diff v<previous> -- package.json`: runtime `dependencies` and `peerDependencies` decide
   whether consumers can take it as a drop-in (devDependencies do not reach them).
4. A clean working tree (`npm version` refuses otherwise).

## Publishing (the user)

```bash
npm whoami
npm version patch
git push --follow-tags
```

- `npm whoami` answering E401: the token in `~/.npmrc` expired or was revoked. `npm login` (web
  sign-in) replaces it; its session lasts a few hours, so publish right after.
- `npm version` builds (`prepack`), makes the commit `x.y.z` and the tag `vx.y.z`, then
  `npm publish` asks for 2FA in the browser or a one-time code.
- If the publish fails after the commit and tag were made, do not run `npm version` again (that
  makes the next version): fix the sign-in and run `npm publish`.
- The registry can take a few minutes to serve a new version (`npm view sp-js-provisioning@x.y.z`).
- No GitHub Releases are made for this repository; the tag is enough.

## Taking it into PP365

| Item | Where |
| --- | --- |
| Exact pins | `SharePointFramework/PortfolioExtensions/package.json`, `SharePointFramework/ProjectExtensions/package.json` |
| Lock | `npm run rush:update` with Node 26 (PP365's `.nvmrc`; `rush.json` refuses other majors) |
| SBOM | `npm run generate-sbom` (only the date and the engine's lines should change) |
| Mentions | CHANGELOG (the release's section), `releasenotes/<version>.md`, `.development-guide/maler/js-provisjoneringsmal.md` and `.development-guide/README.md` ("låst til versjon"), the `pp365-templates` and `pp365-toolchain` skills |
| Tests | `npm test` (heft) in ProjectExtensions and PortfolioExtensions: they build too |
| Commit | `chore: sp-js-provisioning x.y.z [apps-only:PortfolioExtensions,ProjectExtensions]` |

The lock diff should hold only the engine and what its own dependencies add or drop. A bump makes
pnpm re-resolve the SPFx build rig, which can drop or add peer suffixes elsewhere (1.5.1: the
`(postcss@8.5.28)` suffix on webpack, about 60 lines). On a release branch keep only the engine's
hunks: save `git diff -U3` of the lock, check out the committed lock, `git apply` the hunks that name
`sp-js-provisioning`, then `rush install` must pass.

## Which PP365 can take which release

- PP365 1.14.x pins 1.3.15 (PnPjs 3). 1.4.0 and later need PnPjs 4 (PP365 1.15): they cannot go into
  a 1.14 hotfix; a v3-era consumer fails every `new WebProvisioner(web)` with "IWeb is not
  assignable".
- 1.6.0 (the JSOM exit, 2026-11-14) no longer brings `spfx-jsom`: a consumer that imports it declares
  it itself (PP365's ProjectExtensions, ProjectWebParts and shared-library do). PP365 1.15 takes it;
  check a project setup and a hub import of `pp-testprosjekt`, twice, on the test channel.
- prosjektportalen-hosting packages declare `minPPVersion`; a package for 1.14.0 must still work
  with 1.3.15, so a template property that needs a newer engine raises `minPPVersion`.
- Error messages, log levels and re-apply behaviour are what PP365 users see in the setup and the
  package import: say them in the CHANGELOG intro, and in PP365's CHANGELOG and release notes.

## Pitfalls

- 2026-11-10 (1.5.1): `npm whoami` gave E401 with an old token in `~/.npmrc`; `npm login` fixed it.
- 2026-11-10: `rush update` changed about 60 unrelated lines in the lock; a `rush update` with
  nothing changed leaves the lock alone, so it is the bump that triggers the re-resolution.
- The user's commits use the repository's own git identity: no `--author`, keep the
  `Co-Authored-By` trailer.

## Verification checklist

- [ ] CHANGELOG section, date and compare links
- [ ] All checks green; the runtime dependencies and peers diffed against the previous tag
- [ ] After publishing: `npm view sp-js-provisioning@x.y.z version`, the tag on GitHub, CI green
- [ ] PP365: both pins, the lock with only the engine's hunks, SBOM, mentions, heft tests green

## Do not

- Run `npm version`, `npm publish` or `git push` yourself unless the user asks for it.
- Run `provision`, `cleanup` or the integration suites against a site the user did not name, or
  without the user's go-ahead; never print a cookie, token or certificate, or commit `.env`.
- Delete an assertion, skip a test without a reason, or lower a coverage floor to get green.
- Put a 1.4-or-later release into a PP365 1.14 branch.
