---
doc-schema-version: 1
title: "Release Workflow"
summary: "Maintained release path for publishing the scoped ESM-only package and its unscoped alias."
read_when:
  - Preparing a release
  - Updating npm publish or GitHub Release workflow documentation
  - Checking release verification steps
doc_type: "runbook"
---

# Release Workflow

This guide documents the maintained release path for the scoped ESM-only
package, `@shbernal/ts-pptx`, and for `pptx-ts`, the unscoped alias published
beside it.

The alias is not a redirect. npm has one package per name and no forwarding, so
an alias is a second publish of the same content under a second name: same
version, same `dist/`, same `skills/`, one line of `package.json` different and
the dev-only `scripts` block removed.
`scripts/alias-package.mjs` stages that copy and the publish workflow pushes it
after the canonical one. Nothing about cutting a release changes because of it,
and no step below is for the alias alone except the one-time bootstrap.

Publishing is automated by `.github/workflows/publish.yml`. The workflow runs
when a GitHub Release is published and can also be dispatched manually from a
matching tag as a retry path.

## Release Prerequisites

- The npm package is `@shbernal/ts-pptx`, and `pptx-ts` is its unscoped alias.
- `package.json#repository.url` points at `shbernal/ts-pptx`.
- npm trusted publishing is configured **once per package name**, because npm
  exchanges the OIDC token per package. Both configurations are identical apart
  from the package they sit on:
  - package: `@shbernal/ts-pptx`, and again for `pptx-ts`
  - GitHub repository: `shbernal/ts-pptx`
  - workflow filename: `publish.yml`
  - GitHub environment: `npm-publish`
  - allowed action: `npm publish`
- The GitHub Environment `npm-publish` exists before the first automated
  release.
- Each name exists on the registry already. A trusted publisher cannot be
  configured for a package that does not exist yet, so a *new* name needs one
  manual publish first: see "Bootstrapping a New Package Name" at the end of this
  page. That is the only sanctioned local `npm publish` in this project, and it
  is not part of any release.
- Do not add an `NPM_TOKEN` secret for the normal path. The workflow uses OIDC
  with `id-token: write`.

## Version Updates

`package.json` holds the version of record: the publish workflow refuses to
publish unless the tag matches it. The `VERSION` constant in
`src/presentation.ts` that backs `pres.version` is *derived* from it and is not
edited by hand.

Write the `CHANGELOG.md` entry first (release date and summary), stage it, then
bump:

```bash
git add CHANGELOG.md
pnpm version minor --message 'chore(release): v%s' --no-git-checks
```

That bumps `package.json`, runs the `version` lifecycle script
(`scripts/sync-version.mjs`) which rewrites the constant and stages it, then
makes one commit holding all three files and creates the annotated tag `vX.Y.Z`.
Both flags are load-bearing:

- **`--message`**: pnpm's default subject is the bare version (`3.2.0`), which
  is not this repo's commit style. It does **not** read npm's `message` config;
  setting `message` in an `.npmrc` is silently ignored, so the flag is the only
  way to control the subject.
- **`--no-git-checks`**: pnpm otherwise refuses to run against anything but a
  spotless tree (`ERR_PNPM_UNCLEAN_WORKING_TREE`), and the staged `CHANGELOG.md`
  counts. Waiving the check is what keeps the release a single commit rather than
  two. It waives the check for *everything*, so run `git status` first and leave
  scratch files untracked.

Still by hand, and unaffected by the above:

- Demo package versions, when they intentionally track the release version.
  (Neither does today: `demos/node` is on 5.0.2 and `demos/showcases` on 1.0.0,
  so this is normally a no-op.)
- Keep package import examples on the scoped package name:
  `@shbernal/ts-pptx`.

If the constant ever does drift (a hand-edited `package.json`, or a bump made
some other way), `pnpm run version:check` reports it and `pnpm run version:sync`
repairs it. `test/regression/api/public-accessors.test.js` fails in `verify`
either way, and the release path cannot skip that, so a mis-reported version
cannot ship.

## Local Release Gate

Install dependencies (the OOXML oracle needs no install step; `ooxml-validate`
fetches it on first use):

```bash
pnpm install --frozen-lockfile
```

Run the full automated gate before tagging:

```bash
pnpm run check:static
pnpm run verify:full
```

Together these are what CI runs. `check:static` adds `lint` and `format:check`,
which `verify:full` deliberately omits because the git hooks own them, but a
release is exactly the moment to confirm them explicitly.

Check that the target version is not already published, under either name:

```bash
npm view @shbernal/ts-pptx@X.Y.Z version
npm view pptx-ts@X.Y.Z version
```

Both commands should fail with a registry 404 for a new release version.

## Automated npm Publish

1. Merge the release commit into `master`.
2. Push `master` and the `vX.Y.Z` tag that `pnpm version` created: it already
   matches `package.json#version`, which is what the workflow checks:

   ```bash
   git push origin master && git push origin vX.Y.Z
   ```

3. Create a GitHub Release from `vX.Y.Z`.
4. Publish the GitHub Release.

Publishing the GitHub Release starts `.github/workflows/publish.yml`. The
workflow:

- refuses to run outside `shbernal/ts-pptx`
- refuses branch publishes; `GITHUB_REF_TYPE` must be `tag`
- requires the tag name to equal `v${package.json#version}`
- checks that `@shbernal/ts-pptx@X.Y.Z` and `pptx-ts@X.Y.Z` are unpublished,
  and fails only when **both** already exist (see "Retrying a half-published
  release" below)
- installs with `pnpm install --frozen-lockfile`
- installs the OOXML validator
- runs lint, formatting, typecheck, tests, and the package boundary checks
  (`package:lint` + `test:package`)
- publishes `@shbernal/ts-pptx` with
  `npm publish --access public --provenance --ignore-scripts`
- stages the alias with `node scripts/alias-package.mjs --out .tmp/alias-package`
  and publishes that directory the same way

npm trusted publishing automatically exchanges the GitHub Actions OIDC token for
publish credentials. The explicit `--provenance` flag keeps provenance required
even if npm defaults change.

## Manual Workflow Retry

Use this only after fixing a failed publish workflow without changing the
release artifact:

```bash
gh workflow run publish.yml --repo shbernal/ts-pptx --ref vX.Y.Z
```

The selected ref must be the release tag, not `master`.

### Retrying a half-published release

Publishing two names is not atomic: the canonical publish can succeed and the
alias fail after it, leaving the version on npm under one name only. The
workflow is written to be re-runnable in exactly that state. Its guard fails
only when **both** names already carry the version, and each publish step is
skipped when its own name already does, so re-dispatching on the same tag
finishes the missing half and touches nothing else.

The ordering is what makes this safe rather than merely convenient. The alias is
published last, so a failure in it can never take down a release that has
already gone out; the worst case is a version that exists under the scoped name
and needs one re-dispatch to appear under the alias. Do not reach for a local
`npm publish` here. The retry path is the workflow.

## Post-Publish Checks

Verify npm and GitHub agree on the release:

```bash
npm view @shbernal/ts-pptx@X.Y.Z version dist-tags --json
npm view pptx-ts@X.Y.Z version dist-tags --json
gh release view vX.Y.Z --repo shbernal/ts-pptx
```

Then say on each issue the release closes **which version carries the fix**:

```bash
gh issue comment <N> --repo shbernal/ts-pptx --body "Released in X.Y.Z."
```

One line each, and it is the only place a consumer can learn it. Issues here close
when the fix merges, which is the right moment for this repo and the wrong signal
for a consumer: merged and unreleased is a state that can last weeks, and a
workaround deleted on the strength of a closed issue breaks against the version
that is actually installed. The skill this package ships tells consumers to trust
the published version over the issue state for exactly that reason: this comment
is what makes the two agree. `CHANGELOG.md` already cites the numbers, so the list
is the entry you just wrote.

### One downstream to watch, and it is not a blocker

The site's demos page renders its preview with [`pptx-html`](https://www.npmjs.com/package/pptx-html),
which depends on `@shbernal/ts-pptx` at a caret range and therefore installs its **own**
published copy: see `www/README.md` for why that duplication is deliberate rather than an
oversight.

The consequence lands on a **major**. When this package goes 4.x, `pptx-html`'s reader is
still built against 3.x, so the preview keeps rendering decks the *old* writer produced and
the docs build keeps resolving, but the page stops demonstrating the version it sits beside
until `pptx-html` ships a matching release. Nothing in this repo's gates detects that, because
nothing here asserts what the preview looks like (docs/testing.md, "Demos Are Not Tests").

So: release, then open an issue on `pptx-html`. Do not hold a release for it, and do not
pin the site to the workspace copy to avoid it: that trade makes the first breaking change
break the docs *deploy* of the release introducing it, which is strictly worse.

## Package Surface Checks

The package should ship:

- `dist/index.js`
- `dist/index.d.ts`
- `dist/inspect.js`
- `dist/measure.js`
- `dist/read.js`
- `dist/script.js`
- `dist/math.js`
- `dist/zip.js`
- `dist/html.js`
- `dist/node.js`
- `dist/browser.js`
- package `exports["."].default`
- package `exports["."].types`
- package subpaths for `./inspect`, `./measure`, `./read`, `./script`,
  `./math`, `./zip`, `./html`, `./node`, and `./browser`
- scoped imports for `@shbernal/ts-pptx`,
  `@shbernal/ts-pptx/inspect`, `@shbernal/ts-pptx/measure`,
  `@shbernal/ts-pptx/read`, `@shbernal/ts-pptx/script`,
  `@shbernal/ts-pptx/math`, `@shbernal/ts-pptx/zip`,
  `@shbernal/ts-pptx/html`, `@shbernal/ts-pptx/node`, and
  `@shbernal/ts-pptx/browser`

(`pnpm run test:package` exercises all ten end-to-end, out of an installed
tarball: every one is imported and checked for a sample of load-bearing named
exports, and all but `./browser` are additionally put through esbuild on the
`node` platform. Both lists come off one `EXPORT_MATRIX` in
`scripts/package-smoke.mjs`, so a new subpath is covered by being added there,
but it still has to be added *here* by hand, and this list has drifted before.)

The package should not ship or document:

- CommonJS support
- IIFE/global browser bundle support
- direct CDN script-tag support as a maintained workflow
- `types/`
- `src/bld/`
- `dist/pptxgen.cjs.js`
- `dist/pptxgen.js`
- `dist/pptxgen.es.js`
- `dist/pptxgen.bundle.js`
- `dist/pptxgen.min.js`

## Bootstrapping a New Package Name

Read this only when the project starts publishing under a name it has never
published under before. It ran once for `@shbernal/ts-pptx` and once for
`pptx-ts`, and it is not part of cutting a release.

npm configures a trusted publisher on a package's settings page, which means the
package has to exist before OIDC can be enabled for it
([npm/cli#8544](https://github.com/npm/cli/issues/8544)). That is a genuine
chicken and egg: the workflow cannot create the package, and the package cannot
be created by the workflow. It is broken with one local publish, and then never
again for that name.

1. Confirm the name is free. A 404 is the answer you want:

   ```bash
   npm view <name> version
   ```

   npm also rejects a new name that is too similar to an existing one, and that
   check runs at publish time rather than here. If step 3 fails on similarity,
   the name is not available and npm support is the only appeal.

2. Stage the package at a throwaway version. Do not bootstrap with a real
   release version, and do not bootstrap with an empty placeholder: npm's
   acceptable content policy forbids content that exists only to reserve a name,
   so the first publish carries the actual build.

   ```bash
   node scripts/alias-package.mjs --out .tmp/alias-package --version 0.0.1
   ```

   Read `.tmp/alias-package/package.json` before continuing. It should differ
   from the repository's own manifest on the `name` line, on the `version` line,
   and by the absence of the `scripts` block, and nowhere else.

3. Publish it, from your own npm account:

   ```bash
   npm login
   npm publish .tmp/alias-package --access public --ignore-scripts
   ```

   No `--provenance`: provenance comes from the CI environment, and this publish
   is not in one.

4. On npmjs.com, open the new package's settings and add a trusted publisher
   with the values listed under Release Prerequisites above.

5. Deprecate the bootstrap version, which exists for no other reason than to
   have created the package:

   ```bash
   npm deprecate <name>@0.0.1 "Bootstrap release. Install the current version."
   ```

6. Delete the staging directory and confirm the working tree is clean. The
   staging script never writes into the repository's own `package.json`, so
   there is nothing to revert, but check anyway.

From the next release onward the name is published by CI with provenance like
any other, and the rule that admits no exceptions is back in force: no local
`npm publish`.
