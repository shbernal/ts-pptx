---
doc-schema-version: 1
title: "Release Workflow"
summary: "Maintained release path for publishing the scoped ESM-only package."
read_when:
  - Preparing a release
  - Updating npm publish or GitHub Release workflow documentation
  - Checking release verification steps
doc_type: "runbook"
---

# Release Workflow

This guide documents the maintained release path for the scoped ESM-only
package, `@shbernal/ts-pptx`.

Publishing is automated by `.github/workflows/publish.yml`. The workflow runs
when a GitHub Release is published and can also be dispatched manually from a
matching tag as a retry path.

## Release Prerequisites

- The npm package is `@shbernal/ts-pptx`.
- `package.json#repository.url` points at `shbernal/ts-pptx`.
- npm trusted publishing is configured for:
  - package: `@shbernal/ts-pptx`
  - GitHub repository: `shbernal/ts-pptx`
  - workflow filename: `publish.yml`
  - GitHub environment: `npm-publish`
  - allowed action: `npm publish`
- The GitHub Environment `npm-publish` exists before the first automated
  release.
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

Install dependencies and the OOXML validator:

```bash
pnpm install --frozen-lockfile
./tools/ooxml-validator/install.sh
```

Run the full automated gate before tagging:

```bash
pnpm run check:static
pnpm run verify:full
```

Together these are what CI runs. `check:static` adds `lint` and `format:check`,
which `verify:full` deliberately omits because the git hooks own them, but a
release is exactly the moment to confirm them explicitly.

Check that the target version is not already published:

```bash
npm view @shbernal/ts-pptx@X.Y.Z version
```

The command should fail with a registry 404 for a new release version.

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
- checks that `@shbernal/ts-pptx@X.Y.Z` is unpublished
- installs with `pnpm install --frozen-lockfile`
- installs the OOXML validator
- runs lint, formatting, typecheck, tests, and the package boundary checks
  (`package:lint` + `test:package`)
- publishes with `npm publish --access public --provenance --ignore-scripts`

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

## Post-Publish Checks

Verify npm and GitHub agree on the release:

```bash
npm view @shbernal/ts-pptx@X.Y.Z version dist-tags --json
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
