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

1. Update `package.json` version.
2. Update `src/presentation.ts` version.
3. Update `CHANGELOG.md` with the release date and summary.
4. Update demo package versions when they intentionally track the release
   version.
5. Keep package import examples on the scoped package name:
   `@shbernal/ts-pptx`.

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
which `verify:full` deliberately omits because the git hooks own them — but a
release is exactly the moment to confirm them explicitly.

Check that the target version is not already published:

```bash
npm view @shbernal/ts-pptx@X.Y.Z version
```

The command should fail with a registry 404 for a new release version.

## Automated npm Publish

1. Merge the release commit into `master`.
2. Create a tag named exactly `vX.Y.Z`, matching `package.json#version`.
3. Push `master` and the tag.
4. Create a GitHub Release from `vX.Y.Z`.
5. Publish the GitHub Release.

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
`scripts/package-smoke.mjs`, so a new subpath is covered by being added there —
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
