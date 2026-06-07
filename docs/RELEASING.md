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
package, `@shbernal/pptxgenjs`.

Publishing is automated by `.github/workflows/publish.yml`. The workflow runs
when a GitHub Release is published and can also be dispatched manually from a
matching tag as a retry path.

## Release Prerequisites

- The npm package is `@shbernal/pptxgenjs`.
- `package.json#repository.url` points at `shbernal/PptxGenJS`.
- npm trusted publishing is configured for:
  - package: `@shbernal/pptxgenjs`
  - GitHub repository: `shbernal/PptxGenJS`
  - workflow filename: `publish.yml`
  - GitHub environment: `npm-publish`
  - allowed action: `npm publish`
- The GitHub Environment `npm-publish` exists before the first automated
  release.
- Do not add an `NPM_TOKEN` secret for the normal path. The workflow uses OIDC
  with `id-token: write`.

## Version Updates

1. Update `package.json` version.
2. Update `src/pptxgen.ts` version.
3. Update `CHANGELOG.md` with the release date and summary.
4. Update demo package versions when they intentionally track the release
   version.
5. Keep package import examples on the scoped package name:
   `@shbernal/pptxgenjs`.

## Local Release Gate

Install dependencies and the OOXML validator:

```bash
pnpm install --frozen-lockfile
./tools/ooxml-validator/install.sh
```

Run the full automated gate before tagging:

```bash
pnpm run lint
pnpm run format:check
pnpm run typecheck
pnpm test
pnpm run build
pnpm run package:lint
pnpm run pack:check
pnpm run test:package
pnpm run test:demos
npm pack --dry-run --ignore-scripts
```

Check that the target version is not already published:

```bash
npm view @shbernal/pptxgenjs@X.Y.Z version
```

The command should fail with a registry 404 for a new release version.

## Automated npm Publish

1. Merge the release commit into `mainline`.
2. Create a tag named exactly `vX.Y.Z`, matching `package.json#version`.
3. Push `mainline` and the tag.
4. Create a GitHub Release from `vX.Y.Z`.
5. Publish the GitHub Release.

Publishing the GitHub Release starts `.github/workflows/publish.yml`. The
workflow:

- refuses to run outside `shbernal/PptxGenJS`
- refuses branch publishes; `GITHUB_REF_TYPE` must be `tag`
- requires the tag name to equal `v${package.json#version}`
- checks that `@shbernal/pptxgenjs@X.Y.Z` is unpublished
- installs with `pnpm install --frozen-lockfile`
- installs the OOXML validator
- runs lint, formatting, typecheck, tests, package checks, package smoke tests,
  demo smoke tests, and `npm pack --dry-run --ignore-scripts`
- publishes with `npm publish --access public --provenance --ignore-scripts`

npm trusted publishing automatically exchanges the GitHub Actions OIDC token for
publish credentials. The explicit `--provenance` flag keeps provenance required
even if npm defaults change.

## Manual Workflow Retry

Use this only after fixing a failed publish workflow without changing the
release artifact:

```bash
gh workflow run publish.yml --repo shbernal/PptxGenJS --ref vX.Y.Z
```

The selected ref must be the release tag, not `mainline`.

## Post-Publish Checks

Verify npm and GitHub agree on the release:

```bash
npm view @shbernal/pptxgenjs@X.Y.Z version dist-tags --json
gh release view vX.Y.Z --repo shbernal/PptxGenJS
```

## Package Surface Checks

The package should ship:

- `dist/index.js`
- `dist/index.d.ts`
- `dist/core.js`
- `dist/node.js`
- `dist/browser.js`
- `dist/standalone.js`
- package `exports["."].default`
- package `exports["."].types`
- package subpaths for `./core`, `./node`, `./browser`, and `./standalone`
- scoped imports for `@shbernal/pptxgenjs`,
  `@shbernal/pptxgenjs/core`, `@shbernal/pptxgenjs/node`,
  `@shbernal/pptxgenjs/browser`, and `@shbernal/pptxgenjs/standalone`

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
