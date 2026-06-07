# Release Checklist

This guide documents the maintained release path for the scoped ESM-only
package, `@shbernal/pptxgenjs`.

Publishing is manual until the first scoped release exists and the npm package
ownership/trusted-publishing setup is confirmed. Do not add or run automated npm
publishing without an explicit release-automation task.

## Version Updates

1. Update `package.json` version.
2. Update `src/pptxgen.ts` version.
3. Update `CHANGELOG.md` with the release date and summary.
4. Update demo package versions when they intentionally track the release
   version.
5. Keep package import examples on the scoped package name:
   `@shbernal/pptxgenjs`.

## Automated Release Gate

Install dependencies and the OOXML validator:

```bash
pnpm install --frozen-lockfile
./tools/ooxml-validator/install.sh
```

Run the full automated gate:

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

## Demo Checks

The maintained demos are:

- `demos/node`
- `demos/vite-demo`

Run both with:

```bash
pnpm run test:demos
```

## Manual Pack Check

```bash
mkdir -p /tmp/pptxgenjs-release
pnpm pack --pack-destination /tmp/pptxgenjs-release
```

Inspect the generated tarball before publishing:

```bash
tar -tf /tmp/pptxgenjs-release/shbernal-pptxgenjs-*.tgz
```

## Manual npm Publish

Only publish after confirming npm ownership and the exact target version:

```bash
npm publish /tmp/pptxgenjs-release/shbernal-pptxgenjs-*.tgz --access public
```

For a prerelease, publish with an explicit tag:

```bash
npm publish /tmp/pptxgenjs-release/shbernal-pptxgenjs-*.tgz --access public --tag beta
```

## GitHub Release

1. Merge the release branch into the release branch target.
2. Copy the changelog entry into a new GitHub release.
3. Use `vX.Y.Z` as the release tag.
4. State the npm package status in the release notes. For the first scoped
   GitHub release, npm publishing may remain manual or pending.
