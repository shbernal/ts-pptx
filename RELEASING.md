# PptxGenJS Release Checklist

This guide documents the maintained release path for the ESM-only package.

## Version Updates

1. Update `package.json` version.
2. Update `src/pptxgen.ts` version.
3. Update `CHANGELOG.md` with the release date and summary.
4. Update demo package versions when they intentionally track the release
   version.

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

## Beta Publish

```bash
pnpm publish --tag beta
```

## Stable Publish

```bash
pnpm publish
```

## GitHub Release

1. Merge the release branch into the release branch target.
2. Copy the changelog entry into a new GitHub release.
3. Use `vX.Y.Z` as the release tag.
4. Publish the release after npm publishing succeeds.
