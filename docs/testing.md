# Testing Guide

Use `pnpm` for repository scripts. The package declares Node.js `>=24`.

## Standard Validation

For source changes, run:

```bash
pnpm run build
pnpm run typecheck
pnpm run test:unit
```

For documentation-only changes, no automated test is required unless the docs
change package, build, or testing claims.

## OOXML Schema Validation

Install the validator once:

```bash
./tools/ooxml-validator/install.sh
```

Run schema fixtures:

```bash
pnpm run test:schema
```

Use this path for emitted OOXML changes. Add or update focused fixtures in
`test/schema.test.js`.

## Full Test Command

The default test command builds first, then runs regression and schema
validation:

```bash
pnpm test
```

## Package Boundary Checks

Build package artifacts:

```bash
pnpm run build:dist
pnpm run types:build
```

Check the packed package:

```bash
pnpm run pack:check
pnpm run test:package
```

`pack:check` uses `pnpm pack --dry-run`. `test:package` creates a packed
package with pnpm, installs it with npm and pnpm, verifies that the ESM entry
and declarations are present, verifies that old CJS and IIFE artifacts are
absent, runs an ESM import smoke test, checks that CommonJS `require` is
unsupported, and typechecks a minimal TypeScript consumer.

## Demo Smoke Tests

Run both maintained demo smoke tests:

```bash
pnpm run test:demos
```

Run one target:

```bash
pnpm run test:demo:node
pnpm run test:demo:vite
```

The demo smoke command builds package artifacts first, then runs the maintained
workspace demos with pnpm. The Node demo validates ESM package usage in a Node
application. The Vite demo validates a modern browser app path through React,
TypeScript, and Vite.

## Manual Visual Checks

Automated tests prove package shape and generated XML structure. Manual visual
checks are still useful for user-visible PowerPoint behavior:

1. Generate a small deck from the Node demo or Vite demo.
2. Open it in Microsoft PowerPoint when available.
3. Check import behavior in Keynote, LibreOffice Impress, or Google Slides when
   the change affects cross-app compatibility.
4. For browser download behavior, prefer `demos/vite-demo`.
