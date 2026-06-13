---
doc-schema-version: 1
title: "Testing Guide"
summary: "Regression, schema, package, demo, and manual verification commands."
read_when:
  - Choosing verification commands
  - Updating test scripts or package smoke checks
  - Changing emitted OOXML or package exports
doc_type: "guide"
---

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

## Regression Suite Layout

Regression tests live in `test/regression/` and are organized by behavior, not
by historical bug number. File names should describe the contract being tested,
such as `object-identity.test.js`, `content-type-defaults.test.js`, or
`slide-master-placeholders.test.js`.

Each regression file calls `defineRegressionSuite()` from `test/helpers.js`.
The optional second argument records legacy provenance, for example
`legacy bug-21`, so old issue references remain traceable without making the
suite name opaque.

Prefer public API deck generation plus focused package/XML assertions:

- Use `build()` to create a presentation and inspect the generated package.
- Use `readEntry()` for specific package parts such as `ppt/slides/slide1.xml`.
- Use helper assertions such as `assertContentTypeDefault()`,
  `assertContentTypeOverride()`, `assertXmlOrder()`, and
  `assertNonVisualDrawingProperty()` when they match the behavior under test.
- Keep raw XML substring or regex assertions local and narrowly targeted when a
  helper would hide the OOXML detail being tested.

Add a regression test when a public API call must keep producing a specific
package part, relationship, OOXML element, attribute, or absence of generated
parts. Name the file after the behavior, and include bug or upstream issue
context in the suite metadata or test name only when it helps future triage.

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

## Read/Round-Trip Suite (`pptxgenjs/read`)

The lossless read/edit subsystem (`src/read/`) has its own harness:

```bash
pnpm run test:read
```

It runs `test/read/roundtrip.test.js` against real, PowerPoint-authored decks
in `test/read/fixtures/` (provenance in that directory's README): part-set
stability, per-part byte-identity for untouched parts, lazy-parse guarantees,
save idempotence, content-type/relationship resolution, the dirty
(mutate-and-reserialize) path, and schema validation of saved output. The
schema cases require the OOXML validator above and are skipped with a notice
when it is not installed.

Changes under `src/read/` should run this suite; new read/edit capabilities
should extend it (and grow the fixture set) alongside the code.

## Full Test Command

The default test command builds first, then runs regression and schema
validation:

```bash
pnpm test
```

## Package Boundary Checks

Build package artifacts:

```bash
pnpm run build
```

Check the packed package:

```bash
pnpm run package:lint
pnpm run pack:check
pnpm run test:package
```

`package:lint` runs package export/type validation. `pack:check` uses
`pnpm pack --dry-run`. `test:package` creates a packed package with pnpm,
installs it with npm and pnpm, verifies that the ESM entries and declarations
are present, verifies that old generated artifacts are absent, runs an ESM
import smoke test, checks that the package has no CJS export condition, and
typechecks a minimal TypeScript consumer.

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

Node demo decks are written to `demos/node/output/`, which is ignored by git.
Re-running a demo command replaces the previous deck with the same name.
