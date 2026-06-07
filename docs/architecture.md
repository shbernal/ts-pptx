---
doc-schema-version: 1
title: "Architecture"
summary: "How PptxGenJS is structured and where major responsibilities live."
read_when:
  - Changing module boundaries
  - Explaining architecture or ownership decisions
  - Reviewing whether a new feature belongs in the current structure
doc_type: "architecture"
---

# Architecture

PptxGenJS is a TypeScript library that turns a presentation object model into an
OOXML `.pptx` package. Consumer projects should import only the public package
exports and let this repository own the internal OOXML generation details.

## Responsibilities

- `src/index.ts`, `src/node.ts`, `src/browser.ts`, `src/standalone.ts`, and
  `src/core.ts` define the public entry points described by `package.json`
  exports.
- `src/pptxgen.ts` owns the main presentation class and package export flow.
- `src/slide.ts` owns slide-level object collection and public slide methods.
- `src/gen-*.ts` files own internal generation primitives for XML, charts,
  objects, media, and tables.
- `src/core-interfaces.ts` and `src/core-enums.ts` define the public typed
  contract.
- `scripts/package-smoke.mjs` verifies the packed package boundary from a
  consumer perspective.

## Boundaries

- The maintained runtime package is ESM-only.
- CommonJS and IIFE/global browser bundles are not maintained package targets.
- `dist/` is generated release output, not hand-edited source.
- Internal OOXML generators are implementation details unless deliberately
  exposed through `package.json` exports and public declarations.
- Downstream deck-production workflows belong in `downstream` unless the
  behavior is broadly reusable for PptxGenJS consumers.

## Data And Control Flow

1. Consumers create a presentation through a public PptxGenJS entry point.
2. Public methods collect slides and slide objects into internal structures.
3. The export flow calls internal generators to create package parts and OOXML.
4. Runtime adapters write the result for Node or browser environments.
5. Package smoke tests verify that consumers can import only supported public
   entry points.

## Extension Points

- Add public API only through exported entry points and generated declarations.
- Add OOXML behavior with focused regression or schema fixtures.
- Add package-boundary checks when changing public exports, runtime targets, or
  declaration output.
