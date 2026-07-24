---
doc-schema-version: 1
title: "Reference"
summary: "Stable commands, configuration, APIs, and generated references for ts-pptx."
read_when:
  - Looking up command, config, or API details
  - Adding a reference page
  - Verifying docs against exported contracts
doc_type: "reference"
---

# Reference

Use this area for stable public contracts. Implementation details such as
internal XML generators stay out of the public API reference unless they are
deliberately exposed through the package boundary.

## Public API

The generated TypeDoc reference is built from public entry points:

- `src/index.ts`
- `src/core.ts`
- `src/node.ts`
- `src/browser.ts`
- `src/standalone.ts`

`src/inspect.ts`, `src/measure.ts`, `src/read.ts`, `src/math.ts`, and
`src/zip.ts` are separate public subpaths (see
[Runtime And Package Support](../runtime-and-package-support.md)) not covered
by the generated TypeDoc reference; they're documented by hand instead:
[PPTX Inspection](pptx-inspection.md), [PPTX Read / Round-Trip](pptx-read.md),
[Measured Text Fit](../measured-text-fit.md), and
[Math Equations](../math-latex.md). `src/zip.ts` has no dedicated page yet —
it is internal OPC/zip plumbing shared by `read` and `inspect`.

After running `pnpm run docs:api`, read the generated reference at
`reference/api/index.md`.

## Object Identity

[Object Identity](object-identity.md) documents the `objectName` contract for
stable PowerPoint Selection Pane names and the current `altText` serialization
contract for images and charts.

## Layout Units

[Layout Units](layout-units.md) documents standard slide-layout constants and
unit helpers for converting inches, points, pixels-at-DPI, and EMUs.

## PPTX Inspection

[PPTX Inspection](pptx-inspection.md) documents low-level package inspection,
slide/object extraction, and geometry helpers for downstream tools.

## Package Boundary

- Supported import paths are declared in `package.json` exports.
- Typed consumer contracts come from generated declarations.
- Internal `src/gen-*.ts` generation primitives are not consumer import paths.

## Commands

Stable verification commands are documented in [Testing Guide](../testing.md).
Development commands are documented in
[Development Guide](../development.md).
