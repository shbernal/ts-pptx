---
doc-schema-version: 1
title: "PPTX Inspection"
summary: "Low-level package inspection and geometry helpers for generated or edited PPTX files."
read_when:
  - Inspecting generated PPTX files
  - Building downstream linting or review tools
  - Checking object names, boxes, text, colors, or slide parts
doc_type: "reference"
---

# PPTX Inspection

The `@shbernal/pptxgenjs/inspect` subpath exposes low-level primitives for
tools that need to examine a PPTX package after generation or manual editing.
It is intentionally separate from the presentation-authoring API.

```ts
import { inspectPptx, loadPptxPackage, listPptxParts } from "@shbernal/pptxgenjs/inspect"
```

`inspectPptx(input)` loads a PPTX package and returns:

- `slideSize`: presentation width and height in inches.
- `slides[]`: generated slide entries in package order.
- `slides[].elements[]`: normalized objects with `id`, `name`, `kind`,
  `zIndex`, `box`, `text`, `textRuns`, `fontSizes`, `colors`, `fill`, `line`,
  `shapeType`, and `textWrap`.

The subpath also exports package helpers such as `loadPptxPackage()`,
`listPptxParts()`, and `readPptxTextPart()`, plus geometry helpers such as
`boxAnchor()` and `overlapArea()`.

Downstream tools should keep policy decisions outside this package. For
example, PptxGenJS can report object boxes and overlap area, while a deck
production tool decides which margins, overlaps, colors, or semantic
relationships are acceptable for its workflow.
