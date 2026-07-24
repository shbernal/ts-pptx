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

The `@shbernal/ts-pptx/inspect` subpath exposes low-level primitives for
tools that need to examine a PPTX package after generation or manual editing.
It is intentionally separate from the presentation-authoring API.

```ts
import { inspectPptx, loadPptxPackage, listPptxParts } from "@shbernal/ts-pptx/inspect"
```

`inspectPptx(input)` loads a PPTX package and returns:

- `slideSize`: presentation width and height in inches.
- `slides[]`: generated slide entries in package order.
- `slides[].elements[]`: normalized objects with `id`, `name`, `kind`,
  `zIndex`, `box`, `rotation`, `flipH`, `flipV`, `parentZIndex`, `childZIndices`,
  `text`, `textRuns`, `paragraphs`, `fontSizes`, `colors`, `fill`, `line`,
  `shapeType`, `textWrap`, `autofit`, `autofitFontScale`, and `bodyInsets`.

## Geometry, groups, and z-order

`box` is **slide-absolute** inches, composing every enclosing group transform, so
boxes are directly comparable whether or not an element is grouped. A group
(`<p:grpSp>`) authors its children in a private coordinate space (`a:chOff`/`a:chExt`)
that need not match its slide frame — PowerPoint makes it non-identity as soon as a
user resizes a group — so a child's raw `a:xfrm` is not placeable on the slide and is
never what you get here.

For a rotated element, `box` is the *unrotated* placement box (what PowerPoint writes
after Ungroup) and `rotation` (degrees, `[0, 360)`) / `flipH` / `flipV` report its
effective orientation after group composition. This mirrors `Shape.absoluteFrame` on
the read API, which shares the same implementation.

An element whose position cannot be resolved — an enclosing group with a degenerate
(zero) `a:chExt` — is omitted with a warning rather than reported at a wrong position.

`zIndex` is `0`-based paint order: a depth-first walk of the shape tree in document
order, so higher draws on top. Elements are linked by it:

- `kind: 'group'` is a group container. It has an id, name, box, and fill, but no text
  of its own.
- `parentZIndex` is the enclosing group's `zIndex`, or `null` at slide level.
- `childZIndices` lists a group's direct children in document order; empty for
  every other kind.

A group's box overlaps its children by construction, so tools that reason about layout
(overlap, coverage) usually want the leaves only — filter out `kind === 'group'`, or
keep elements with `childZIndices.length === 0`.

`autofit` and `bodyInsets` describe the text frame's `a:bodyPr` so a consumer can
tell a bounded text box from an auto-growing one and compute its inner box:

- `autofit`: `'none'` (fixed height — a genuine overflow candidate), `'normAutofit'`
  (shrink text to fit, ts-pptx `fit: 'shrink'`), or `'spAutoFit'` (resize shape to
  fit text, `fit: 'resize'` — the authored height is an output, so it cannot
  overflow). `null` for elements without a text frame (e.g. images).
- `bodyInsets`: `{ left, top, right, bottom }` in inches, with PowerPoint defaults
  applied when absent (0.1in left/right, 0.05in top/bottom). Subtract from `box` to
  get the inner text box. `null` for elements without a text frame.

The subpath also exports package helpers such as `loadPptxPackage()`,
`listPptxParts()`, and `readPptxTextPart()`, plus geometry helpers such as
`boxAnchor()` and `overlapArea()`.

Downstream tools should keep policy decisions outside this package. For
example, ts-pptx can report object boxes and overlap area, while a deck
production tool decides which margins, overlaps, colors, or semantic
relationships are acceptable for its workflow.
