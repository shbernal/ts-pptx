---
doc-schema-version: 1
title: "Inspect a package"
summary: "Report what is on each slide of a .pptx, and where, with pptx-ts/inspect: slide and element fields, graphic frames, group geometry and z-order, and the package helpers."
read_when:
  - Inspecting generated or edited PPTX files
  - Building a linter, overlap check, or deck diff over slide contents
  - Checking object names, boxes, text, colors, or slide parts
  - Deciding between pptx-ts/inspect and pptx-ts/read
doc_type: "reference"
---

# Inspect a package

`pptx-ts/inspect` reports what is on each slide of a `.pptx`, and where.
It returns one flat list of elements per slide, each with a slide-absolute box in inches.

```ts
import { inspectPptx, overlapArea } from 'pptx-ts/inspect'

const { slides } = await inspectPptx('deck.pptx')
for (const slide of slides) {
  const leaves = slide.elements.filter((element) => element.kind !== 'group')
  leaves.forEach((a, i) => {
    for (const b of leaves.slice(i + 1)) {
      if (overlapArea(a.box, b.box) > 0) console.log(`${slide.path}: "${a.name}" overlaps "${b.name}"`)
    }
  })
}
```

[`inspectPptx`](api/inspect/functions/inspectPptx.md) takes a filesystem path (Node), the archive bytes (`Uint8Array`, `ArrayBuffer`, `Blob` or `number[]`), or a promise of either.

## Inspect or read?

`pptx-ts/inspect` is a flat projection over the read model, not a second reader.
Both load the package through the same `OpcPackage` and parser, so they cannot disagree about what a deck says.

| Use `pptx-ts/inspect` when | Use `pptx-ts/read` when |
| --- | --- |
| You want every element on every slide as one flat list | You want to walk the shape tree, or change the deck and save it |
| You compare boxes across groups in slide-absolute inches | You need table cells, chart series, SmartArt nodes, speaker notes, comments or animations |
| An element should report only the geometry its slide states | A placeholder's box should resolve through its layout and master |
| Per-run formatting of text boxes is enough | You need per-run formatting inside a table cell or SmartArt node |

[`loadPptxPackage`](api/inspect/functions/loadPptxPackage.md) returns the `OpcPackage` that [`Presentation.fromPackage()`](api/read/classes/Presentation.md) accepts, so a tool can start here and move to the read model without loading the bytes again.
[The read object model](read-object-model.md) describes that model.

## Slides

`inspectPptx` returns `slideSize` and `slides`.
`slideSize` is the deck's `p:sldSz` in inches, rounded to three decimals, or 10 × 7.5 in when the deck declares none.
`slides` follows presentation order (`p:sldIdLst`), the order PowerPoint shows.
Part names stop matching that order once a deck is reordered.

| Field | Type | Holds |
| --- | --- | --- |
| `index` | `number` | 0-based position in presentation order |
| `name` | `string` | `p:cSld/@name`, or `Slide N` (1-based) when unset |
| `path` | `string` | The slide part's zip path, without a leading `/` |
| `size` | `PptxSlideSize` | The deck's `slideSize`, repeated on every slide |
| `elements` | `PptxSlideElement[]` | The slide's elements in paint order |
| `text` | `string` | The non-empty `text` of every element, joined with single spaces |
| `wordCount` | `number` | The number of whitespace-separated words in `text` |

Hidden slides are reported, and neither a slide nor an element carries a hidden flag.
An OPC package with no readable presentation part inspects as zero slides rather than throwing.

## Element fields

Each element is a [`PptxSlideElement`](api/inspect/interfaces/PptxSlideElement.md).
A text frame is the `p:txBody` of a `p:sp`, and no other kind of element has one.
Explicit means set on the element itself, with nothing resolved from the layout, master or theme.

| Field | Type | Unit | Holds | `null` or empty when |
| --- | --- | --- | --- | --- |
| `id` | `string \| number` | none | `p:cNvPr/@id` | Never; `zIndex + 1` stands in when the attribute is missing |
| `name` | `string` | none | `p:cNvPr/@name` | Never; the kind and `zIndex + 1` stand in when unset, as in `shape 4` |
| `kind` | `PptxSlideElementKind` | none | `'group'` for `p:grpSp`, `'graphicFrame'` for `p:graphicFrame`, `'text'` for a `p:sp` with text, `'image'` for `p:pic`, `'shape'` for the rest (a `p:sp` without text, `p:cxnSp`) | Never |
| `graphicKind` | `PptxGraphicKind \| null` | none | What a graphic frame hosts; see [Graphic frames](#graphic-frames) | `null` unless `kind` is `'graphicFrame'` |
| `zIndex` | `number` | none | 0-based paint order within the slide | Never |
| `box` | `PptxBox` | inches | Slide-absolute `x`, `y`, `w`, `h` of the unrotated placement box | Never; an element with no resolvable box is omitted |
| `rotation` | `number` | degrees clockwise, `[0, 360)` | Effective rotation after group composition | Never; `0` when unrotated |
| `flipH`, `flipV` | `boolean` | none | Effective flips after group composition | Never |
| `parentZIndex` | `number \| null` | none | `zIndex` of the enclosing group | `null` at slide level |
| `childZIndices` | `number[]` | none | `zIndex` of each direct child, in document order | Empty unless `kind` is `'group'` |
| `text` | `string` | none | The runs joined, whitespace collapsed to one trimmed line; for a graphic frame, see [Graphic frames](#graphic-frames) | Empty with no text frame (pictures, connectors, groups) or no run text |
| `textWrap` | `string \| null` | none | Explicit `a:bodyPr/@wrap`, such as `square` or `none` | `null` with no text frame or no `a:bodyPr`, or when the attribute is unset |
| `autofit` | `PptxAutofitMode \| null` | none | Explicit `a:bodyPr` autofit child: `'none'`, `'normAutofit'` or `'spAutoFit'` | `null` with no text frame or no `a:bodyPr` |
| `autofitFontScale` | `number \| null` | percent | `a:normAutofit/@fontScale`, such as `62.5` | `null` unless the frame has `a:normAutofit` with a baked scale |
| `bodyInsets` | `PptxBodyInsets \| null` | inches | `left`, `top`, `right`, `bottom` from `a:bodyPr`, with PowerPoint's defaults for unset sides (0.1 left and right, 0.05 top and bottom) | `null` with no text frame or no `a:bodyPr` |
| `textRuns` | `PptxTextRun[]` | none | Every run of every paragraph, in document order | Empty with no text frame, and always for a graphic frame |
| `paragraphs` | `PptxParagraph[]` | none | The same runs grouped by `a:p` | Empty with no text frame, and always for a graphic frame |
| `fontSizes` | `number[]` | points | Distinct explicit run sizes (`a:rPr/@sz`), in first-seen order | Empty when no run sets a size |
| `colors` | `string[]` | hex, no `#` | Distinct explicit RGB run colors (`a:srgbClr`), in first-seen order | Empty when no run sets an RGB color; theme colors are not reported |
| `fill` | `string \| null` | hex, no `#` | Explicit RGB solid fill in `p:spPr`, or `p:grpSpPr` for a group | `null` for a theme color, gradient, picture fill or no fill, and for a graphic frame |
| `line` | `string \| null` | hex, no `#` | Explicit RGB outline color (`a:ln`) | `null` for a theme color or no outline, and for a group or graphic frame |
| `shapeType` | `string \| null` | none | Preset geometry (`a:prstGeom/@prst`), such as `rect` or `straightConnector1` | `null` for custom geometry, groups and graphic frames |

`text` joins runs with no separator, including across paragraphs.
The last word of one paragraph and the first word of the next therefore count as one word.

`textRuns[].text` is the run's `a:t` verbatim, including the whitespace an `xml:space="preserve"` run carries.
Each run is a [`PptxTextRun`](api/inspect/interfaces/PptxTextRun.md): size, RGB color, typeface, bold, italic, strike, highlight and character spacing, as the run sets them.
An unset value is `null`, except `bold` and `italic`, which are `false`.

`autofit` tells a bounded text box from one that grows:

- `'none'` keeps a fixed height, so its text can overflow.
- `'normAutofit'` (write-side `fit: 'shrink'`) shrinks the text to fit.
  A bare `<a:normAutofit/>` has `autofitFontScale: null`, and PowerPoint draws it at 100% until the text is edited.
- `'spAutoFit'` (write-side `fit: 'resize'`) resizes the shape to fit, so its authored height is an output.

Subtract `bodyInsets` from `box` to get the inner text box.

## Graphic frames

A `p:graphicFrame` is one element with `kind: 'graphicFrame'`, its own box and its own `zIndex`.
`graphicKind` comes from the frame's `a:graphicData/@uri`.
The structure inside the frame is not flattened; walk `pptx-ts/read` for it.

| `graphicKind` | Hosts | `text` | `textRuns` and `paragraphs` | Structure |
| --- | --- | --- | --- | --- |
| `'table'` | A table (`a:tbl`) | Cell text in row order, separated by spaces | Empty | Not reported: no rows, columns or cells |
| `'chart'` | A classic chart (`c:chart`) | Empty, because data labels and axis titles are chart data, not slide text | Empty | Not reported: no series or axes |
| `'chartEx'` | A 2016-family chart (`cx:chartSpace`), such as a waterfall, funnel or treemap | Empty | Empty | Not reported |
| `'diagram'` | A SmartArt graphic | Node text | Empty | Not reported: no nodes |
| `'other'` | A payload this library does not model, such as an OLE object or a 3D model | Empty | Empty | Not reported; the box marks where it sits |

Frame text counts toward the slide's `text` and `wordCount`.
That matches `Slide.text` on the read model, which also leaves chart text out.

## Geometry, groups, and z-order

`box` composes every enclosing group transform, so boxes compare directly whether or not an element is grouped.
A group places its children in a private coordinate space (`a:chOff` and `a:chExt`), and a child's raw `a:xfrm` never reaches this surface.
For a rotated element, `box` is the unrotated placement box, the box PowerPoint writes after Ungroup.
`rotation`, `flipH` and `flipV` carry the orientation.
These values are the read model's `absoluteFrame` in inches, and [the read object model](read-object-model.md#absolute-frame-and-groups) explains how it resolves.

An element whose box cannot be resolved is omitted.

| Cause | Warning code |
| --- | --- |
| The element has no transform of its own, such as a placeholder that inherits its box from the layout | None |
| An enclosing group has no usable `a:xfrm` | `inspect/group-transform-missing` |
| An enclosing group has a zero `a:chExt` | `inspect/group-transform-degenerate` |

[Errors and warnings](../errors-and-warnings.md#warnings) shows how to route the warnings.
An omitted element keeps its `zIndex`, so the numbers can have gaps.

`zIndex` is 0-based paint order, from a depth-first walk of the shape tree in document order.
A higher value draws on top, and a group comes immediately before its own children.

- `kind: 'group'` is a group container with its own id, name, box and fill, and no text.
- `parentZIndex` points from an element to its enclosing group.
- `childZIndices` points from a group to its direct children.

A group's box overlaps its children, so an overlap or coverage check usually wants leaves only.
Filter out `kind === 'group'`, or keep elements whose `childZIndices` is empty.

## Package and geometry helpers

| Member | What it does |
| --- | --- |
| [`loadPptxPackage`](api/inspect/functions/loadPptxPackage.md) | Loads the input as an `OpcPackage` |
| [`listPptxParts`](api/inspect/functions/listPptxParts.md) | Lists every part as a zip path, sorted |
| [`readPptxTextPart`](api/inspect/functions/readPptxTextPart.md) | Reads a part as UTF-8 text, or `null` when the part is absent |
| [`readPptxBinaryPart`](api/inspect/functions/readPptxBinaryPart.md) | Reads a part's current bytes, or `null` when the part is absent |
| [`readPresentationSize`](api/inspect/functions/readPresentationSize.md) | Reads `slideSize` from a loaded package |
| [`extractSlides`](api/inspect/functions/extractSlides.md) | Builds `slides` from a loaded package |
| [`boxAnchor`](api/inspect/functions/boxAnchor.md) | Returns a box's left, center or right x, or its top, middle or bottom y |
| [`overlapArea`](api/inspect/functions/overlapArea.md) | Returns the area two boxes share, or `0` when they do not overlap |
| [`DEFAULT_INSPECT_SLIDE_SIZE`](api/inspect/variables/DEFAULT_INSPECT_SLIDE_SIZE.md) | The 10 × 7.5 in size reported when a deck declares none |

The input must be a real OPC package.
A zip that holds slide XML but no `[Content_Types].xml` throws a `PackageReadError` with the code `package/not-an-opc-package`.

This surface reports boxes, text and overlap areas.
The tool that consumes them decides which margins, overlaps and colors are acceptable.
