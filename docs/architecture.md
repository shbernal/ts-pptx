---
doc-schema-version: 1
title: "Architecture"
summary: "How PptxGenJS is structured and where major responsibilities live."
read_when:
  - Changing module boundaries
  - Explaining architecture or ownership decisions
  - Reviewing whether a new feature belongs in the current structure
  - Finding which file and function implements a given task
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
- `src/pptxgen.ts` owns the main presentation class: presentation-level state,
  the authoring API façade (`addSlide`, `defineSlideMaster`, …), and the metadata
  and enum accessors. `write`/`writeFile`/`stream` are thin façades over the
  packaging layer.
- `src/package/assemble.ts` owns package assembly, split into two composable halves:
  `buildPackageParts` turns an authored deck into every OOXML part in emission order
  (`[Content_Types].xml`, the rels graph, docProps, theme, per-slide/layout/master
  parts, comments, chart/media rels), and `zipPackageParts` compresses that ordered
  list to the requested output shape. `writePackage` is their composition — the entry
  behind `write`/`writeFile`/`stream`. It takes a structural `PackageSource` the
  presentation class satisfies, so it does not depend on the class. The assembly half
  is also exposed publicly as `pptx.toParts()` (returning `PackagePart[]` = `{ path,
  data }`, dropping the internal-only STORE/DEFLATE hint): **part paths and their
  emission order are a stability-guaranteed observable contract** — adding a part later
  is back-compatible if existing paths/order do not shift; renaming/reordering is
  breaking. The per-part bytes are identical to what `write()` compresses.
- `src/slide.ts` owns slide-level object collection and public slide methods.
- `src/gen/` holds the internal OOXML generators as a layered tree mirroring
  `src/read/`: `gen/define/*` normalizes user options onto the slide model, and
  `gen/{drawingml,slide,pres,opc,chart,table,anim}/*` serialize that model to
  OOXML at export time. Chart emission is split per plot family under `gen/chart/`
  (`chart-parts` → `chart-axes` / `plot-*` → `chart-xml`) behind the `makeChartType`
  dispatch. `src/gen-utils.ts` holds only the cross-cutting helpers that belong to
  no single part (XML escaping, object names, rel ids).
- `src/core-interfaces.ts` and `src/core-enums.ts` define the public typed contract.
  `core-interfaces.ts` is a re-export barrel over `src/types/*` (split by domain).
  The generator-internal `*Internal` wire shapes live in `src/types/internal.ts`
  and are **not** re-exported — internal code imports them from there directly, the
  same non-published convention as `units-internal.ts`.
- `scripts/package-smoke.mjs` verifies the packed package boundary from a
  consumer perspective.

## Where Does X Live? (task → file → function)

A starting point for "which function do I touch?". Two-phase pattern for most content:
an `add*Definition` in `gen/define/*` normalizes user options onto the slide model,
then a serializer under `gen/{slide,drawingml,chart,anim,pres,opc}/*` emits OOXML at
export time. Each module opens with a TSDoc header stating its job; larger files add
`// ===== region =====` banners — grep those to jump within a file.

| Task | Add / normalize (`src/…`) | Emit OOXML (`src/…`) |
| --- | --- | --- |
| Add text | `gen/define/text.ts` `addTextDefinition` | `gen/drawingml/text-body.ts` `genXmlTextBody` |
| Add a shape | `gen/define/shape.ts` `addShapeDefinition` | `gen/drawingml/geometry.ts` `genXmlPresetGeom` / `genXmlCustGeom` |
| Add a connector | `gen/define/connector.ts` `addConnectorDefinition` | `gen/slide/object.ts` `slideObjectToXml` |
| Add an image | `gen/define/image.ts` `addImageDefinition` | `gen/slide/object.ts` `slideObjectToXml` |
| Add audio/video | `gen/define/media.ts` `addMediaDefinition` | `gen/slide/object.ts` `slideObjectToXml` + `gen/anim/timing.ts` `slideTimingToXml` |
| Add a chart | `gen/define/chart.ts` `addChartDefinition` | `gen/chart/chart-xml.ts` `makeXmlCharts` / `makeChartType` (+ `gen/chart/embed-xlsx.ts` `buildEmbeddedWorksheet`) |
| Add a table | `gen/define/table.ts` `addTableDefinition`; auto-paging `gen/table/autopage.ts` `getSlidesForTableRows` | `gen/slide/object.ts` `slideObjectToXml` (table branch) |
| Group objects | `gen/define/group.ts` `addGroupDefinition` / `groupObjectsDefinition` | `gen/slide/object.ts` `slideObjectToXml` |
| Notes | `gen/define/notes.ts` `addNotesDefinition` | `gen/slide/notes.ts` `makeXmlNotesSlide` |
| Comments | `gen/define/comment.ts` `addCommentDefinition` | `gen/slide/comments.ts` `makeXmlComments` |
| Slide number | `pptxgen.ts` `setSlideNumber` | `gen/slide/object.ts` `slideObjectToXml` (`SLDNUMFLDID`) |
| Transitions / animations | slide props (`slide.ts`) | `gen/anim/transition.ts` `slideTransitionToXml` / `gen/anim/animation.ts` `buildAnimationSeq` |
| Slide master / layout | `gen/define/master.ts` `createSlideMaster` | `gen/slide/master.ts` `makeXmlMaster` / `gen/slide/layout.ts` `makeXmlLayout` |
| Theme colors | — | `gen/pres/theme.ts` `buildThemeClrScheme` / `makeXmlTheme` |
| Coordinates & units (in → EMU) | `units.ts` (strict public primitives); `units-internal.ts` `getSmartParseNumber` (lenient generator layer) | — |
| Colors, fills, borders, shadows | — | `gen/drawingml/color.ts` `createColorElement`; `gen/drawingml/fill.ts` `genXmlColorSelection` / `genXml*Fill`; `gen/drawingml/line.ts` `genXmlLineFill` / `createLineCap`; `gen/drawingml/effect.ts` `createShadowElement` / `createGlowElement` |
| Package assembly & export | `package/assemble.ts` `buildPackageParts` (parts) + `zipPackageParts` (zip) → `writePackage` (behind `pptxgen.ts` `write` / `writeFile` / `stream`); `toParts` exposes the parts | `gen/opc/content-types.ts` `makeXmlContTypes` / `gen/opc/root-rels.ts` `makeXmlRootRels` / per-part rels |
| HTML `<table>` → slides (live DOM) | `browser.ts` `tableToSlides` (browser/standalone build only) | `gen/table/html-dom.ts` `genTableToSlides` |
| Public API surface | `pptxgen.ts` (class), `slide.ts` (slide methods) | — |
| Option / type definitions | `core-interfaces.ts` | — |
| Enums & shared constants | `core-enums.ts` | — |

## Boundaries

- The maintained runtime package is ESM-only.
- CommonJS and IIFE/global browser bundles are not maintained package targets.
- `dist/` is generated release output, not hand-edited source.
- Internal OOXML generators are implementation details unless deliberately
  exposed through `package.json` exports and public declarations.
- Platform differences go through the `RuntimeAdapter` seam (`src/runtime/*`): the
  `node`/`browser` entry subclasses inject the matching adapter into the shared core
  class. Live-DOM features that only work in a browser (currently `tableToSlides`)
  are defined on the browser entry subclass, not the core class, so they stay off
  the Node build and out of the shared chunk — their code bundles into the
  browser/standalone chunks alone.
- Downstream deck-production workflows belong in the consuming project unless the
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
