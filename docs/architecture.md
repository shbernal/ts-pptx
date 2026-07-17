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
- `src/pptxgen.ts` owns the main presentation class and package export flow.
- `src/slide.ts` owns slide-level object collection and public slide methods.
- `src/gen-*.ts` files own internal generation primitives for XML, charts,
  objects, media, and tables.
- `src/core-interfaces.ts` and `src/core-enums.ts` define the public typed
  contract.
- `scripts/package-smoke.mjs` verifies the packed package boundary from a
  consumer perspective.

## Where Does X Live? (task → file → function)

A starting point for "which function do I touch?". Two-phase pattern for most content:
an `add*Definition` in `gen-objects.ts` normalizes user options onto the slide model,
then `gen-xml.ts` (or `gen-charts.ts`) serializes it to OOXML at export time. Each file
opens with a module-map header and `// ===== region =====` banners — grep those to
jump within a file.

| Task | Add / normalize (`src/…`) | Emit OOXML (`src/…`) |
| --- | --- | --- |
| Add text | `gen-objects.ts` `addTextDefinition` | `gen-xml.ts` `genXmlTextBody` |
| Add a shape | `gen-objects.ts` `addShapeDefinition` | `gen-xml.ts` `genXmlPresetGeom` / `genXmlCustGeom` |
| Add a connector | `gen-objects.ts` `addConnectorDefinition` | `gen-xml.ts` `slideObjectToXml` |
| Add an image | `gen-objects.ts` `addImageDefinition` | `gen-xml.ts` `slideObjectToXml` |
| Add audio/video | `gen-objects.ts` `addMediaDefinition` | `gen-xml.ts` `slideObjectToXml` + `slideTimingToXml` |
| Add a chart | `gen-objects.ts` `addChartDefinition` | `gen-charts.ts` `makeXmlCharts` / `makeChartType` (+ `buildEmbeddedWorksheet`) |
| Add a table | `gen-objects.ts` `addTableDefinition`; auto-paging `gen-tables.ts` `getSlidesForTableRows` | `gen-xml.ts` `slideObjectToXml` (table branch) |
| Group objects | `gen-objects.ts` `addGroupDefinition` / `groupObjectsDefinition` | `gen-xml.ts` `slideObjectToXml` |
| Notes | `gen-objects.ts` `addNotesDefinition` | `gen-xml.ts` `makeXmlNotesSlide` |
| Comments | `gen-objects.ts` `addCommentDefinition` | `gen-xml.ts` `makeXmlComments` |
| Slide number | `pptxgen.ts` `setSlideNumber` | `gen-xml.ts` `slideObjectToXml` (`SLDNUMFLDID`) |
| Transitions / animations | slide props (`slide.ts`) | `gen-xml.ts` `slideTransitionToXml` / `buildAnimationSeq` |
| Slide master / layout | `gen-objects.ts` `createSlideMaster` | `gen-xml.ts` `makeXmlMaster` / `makeXmlLayout` |
| Theme colors | — | `gen-xml.ts` `buildThemeClrScheme` / `makeXmlTheme` |
| Coordinates & units (in → EMU) | `gen-utils.ts` `getSmartParseNumber`; `units.ts` | — |
| Colors, fills, borders, shadows | `gen-utils.ts` `createColorElement` / `genXml*Fill` / `createShadowElement` | — |
| Package assembly & export | `pptxgen.ts` `exportPresentation` (`write` / `writeFile` / `stream`) | `gen-xml.ts` `makeXmlContTypes` / `makeXmlRootRels` / per-part rels |
| Public API surface | `pptxgen.ts` (class), `slide.ts` (slide methods) | — |
| Option / type definitions | `core-interfaces.ts` | — |
| Enums & shared constants | `core-enums.ts` | — |

## Boundaries

- The maintained runtime package is ESM-only.
- CommonJS and IIFE/global browser bundles are not maintained package targets.
- `dist/` is generated release output, not hand-edited source.
- Internal OOXML generators are implementation details unless deliberately
  exposed through `package.json` exports and public declarations.
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
