---
doc-schema-version: 1
title: "Architecture"
summary: "How ts-pptx is structured and where major responsibilities live."
read_when:
  - Changing module boundaries
  - Explaining architecture or ownership decisions
  - Reviewing whether a new feature belongs in the current structure
  - Finding which file and function implements a given task
doc_type: "architecture"
---

# Architecture

ts-pptx is a TypeScript library that turns a presentation object model into an
OOXML `.pptx` package. Consumer projects should import only the public package
exports and let this repository own the internal OOXML generation details.

## Responsibilities

- `src/index.ts`, `src/node.ts`, and `src/browser.ts` define the public entry
  points described by `package.json` exports.
- `src/pptxgen.ts` owns the main presentation class: presentation-level state,
  the authoring API façade (`addSlide`, `defineSlideMaster`, …), and presentation
  metadata. Enums are imported from the package entry, not read off the instance.
  `write`/`writeFile`/`stream` are thin façades over the packaging layer.
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
  dispatch, and shape emission is split per shape kind under `gen/slide/objects/`
  behind the `slideObjectToXml` dispatch — `gen/slide/object.ts` keeps only that
  walk, the group and slide-number branches that consume its shape-id counter, and
  the slide `.rels`. `src/gen-utils.ts` holds only the cross-cutting helpers that
  belong to no single part (XML escaping, object names, rel ids).
- `src/core-interfaces.ts` and `src/core-enums.ts` define the public typed contract.
  `core-interfaces.ts` is a re-export barrel over `src/types/*` (split by domain).
  The generator-internal `*Internal` wire shapes live in `src/types/internal.ts`
  and are **not** re-exported — internal code imports them from there directly, the
  same non-published convention as `units-internal.ts`.
- `src/script/` turns a deck read through `src/read/` into a serializable
  description of the write-API calls that would rebuild it (`readModelToIr`),
  and prints that description as a runnable TypeScript module. The two halves
  meet only at the IR: `from-read/` knows OOXML and the read model, `print/`
  knows only strings, and neither can see the other. That is what makes the
  mapping testable without a printer and keeps "how a number is spelled" from
  changing what a deck means. Two printers sit over the one IR and differ only
  in where the deck's *chrome* comes from. `printScript` anchors its output on a
  template — it reuses the *source deck itself*, because `fromTemplate` strips a
  package's slides while leaving masters, layouts, theme, and document properties
  byte-identical, so only slide content is ever regenerated.
  `printStandaloneScript` emits a module that depends on nothing but this
  package, re-authoring the theme and one `defineSlideMaster` per source layout
  from what the read model exposes. The split is not a preference: a theme's
  `a:fmtScheme`, a master's `p:txStyles`, and master/layout decoration are
  unreachable from *both* directions, so a rebuilt design can only ever be an
  approximation, while a reused one is exact. Each printer therefore returns the
  note set that applies to **its** output — suppressing what its tier rescues and
  adding what its tier costs — and that set, not `DeckIr.fidelity`, is what a
  round-trip check excludes.
  It is its own subsystem because it depends on **both** halves — the read model
  and the write option types — so it fits inside neither, and because `src/read/`
  is documented as isomorphic (bytes in, bytes out), which a converter emitting
  source text would quietly break for every `ts-pptx/read` consumer. Losses are
  data, not log lines: anything that cannot survive is a `FidelityNote` on the
  IR, which is what lets a round-trip check exclude exactly the declared losses
  and treat every other difference as a defect. `verify/` is that check:
  `canonicalDeckIr` reduces an IR to the form a comparison can use (dropping
  only values whose explicit and absent spellings are the same OOXML default),
  and `diffDeckIr` compares the source deck's IR against the IR of the deck a
  generated script produced, with the printer's notes as the exclusion list.
  It compares IRs rather than packages because the output can never be
  byte-identical — fresh rel ids, regenerated shape ids — so a byte comparison
  would fail for every deck and measure nothing. Its reach is bounded in two
  ways worth knowing before trusting a clean run: both IRs come from the same
  reader, so a construct the read path cannot see is absent from both, and the
  converter need not be injective, so two source constructs that map to the
  same call compare equal. It detects *asymmetry*; `pnpm run read:census` and
  the IR unit tests cover the rest. The consumer-facing guide — both tiers, the
  measured loss list, and how to read a fidelity note — is
  [PPTX To Script](reference/pptx-to-script.md).
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
| Add a connector | `gen/define/connector.ts` `addConnectorDefinition` | `gen/slide/objects/connector.ts` `renderConnectorObject` |
| Add an image | `gen/define/image.ts` `addImageDefinition` | `gen/slide/objects/image.ts` `renderImageObject` |
| Add audio/video | `gen/define/media.ts` `addMediaDefinition` | `gen/slide/objects/media.ts` `renderMediaObject` + `gen/anim/timing.ts` `slideTimingToXml` |
| Add a chart | `gen/define/chart.ts` `addChartDefinition` | `gen/chart/chart-xml.ts` `makeXmlCharts` / `makeChartType` (+ `gen/chart/embed-xlsx.ts` `buildEmbeddedWorksheet`) |
| Add a table | `gen/define/table.ts` `addTableDefinition`; auto-paging `gen/table/autopage.ts` `getSlidesForTableRows` | `gen/slide/objects/table.ts` `renderTableObject` |
| Group objects | `gen/define/group.ts` `addGroupDefinition` / `groupObjectsDefinition` | `gen/slide/object.ts` `slideObjectToXml` (group branch) |
| Notes | `gen/define/notes.ts` `addNotesDefinition` | `gen/slide/notes.ts` `makeXmlNotesSlide` |
| Comments | `gen/define/comment.ts` `addCommentDefinition` | `gen/slide/comments.ts` `makeXmlComments` |
| Slide number | `pptxgen.ts` `setSlideNumber` | `gen/slide/object.ts` `slideObjectToXml` (`SLDNUMFLDID`) |
| Transitions / animations | slide props (`slide.ts`) | `gen/anim/transition.ts` `slideTransitionToXml` / `gen/anim/animation.ts` `buildAnimationSeq` |
| Slide master / layout | `gen/define/master.ts` `createSlideMaster` | `gen/slide/master.ts` `makeXmlMaster` / `gen/slide/layout.ts` `makeXmlLayout` |
| Theme colors | — | `gen/pres/theme.ts` `buildThemeClrScheme` / `makeXmlTheme` |
| Coordinates & units (in → EMU) | `units.ts` (strict public primitives); `units-internal.ts` `getSmartParseNumber` (lenient generator layer) | — |
| Colors, fills, borders, shadows | — | `gen/drawingml/color.ts` `createColorElement`; `gen/drawingml/fill.ts` `genXmlColorSelection` / `genXml*Fill`; `gen/drawingml/line.ts` `genXmlLineFill` / `createLineCap`; `gen/drawingml/effect.ts` `createShadowElement` / `createGlowElement` |
| Package assembly & export | `package/assemble.ts` `buildPackageParts` (parts) + `zipPackageParts` (zip) → `writePackage` (behind `pptxgen.ts` `write` / `writeFile` / `stream`); `toParts` exposes the parts | `gen/opc/content-types.ts` `makeXmlContTypes` / `gen/opc/root-rels.ts` `makeXmlRootRels` / per-part rels |
| HTML `<table>` → slides | `html.ts` `tableToSlides` (the `ts-pptx/html` subpath, any DOM); `browser.ts` `tableToSlides` (method form, delegates) | `gen/table/html-dom.ts` `genTableToSlides` |
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
  browser chunk alone.
- Downstream deck-production workflows belong in the consuming project unless the
  behavior is broadly reusable for ts-pptx consumers.

## Data And Control Flow

1. Consumers create a presentation through a public ts-pptx entry point.
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
