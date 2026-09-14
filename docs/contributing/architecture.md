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

ts-pptx turns a presentation object model into an OOXML `.pptx` package.
Consumers import the public package exports and nothing else. Everything about
how the OOXML gets built stays on this side of that line.

## Responsibilities

- `src/index.ts`, `src/node.ts`, and `src/browser.ts` define the public entry
  points described by `package.json` exports. Each is a `PresentationCore`
  subclass differing only in the `RuntimeAdapter` it injects, and `.` resolves to
  one of the three by condition: `node`, `browser`, or neither. `index.ts` is
  that third case (Deno, Bun, edge workers): it authors and exports bytes like the
  others, and refuses only what needs a host it does not have.
- `src/presentation.ts` owns the main presentation class: presentation-level state,
  the authoring API façade (`addSlide`, `defineSlideMaster`, …), and presentation
  metadata. Enums are imported from the package entry, not read off the instance.
  `write`/`writeFile`/`stream` are thin façades over the packaging layer.
- `src/package/assemble.ts` owns package assembly, in two composable halves.
  `buildPackageParts` turns an authored deck into every OOXML part in emission
  order: `[Content_Types].xml`, the rels graph, docProps, theme, the per-slide,
  layout and master parts, media rels, and whatever the part contributors add.
  `zipPackageParts` compresses that ordered list into the requested output shape.
  `writePackage` composes the two, and is the entry behind
  `write`/`writeFile`/`stream`. It takes a structural `PackageSource` that the
  presentation class happens to satisfy, so it never depends on the class.

  The assembly half is public too, as `pptx.toParts()`. It hands back
  `PackagePart[]`, which is `{ path, data }`, minus the internal STORE/DEFLATE
  hint. **Part paths and their emission order are a stability-guaranteed
  observable contract.** Adding a part later stays back-compatible as long as the
  existing paths and order hold. Renaming or reordering breaks. Per-part bytes
  match what `write()` compresses, exactly.
- `src/slide.ts` owns slide-level state: the object list, the rel arrays, the slide number,
  the background, and the geometry accessors over them. The `add*` methods are not written
  there; they are bound on from the construct families the presentation was composed with.
- `src/gen/` holds the internal OOXML generators, as a layered tree mirroring
  `src/read/`. `gen/define/*` normalizes user options onto the slide model.
  `gen/{drawingml,slide,pres,opc,chart,table,anim}/*` serialize that model to
  OOXML at export time. Chart emission splits per plot family under `gen/chart/`
  (`chart-parts` → `chart-axes` / `plot-*` → `chart-xml`) behind the
  `makeChartType` dispatch. A chart's embedded workbook (`embed-xlsx`) and every formula
  in its chart part are laid out from one `worksheetLayout` in `gen/chart/data-refs.ts`,
  built once per chart, and
  `test/regression/chart/chart-worksheet-invariant.test.js` resolves each formula through
  the workbook to hold the two sides together. Shape emission splits per shape kind under
  `gen/slide/objects/` behind the `slideObjectToXml` dispatch, so
  `gen/slide/object.ts` keeps only that walk, the group and slide-number branches
  that consume its shape-id counter, and the slide `.rels`. `src/gen/utils.ts`
  holds the cross-cutting helpers that belong to no single part: XML escaping,
  object names, rel ids.
- **A consumer composes a presentation from the families it needs.**
  `createPresentation({ use })` sits beside `TsPptx` on every entry.
  `entry-compose.ts` is the shared body; `families.ts` is the `pptx-ts/families`
  subpath a caller takes the values from. It composes the core tier plus what it
  is given, and both forms write the same deck for the same slides, part for
  part. This is a reachability difference, not a second write path.

  There is no second lean entry point, on purpose. One would need the same
  three-condition treatment `.` gets, for nothing export-level shaking does not
  already give. The type follows the composition: `addSlide()` answers a slide
  carrying exactly the composed methods, and each of those answers that same
  slide, so a chain cannot widen back to the full surface.

  What that costs and saves is not an argument but a gate. `bundle-tier:check`
  bundles a composed program, a composed-plus-one-family program, and the
  `TsPptx` ones. The difference between two of those rows is what a family costs
  a consumer. The measured numbers, the core tier, and what a consumer sees when
  a family is missing, are in [Smaller bundles](../bundle-size.md).
- **A construct family is one value, and a presentation is composed with a list of
  them.** What the library knows about charts, or tables, or speaker notes is a
  `ConstructFamily` (`families/shared.ts`). That value carries the methods it
  adds to a slide (`addChart`), the child descriptors it recognises inside a
  slide master or a group (`{ chart: ... }`), the renderer that emits its XML,
  the part contributor that puts its parts in the package, and for two families a
  method on the presentation itself (`measureText`, `tableToSlides`).
  `entry-families.ts` is the full list, and all three entries compose it. The
  browser entry adds the table family's live-DOM half (`families/table-dom.ts`),
  the only piece that needs a `document`.

  This is the bundling constraint the two seams below record, one level up. A
  class method body is never shaken, so `SlideBuilder.addChart` calling
  `addChartDefinition` linked every plot module into every program that can make
  a slide. The methods are bound onto each slide instance instead, not registered
  on the prototype or into a module map. Two presentations composed differently
  in one process each get their own, and a method whose family was not composed
  raises `family/not-composed` naming the family, rather than simply being
  absent.

  What stays core is what no tier can drop: the slide background,
  `createSlideMaster` (handed the child-descriptor table rather than naming
  families itself), the shape-id allocator and `resolveObjectName` every family
  calls, and the export-time autofit bake. Slide *extraction*
  (`gen/extract-slides.ts`, the path `appendSlides` reads) is handed a family
  table too, for the same reason. It re-emits what a slide holds, so naming the
  chart emitters there put all of `gen/chart/` in the graph of every program that
  can build a deck.
- **The shape walk is handed its renderers; it does not import them.** Which renderer
  emits each shape family is a `RendererTable` (`gen/slide/objects/shared.ts`).
  It travels down the write path with the deck state: from
  `PackageSource.renderers`, through `makeXmlSlide` / `makeXmlLayout` /
  `makeXmlMaster`, into the walk. The presentation assembles that table from the
  construct families it was composed with, each of which names its own renderer.
  It is keyed by object kind rather than written as a `switch`, and every
  renderer takes one `RenderContext` and nothing else. So the table has a single
  signature, and `gen/slide/object.ts` stays the only module that knows which
  object kinds exist.

  This is a bundling constraint written into the code. A named import inside a
  reachable function body is retained unconditionally, so a dispatch that called
  `renderChartObject` itself would link the chart emitter into every program that
  writes a slide, text-only ones included. Two reasons it is passed rather than
  registered into a module-level map. The chart-part-id counter in
  `package/assemble.ts` already records the first: a never-reset module global made
  the same input produce different bytes. And two decks built in one process must be
  able to disagree about which families they carry.

  A slide object whose family the table omits throws
  `slide/object-type-not-routed` rather than emitting nothing. It has already
  reserved a `<p:cNvPr>` id that connector bindings, animation targets and group
  bounds may point at, so a silent omission is a dangling reference, which
  PowerPoint reports as a repair.
- **The packager is handed its part contributors; it does not import them.** Which
  construct families put parts in a package is a `PartContributor[]` (`package/parts/shared.ts`),
  carried on `PackageSource` beside the renderer table. The presentation collects
  them from the construct families it was composed with. Three of those have
  parts to add: charts, comments, speaker notes. `package/assemble.ts` keeps the
  skeleton every deck has, calls those contributors at fixed points in it, and
  names no family itself.

  Same bundling constraint as the shape walk, on the other axis. This one is
  about which parts land in the zip rather than which XML lands in a slide.
  Naming `createExcelWorksheet` from the packager linked the whole of
  `gen/chart/` into every program that wrote a deck: plot modules, axes, chartEx
  sidecars, the embedded-workbook writer.

  The zip's part order and `[Content_Types].xml` are both byte-significant, so
  neither can depend on the order a caller happened to list contributors in. Each
  contributor declares an `order` and the packager sorts by it. Content-type
  entries arrive as data, grouped by the slot they land in, so
  `gen/opc/content-types.ts` keeps its fixed sequence and never learns which
  family asked.

  Three things stay on the core path deliberately. The chart-part-id pass is
  package-wide ordering rather than chart knowledge: it must number identically
  for a deck with no charts, and `gen/chart/chartex-xml.ts` derives its series
  GUIDs from the id it assigns. The media half of what used to be one
  chart-and-media function stays, because images are how any deck carries a
  picture; the chart contributor runs immediately before it per target, so the
  zip keeps its chart-then-media interleaving. And `ppt/theme/theme2.xml` stays
  with the theme it pairs with, even though only `notesMaster1.xml.rels`
  references it.
- `src/ooxml/` holds the schema facts that belong to **neither** half of the library:
  relationship-type URIs (`rel-types.ts`), the child-sequence order of each
  complexType a writer or an editor inserts into (`sequence.ts`), the `ST_`
  enumerations (`st-enums.ts`), and the two enum-validation policies
  (`check-enum.ts`). It exists because `src/gen/` and `src/read/` each used to
  keep a private copy of the same constants, and a divergence between them had no
  compile-time signal. A wrong rel URI matches nothing. An out-of-order child
  makes the part invalid, and PowerPoint reports that as a *corrupt file* rather
  than as a bad edit. Two properties keep it from regressing. Each successor list
  is **derived** by slicing one declared sequence rather than written out, and
  each `ST_` union is **derived** from the same tuple the validator checks
  against (`(typeof X)[number]`), so a type and its runtime list cannot drift.
  Nothing here knows whether it is being read or written.
- `src/read/` is layered. `read/opc/` is the package, part and relationship layer.
  `read/oxml/` is the DOM substrate plus the **pure** resolvers: `theme.ts` for
  colour and style-matrix resolution, `placeholder-inherit.ts` for what a
  placeholder inherits from its layout and master chain. `read/api/` is the
  navigable object model. `read/api/ops/` is the deck-level machinery that moves
  parts between packages: part copy, master registry, import, prune, rescale, and
  the `preserve`-mode `flatten.ts`.

  The line between `read/oxml/` and `read/api/ops/` is **mutation**. The
  resolvers answer "what would this be?" and build detached elements. The ops
  write the answers into a live part. That line is what lets the read model's
  getters and the import-time bake share one implementation, which keeps a colour
  reported before export equal to the colour written into the file.

  One file in `read/api/` is not object model. `presentation-imports.ts` holds
  the bodies of `Presentation`'s four import entry points, whose contracts stay
  on the class as the doc comments a caller reads. It is not an `ops/` module,
  because it is not independent of its caller: it reaches back into the deck for
  three `@internal` members it shares with the methods that stayed behind.

  Two `ops/` modules also import from `src/gen/`. `notes-master.ts` and
  `notes-author.ts` reach for `makeXmlNotesMaster`, `makeXmlNotesSlideSkeleton`
  and the element builder. That is deliberate, and it is the one place the read
  half depends on the write emitters. Carrying notes into a deck that has no
  notesMaster means *authoring* one, and an ops-local second copy of those two
  parts would be a second answer to what a notesMaster is. Everything else under
  `ops/` moves parts that already exist.
- `src/measure/` holds the calibrated text-measurement engine behind the
  `pptx-ts/measure` subpath and the export-time autofit bake: `font-metrics.ts`
  (advance widths + the registry), `text-fit.ts` (the wrap simulator and the
  shrink/resize solvers), `paragraphs.ts` (authored object → simulator inputs),
  `table-fit.ts` (`computeTableLayout` and the cell-grid walk), and `fit.ts` (the
  pass that measures and rewrites slide objects before the sync XML build).
  `src/measure.ts` is the public barrel over it. See `docs/contributing/design/text-fit.md`.
- `src/types/index.ts` and `src/enums.ts` define the public typed contract.
  `types/index.ts` is a re-export barrel over its siblings in `src/types/*` (split
  by domain). The generator-internal `*Internal` wire shapes live in
  `src/types/internal.ts` and are **not** re-exported. Internal code imports them from
  there directly. It is the same non-published convention as `units-internal.ts`
  (lenient unit conversion) and `constants-internal.ts` (generator defaults, fixed
  ids, colour palettes), each sitting beside the published module it extends.
- `src/script/` turns a deck read through `src/read/` into a serializable
  description of the write-API calls that would rebuild it (`readModelToIr`), and
  prints that description as a runnable TypeScript module. The two halves meet
  only at the IR: `from-read/` knows OOXML and the read model, `print/` knows
  only strings, and neither can see the other. That is what makes the mapping
  testable without a printer and keeps "how a number is spelled" from changing
  what a deck means. Two printers sit over the one IR, and they differ only in
  where the deck's *chrome* comes from. `printScript` anchors its output on a
  template, reusing the *source deck itself*: `fromTemplate` strips a package's
  slides while leaving masters, layouts, theme and document properties
  byte-identical, so only slide content is ever regenerated.
  `printStandaloneScript` emits a module that depends on nothing but this
  package. It re-authors the theme and one `defineSlideMaster` per source layout
  from what the read model exposes, including that layout's own decoration,
  re-tagged through the same mapper the slides use.

  The split is not a preference. A theme's `a:fmtScheme` and a master's
  `p:txStyles` are unreachable from *both* directions, and a master's own
  decoration has no write-side counterpart, because `defineSlideMaster` creates a
  layout. So a rebuilt design can only ever be an approximation, while a reused
  one is exact. Each printer therefore returns the note set that applies to
  **its** output, suppressing what its tier rescues and adding what its tier
  costs. That set, not `DeckIr.fidelity`, is what a round-trip check excludes.

  It is its own subsystem for two reasons. It depends on **both** halves, the
  read model and the write option types, so it fits inside neither. And
  `src/read/` is documented as isomorphic, bytes in and bytes out, which a
  converter emitting source text would quietly break for every `pptx-ts/read`
  consumer. Losses are data here, not log lines. Anything that cannot survive
  becomes a `FidelityNote` on the IR, which lets a round-trip check exclude
  exactly the declared losses and treat every other difference as a defect.

  `verify/` is that check. `canonicalDeckIr` reduces an IR to the form a
  comparison can use, dropping only values whose explicit and absent spellings
  are the same OOXML default. `diffDeckIr` compares the source deck's IR against
  the IR of the deck a generated script produced, with the printer's notes as the
  exclusion list. It compares IRs rather than packages because the output can
  never be byte-identical: fresh rel ids, regenerated shape ids. A byte
  comparison would fail for every deck and measure nothing.

  Its reach is bounded in two ways worth knowing before you trust a clean run.
  Both IRs come from the same reader, so a construct the read path cannot see is
  absent from both. And the converter need not be injective, so two source
  constructs that map to the same call compare equal. It detects *asymmetry*.
  `pnpm run read:census` and the IR unit tests cover the rest. The
  consumer-facing guide, covering both tiers, the measured loss list and how to
  read a fidelity note, is [PPTX to script](../reference/pptx-to-script.md).
- `scripts/package-smoke.mjs` verifies the packed package boundary from a
  consumer perspective.

## The rule that keeps the tiers real

Three seams carry a construct family to the write path instead of importing it: the
family list a presentation is composed with (`families/shared.ts`), the renderer table
the shape walk is handed (`gen/slide/objects/shared.ts`), and the part contributors the
packager is handed (`package/parts/shared.ts`). All three exist for one reason, and
one sentence covers the whole maintenance burden:

> **A static import from `slide.ts`, `gen/slide/object.ts` or `package/assemble.ts` into
> a family module is what the tier budget is watching for.**

Those three files are what every program that builds a deck reaches. A named import
inside a reachable function body is retained unconditionally, and a class method body is
never shaken at all, so one import added at any of the three puts that family in every
program's graph again, including the ones that compose without it.

Nothing else catches it. The types still check. Every test still passes. The emitted
bytes are identical, and `bundle-size:check` does not budge, because the package ships
exactly what it shipped before. One signal survives, `bundle-tier:check`, where the
composed rows climb toward the `TsPptx` rows until the gate fails on budget. That is
why the tier budget measures composed programs and not only class ones, and why the
`full` row sits beside them as a control.

The rule is about those three composition points, not about families in general. A
family module importing from `gen/` is ordinary and expected. That is where its
emitters live. What it must never do is get named *by* the core path.

## Where does X live? (task → file → function)

A starting point for "which function do I touch?". Two-phase pattern for most content:
an `add*Definition` in `gen/define/*` normalizes user options onto the slide model,
then a serializer under `gen/{slide,drawingml,chart,anim,pres,opc}/*` emits OOXML at
export time. Each module opens with a TSDoc header stating its job; larger files add
`// ===== region =====` banners: grep those to jump within a file.

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
| Slide number | `presentation.ts` `setSlideNumber` | `gen/slide/object.ts` `slideObjectToXml` (`SLDNUMFLDID`) |
| Transitions / animations | slide props (`slide.ts`) | `gen/anim/transition.ts` `slideTransitionToXml` / `gen/anim/animation.ts` `buildAnimationSeq` |
| Slide master / layout | `gen/define/master.ts` `createSlideMaster` | `gen/slide/master.ts` `makeXmlMaster` / `gen/slide/layout.ts` `makeXmlLayout` |
| Theme colors | n/a | `gen/pres/theme.ts` `buildThemeClrScheme` / `makeXmlTheme` |
| Coordinates & units (in → EMU) | `units.ts` (strict public primitives); `units-internal.ts` `getSmartParseNumber` (lenient generator layer) | n/a |
| Colors, fills, borders, shadows | n/a | `gen/drawingml/color.ts` `createColorElement`; `gen/drawingml/fill.ts` `genXmlColorSelection` / `genXml*Fill`; `gen/drawingml/line.ts` `genXmlLineFill` / `createLineCap`; `gen/drawingml/effect.ts` `createShadowElement` / `createGlowElement` |
| Package assembly & export | `package/assemble.ts` `buildPackageParts` (parts) + `zipPackageParts` (zip) → `writePackage` (behind `presentation.ts` `write` / `writeFile` / `stream`); `toParts` exposes the parts | `gen/opc/content-types.ts` `makeXmlContTypes` / `gen/opc/root-rels.ts` `makeXmlRootRels` / per-part rels |
| HTML `<table>` → slides | `html.ts` `tableToSlides` (the `pptx-ts/html` subpath, any DOM); `browser.ts` `tableToSlides` (method form, delegates) | `gen/table/html-dom.ts` `genTableToSlides` |
| Public API surface | `presentation.ts` (class), `slide.ts` (slide methods) | n/a |
| Option / type definitions | `types/index.ts` (barrel over `types/*`) | n/a |
| Enums & shared constants | `enums.ts` (public); `constants-internal.ts` (generator-only) | n/a |

## Boundaries

- The maintained runtime package is one ESM build, and every consumer loads it:
  Node (including through `require()`), bundlers, and browsers reaching it from a
  bundler or an ESM CDN. Upstream also shipped a CommonJS build and an IIFE global;
  this package replaced both with that single artifact.
- `dist/` is generated release output, not hand-edited source.
- Internal OOXML generators are implementation details unless deliberately
  exposed through `package.json` exports and public declarations.
- Platform differences go through the `RuntimeAdapter` seam (`src/runtime/*`): the
  `node`/`browser`/neutral entry subclasses inject the matching adapter into the
  shared core class. Live-DOM features that only work in a browser (currently
  `tableToSlides`) are defined on the browser entry subclass, not the core class. That
  keeps them off the Node build and out of the shared chunk. Their code bundles into
  the browser chunk alone.
- A runtime that resolves neither the `node` nor the `browser` condition gets
  `runtime/neutral.ts`. It implements what is genuinely host-neutral (`fetch`, `btoa`,
  `TextEncoder`, so remote media and fonts load), and throws
  `runtime/file-output-unavailable` from `writeFile` rather than substituting a host it
  does not have. The neutral adapter is the fallback, never a default the other two fall
  back *to*. A capability missing from a real host is a bug in that host's adapter, not
  something the neutral one should paper over.
- `src/read/` imports from `src/gen/` in exactly two places, and both are deliberate:
  `read/api/ops/notes-author.ts` and `read/api/ops/notes-master.ts` build a notes
  slide and a notes master through `gen/slide/notes.ts`. Adding a notes page to a
  loaded deck has to produce the same part the write path produces, and having two
  builders for one part is how they come to disagree. The cost is that
  `pptx-ts/read` pulls `gen/slide/notes.ts` and its `drawingml`/`opc` graph into its
  bundle; that is the trade, and it is stated here so it stays a decision rather
  than an observation. Everything else the two halves share lives in the
  `src/ooxml/` modules, and none of them imports from `src/gen/` or `src/read/`.
  Almost all of them import nothing outside `src/ooxml/` at run time either. The two
  exceptions fail loudly and reach out for that alone: `sequence.ts` imports
  `errors.ts`, and `check-enum.ts`, a warn-or-throw policy, imports `errors.ts` and
  `diagnostics.ts`. That is where a new shared schema fact belongs: a namespace, a
  relationship or content type a writer and a reader both name, a schema sequence or
  enumeration, or the XML prolog.
- Downstream deck-production workflows belong in the consuming project unless the
  behavior is broadly reusable for ts-pptx consumers.

## Data and control flow

1. A consumer creates a presentation through one of the three entry points.
2. `add*` calls pile slides and slide objects into the internal model. No XML yet.
3. Export walks that model, and the generators turn it into package parts.
4. The runtime adapter writes the result, in Node or in a browser.
5. The package smoke test then checks that only the supported entry points
   resolve from outside.

## Extension points

- Add public API only through exported entry points and generated declarations.
- Add OOXML behavior with focused regression or schema fixtures.
- Add package-boundary checks when changing public exports, runtime targets, or
  declaration output.
