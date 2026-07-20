# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- **BREAKING (`inspect`): `PptxSlideElement.box` is now slide-absolute, and `zIndex` is
  document order.** Two silent-wrongness fixes to the same read model, plus the group
  container itself:
  - `box` composes every enclosing `<p:grpSp>` transform (`off + (p - chOff) * (ext / chExt)`,
    with group `@rot`/`@flipH`/`@flipV`) instead of reporting the shape's raw `<a:xfrm>`.
    A group child's own transform is authored in the group's *child* space, so the old
    value was not placeable on the slide. It was right only by coincidence — this
    library's own writer hardcodes an identity child space (`chOff/chExt == off/ext`) —
    and wrong for any deck PowerPoint has touched, where resizing a group alone makes
    `chExt` non-identity. Measured against the `group-transform.pptx` fixture, a child of
    a scaled group was reported ~32% too wide and ~0.45in off-position, with its
    rotation absent. *Migration:* boxes of grouped children change value (top-level
    elements are unaffected); if you were compensating for this downstream, remove the
    workaround. Elements whose position cannot be resolved (an enclosing group with a
    degenerate zero `chExt`) are now omitted with a warning rather than reported wrong.
  - `zIndex` now follows a depth-first walk of `p:spTree` in document order, so it is
    real paint order. It previously came from the harvest order — every `<p:sp>` of a
    node, then every `<p:pic>`, then every `<p:cxnSp>`, then recurse — so an image
    authored between two text boxes sorted after both, and grouped children always
    sorted after every top-level shape. This was wrong for mixed-type slides with or
    without groups. *Migration:* if you relied on `zIndex` to key elements, note the
    values shift; a group's children now immediately follow it.
  - Added `kind: 'group'` for `<p:grpSp>`, carrying its `cNvPr` id/name and `grpSpPr`
    fill. Groups previously reached the output only as a side effect of the generic
    walker recursing into every object value, and the flat element list gave no
    indication which elements were grouped. `parentZIndex` (the enclosing group, or
    `null` at slide level) and `childZIndices` (a group's direct children in document
    order) now expose that structure, and `rotation`/`flipH`/`flipV` report effective
    orientation after group composition. *Migration:* consumers that assume every
    element is a leaf should skip `kind === 'group'` — its box overlaps its children by
    construction, so an overlap linter would otherwise report a false positive per group.

- **BREAKING (types): a freeform `arc` node no longer takes `x`/`y`.** An `<a:arcTo>`
  carries no explicit end point — PowerPoint derives it from the current pen position,
  the radii and the swept angle — and the emitter always discarded the authored `x`/`y`.
  Requiring them made the DSL demand a coordinate it ignored, and forced anyone
  converting an existing PowerPoint freeform into the DSL to invent one. The `arc`
  member of `GeometryPoint` is now `{ curve: { type: 'arc', hR, wR, stAng, swAng } }`.
  *Migration:* delete the `x`/`y` from arc nodes — e.g.
  `{ x: 0, y: h / 2, curve: { type: 'arc', … } }` becomes `{ curve: { type: 'arc', … } }`.
  Note that TypeScript's excess-property check does not reject the old form (`x`/`y`
  appear on other members of the `GeometryPoint` union), so existing code keeps
  compiling and behaves as before; a runtime warning now flags the ignored coordinate.

- **`addGroup()` with no renderable children now warns instead of silently emitting a
  zero-size group.** Auto-bounds over an empty child set is a `0×0` box, so `addGroup([])`
  — or a group whose children were all unsupported kinds and skipped — used to produce a
  degenerate `<p:grpSp>` with `<a:ext cx="0" cy="0"/>` with no signal. It now warns (naming
  the group), matching the partial-frame fallback and the project's rule against silently
  emitting degenerate geometry. The group is still emitted (not dropped), so the only change
  is the warning. This closed the last write-side coverage gap for groups; the new
  [Grouping objects](docs/groups.md) guide and `demos/modules/demo_group.mjs` runnable demo
  document `addGroup()`/`groupObjects()`, and regression tests now cover group
  `rotate`/`flipH`/`flipV`, `objectLock` (plus its unsupported-flag warning), `altText`, and
  the empty-group case.

### Removed

- **BREAKING: the duplicate unit constants `EMU` and `ONEPT` are removed from
  `core-enums.ts`.** They were aliases for `EMU_PER_INCH` (914400) and `EMU_PER_POINT`
  (12700) in `units.ts`, and both pairs were exported from every entrypoint — two
  published names for the same number, which is how unit mix-ups survive. `units.ts` is
  now the single source for unit constants. *Migration:* `EMU` → `EMU_PER_INCH`,
  `ONEPT` → `EMU_PER_POINT`; the values are unchanged, so this is a rename only. All
  internal call sites were migrated in the same change and the demo decks re-emit
  byte-identically (1,437 package parts).

- **BREAKING (types): the legacy `I`-prefixed type aliases are removed.** The ~26
  deprecated `@deprecated v4.0.0` aliases in `core-interfaces.ts` (e.g. `IChartOpts`,
  `IChartMulti`, `ISlideObject`, `ISlideRelMedia`, `IPresentationProps`) — thin
  re-exports of the un-prefixed names kept only for older imports — are gone. They
  had outlived their transition window and stood in tension with the fork's
  "no external backward-compat obligation" policy. *Migration:* drop the `I` prefix
  (`IChartOpts` → `ChartOpts`, `ISlideObject` → `SlideObject`, etc.); the two internal
  shapes map to the `*Internal` convention (`IChartOptsLib` → `ChartOptsInternal`,
  `IOptsChartData` → `OptsChartDataInternal`, `IPresentationProps` → `PresentationPropsInternal`).
  These were type-only exports, so runtime behavior and emitted OOXML are unchanged.

- **BREAKING: removed two unused exported enums/constants from `core-enums.ts`.**
  `DEF_SLIDE_BKGD` (a `'FFFFFF'` constant) and `MasterObjectType` (an enum of
  `chart`/`image`/`line`/`rect`/`text`/`placeholder`) had no internal callers —
  slide-background defaulting never read `DEF_SLIDE_BKGD`, and `MasterObjectType`
  was a leftover from the `MASTER_OBJECTS` → `MasterObjectType` rename that was
  never wired back up to anything. Both were still exported through the public
  `core-enums.js` barrel (`index`/`core`/`browser`/`node`/`standalone`), so removing
  them is a public-API break even though nothing in this codebase used them.
  *Migration:* there is no replacement — inline the literal `'FFFFFF'` if you were
  importing `DEF_SLIDE_BKGD`, and drop any reference to `MasterObjectType` (it was
  never read to distinguish object kinds; use `SlideObjectType` for that).

- **BREAKING: `AddSlideProps.masterName` is removed — use `masterTitle`.** The
  `@deprecated v4.0.0` alias (kept for older call sites, and consistent with the
  slide master's own `title`) has outlived its transition window. *Migration:*
  `pptx.addSlide({ masterName: 'X' })` → `pptx.addSlide({ masterTitle: 'X' })`.

- **BREAKING: the `radarStyle` wire spellings `'standard'`/`'marker'` are removed —
  use `'radar'`/`'markers'`.** These `@deprecated v4.0.0` aliases were the raw
  ECMA-376 `ST_RadarStyle` values; the PowerPoint-UI-facing names have been
  canonical since v4.0.0. *Migration:* `radarStyle: 'standard'` → `'radar'`,
  `radarStyle: 'marker'` → `'markers'` (`'filled'` is unaffected).

- **BREAKING: `ShadowProps.opacity` is removed — use `transparency`.** The
  `@deprecated v4.0.0` 0.0-1.0 alias for the PowerPoint-UI-aligned `transparency`
  (0-100) option is gone. `opacity` also doubled as the internal wire-normalized
  alpha every emit site reads; that internal shape is now the (non-public)
  `ShadowPropsInternal` type, so a caller can no longer reach it by constructing
  an object with an `opacity` field — a legacy/untyped caller still passing one
  is silently ignored (falls back to the shadow's default alpha) rather than
  accidentally still working. *Migration:* `opacity: 0.75` → `transparency: 25`
  (`transparency = (1 - opacity) * 100`).

- **BREAKING: the positional `addChart(type, data, options)` overload is removed —
  pass `type` on the options object.** The `@deprecated v4.0.0` three-argument form
  predates the canonical `addChart(data, { type, ...options })` / combo-chart
  `addChart(charts, options)` signatures. *Migration:*
  `addChart(pptx.ChartType.bar, data, options)` →
  `addChart(data, { type: pptx.ChartType.bar, ...options })`.

### Fixed

- **The slide-number placeholder no longer emits a hardcoded `cNvPr` id that can collide.**
  The placeholder was written with a literal `<p:cNvPr id="25">`, while every other shape
  allocates its id `idx + 2` from the slide's objects and group children allocate ids past
  that range — so a sufficiently populated slide (24 top-level objects is enough on its own;
  groups reach it sooner) emitted id 25 twice, a duplicate `<p:cNvPr>` id PowerPoint reports
  as a repair (`0x80070570`). The placeholder now takes its id from the same monotonic
  counter as every other object (the next free slot after the whole object/group walk), so it
  cannot alias a shape or group-child id regardless of slide population.
- **An out-of-range animation `shapeIndex` now warns and drops the effect instead of
  emitting a dangling target.** `addAnimation({ shapeIndex })` validated only that the
  index was `>= 0`, then emitted `spid = shapeIndex + 2` with no upper bound — so an index
  past the last top-level object produced a `<p:spTgt spid>` naming no shape on the slide, a
  dangling target PowerPoint reports as a repair (`0x80070570`). The index is now bounded to
  the slide's top-level objects (the range whose `spid` mirrors the writer's `id + 2`
  allocation; group children remain addressable by `objectName` only, as documented). An
  out-of-range index warns (`shapeIndex N is out of range (slide has M top-level object(s))`)
  and drops the effect, exactly like an unresolvable `objectName` — no timing tree, no
  dangling spid. A numeric `shapeIndex` is now also handled exclusively: a negative value
  warns and drops rather than silently falling through to `objectName`, surfacing the caller
  error instead of masking it.
- **Connectors and animations can now reference a shape inside a group by `objectName`.**
  Both resolved names only against the slide's top-level object list, which `addGroup()`
  splices its children out of, so both silently failed on a grouped target: an animation
  was dropped with no warning at all, and a connector binding fell back to static endpoint
  coordinates while warning `no shape with that objectName on the slide` — a shape that did
  exist, and was visible in the Selection Pane. Group children are `<p:cNvPr>`-named on the
  same slide and share its id space, so they are now resolved like any other object, at any
  nesting depth. Verified in desktop PowerPoint: a connector bound to a grouped shape at
  both ends reports those shapes as its connected endpoints, and animations targeting
  grouped shapes appear on the correct shapes in the timeline. Where a name is genuinely
  unresolvable, the connector's warning now says `no object with that objectName`, and an
  animation that names a missing object — or names no target at all — warns that its effect
  was dropped instead of vanishing silently. A top-level object still wins over a group
  child of the same name, so no existing deck's bindings change.
- **BREAKING (`addGroup`): a partial group frame no longer emits a degenerate group.** The
  frame is now all-or-nothing — pass all four of `x/y/w/h` to set it explicitly, or none to
  get auto-bounds (the bounding box of the children). Passing *some* axes previously let the
  unset ones fall through to the shared per-object defaults (`x=0`, `y=0`, `cx=75%` of the
  layout width, and `cy=0`), so `addGroup([rect], { x: 5, y: 2 })` silently emitted a
  zero-height group whose width was a slide-width fraction; on re-read every child of it
  resolved to `null` through the degenerate-`chExt` guard. A partial frame now warns and
  falls back to auto-bounds on every axis. *Migration:* calls passing a partial frame change
  geometry — from a broken group to its children's bbox — and gain a warning; pass all four
  axes to keep an explicit frame. Note that per-axis fallback was rejected deliberately: the
  writer keeps an identity child space (`chOff/chExt == off/ext`), so a group's frame never
  moves or scales its children — it only places the selection handle and the rotate pivot.
  `{ x: 5 }` would have left the children where they were and put the group's box somewhere
  they are not, which reads like a reposition but is not one.
- **Default object names no longer collide across a group boundary, and group names no
  longer depend on what else the process has built.** Two defects in one naming model:
  - `addGroup()` moves its children off the slide's top-level object list, but default
    Selection Pane names (`Shape 0`, `Text 1`, `Image 0`, …) were derived by *counting*
    that list — so a grouped child never advanced the count and the next top-level object
    of the same kind reused its name. `addGroup([{ rect }])` followed by
    `addShape('rect', …)` emitted `name="Shape 0"` twice on one slide. (`<p:cNvPr>` ids
    were already unique, so packages were valid; the collision was in Selection Pane
    identity, which name-keyed consumers rely on.) Names now come from a monotonic
    per-slide, per-kind counter that a group child consumes at any nesting depth.
    *Migration:* an object that follows a group in add order takes a higher default
    index than before; pass an explicit `objectName` if you depend on the exact string.
  - `Group N` counted from a module-global that was never reset, so three identical,
    independent presentations built in one process named their groups `Group 1`,
    `Group 2`, `Group 3` — same input, different bytes. Group names are now per slide
    and 1-based, matching PowerPoint's own default. *Migration:* the first default group
    on every slide is `Group 1`.
  - The duplicate-`objectName` warning now recurses into groups. It mapped only
    top-level objects, so it could not see either collision above — or a name a caller
    explicitly duplicated across the boundary.

- **Freeform arc angles are no longer wrapped into `0..360`, and a non-finite angle now
  throws.** `stAng`/`swAng` on an `<a:arcTo>` were converted with the *shape rotation*
  helper, which wraps once (`d > 360 ? d - 360 : d`) and coerces a nullish or `NaN` input
  to `0`. A sweep is not modular: `swAng: 400` silently drew a 40° arc instead of a 400°
  one, and a `NaN` angle from upstream arithmetic silently drew a zero-length arc rather
  than reporting the mistake. Arc angles now convert via a dedicated `convertArcAngle`,
  which preserves the authored angle and throws on a non-finite value. Shape/chart-label
  rotation is unaffected.

### Added

- **`slide.groupObjects(objectNames, options?)` groups objects already on the slide by
  `objectName`.** The counterpart to `addGroup()` for slides composed from independent
  renderers: those renderers each emit finished objects with stable names, so grouping
  them after the fact previously meant re-authoring every child descriptor through
  `addGroup()`. `groupObjects()` instead lifts the named top-level objects into one
  `<p:grpSp>` in place. Grouping is visually a no-op — children keep their slide-absolute
  geometry, their ids, their rels, and their relative z-order; the wrapper takes the
  topmost member's former slot in the stack, and the children's order follows the existing
  z-order rather than the order they are named (naming order is a selection, not a
  restack). The frame is all-or-nothing exactly as with `addGroup()` (omit `x/y/w/h` for
  auto-bounds over the members). Existing groups may be named, so consumers can nest logical
  groups. Every failure throws rather than warns — a name that matches nothing, matches an
  object already inside another group, is ambiguous across two same-named objects, or names
  an ungroupable kind (charts, media, tables, placeholders) — because the alternative is
  leaving the intended object silently loose on the slide. Verified in desktop PowerPoint:
  the nested tree opens without repair, ids stay unique, and a connector bound to a member
  still resolves after the lift.

- **Measured `fit:'shrink'` now shrinks non-wrapping text that overflows horizontally.**
  The measured-fit pass already baked a real `fontScale` onto `<a:normAutofit>` for
  vertical overflow (text taller than its box) when a font's metrics are registered via
  `registerFontMetrics`. But `normAutofit` is a vertical mechanism, so a `wrap:false`
  (`wrap="none"`) frame never triggered it: a single unwrapped line always fits the box
  height, so no scale was baked and the line ran out of the box sideways in PowerPoint
  (LibreOffice re-derives shrink-to-fit at render time and hid it). The shrink solver now
  carries the frame's wrap mode: a non-wrapping frame lays out one line per paragraph and
  the solver additionally keeps the widest line within the box width, on the same 2.5%
  `fontScale` grid and with the same conservative width inflation, so the baked scale
  never overflows PowerPoint. Gated on registered metrics, so decks that do not opt into
  measured fit are unchanged; wrapping frames, table cells, and `measureText` are
  unaffected.

- **`inspect` read model exposes run font properties, paragraph structure, and the baked
  autofit `fontScale`.** `PptxTextRun` now carries `fontFace`, `bold`, `italic`, and
  `charSpacingPt` alongside `text`/`fontSizePt`/`color`; `PptxSlideElement` gains
  `paragraphs: PptxParagraph[]` (runs grouped by their source `<a:p>`, preserving the
  line boundaries the flat `textRuns` list discards) and `autofitFontScale` (the
  `<a:normAutofit@fontScale>` value as a percent, or `null` for a bare `<a:normAutofit/>`).
  This lets a consumer re-derive a frame's rendered text extent — e.g. an overflow linter
  that must measure each `wrap="none"` line at the size PowerPoint will actually draw it —
  entirely off `inspectPptx`, with no raw slide-XML parsing of its own. All fields are
  additive; existing fields are unchanged.

### Internal

- **The `gen-*.ts` re-export barrels are gone (no API change).** `gen-charts.ts`,
  `gen-objects.ts`, `gen-tables.ts` and `gen-xml.ts` had no behavior of their own — they
  only forwarded to the `gen/**` tree so that `pptxgen.ts` and `slide.ts` could keep doing
  `import * as genXml from './gen-xml.js'`. Both costs are now paid off: the indirection
  layer is deleted, and the namespace imports (which defeat tree-shaking, notable for a
  package that declares `"sideEffects": false` and ships ten export subpaths) are replaced
  by named imports direct from `gen/**`. `gen-media.ts` also moved to `gen/media.ts`, so it
  no longer reads as part of the unrelated `media/` directory. None of the barrels were
  listed in `package.json` `exports` or re-exported from an entrypoint, so the published
  surface is byte-for-byte the same (1,150 exports across all ten entrypoints) and the demo
  decks re-emit byte-identically (1,437 package parts).

- **`gen-utils.ts` unit conversion and media decoding split out (no API change).** The
  repo's highest-fan-in module was three unrelated concerns in one file. Unit/number
  conversion moved to `units-internal.ts` — the lenient, warning-emitting layer over the
  strict public primitives in `units.ts` — and base64/image-header decoding plus OPC media
  content types moved to `media/base64.ts`, `media/content-type.ts` and `media/image-size.ts`
  (top-level, since both the write and read sides use them). `gen-utils.ts` drops from 1,062
  to 566 lines and is now only DrawingML fragment builders and naming/XML-escaping helpers.
  Pure code motion, verified two ways: the published `.d.ts` surface is unchanged across all
  ten entrypoints (1,160 exports, no diff), and the demo decks re-emit byte-identically
  (1,437 package parts).

- **`core-interfaces.ts` split into a `src/types/` tree (no API change).** The 3,491-line
  god-module — the one file every other module imports from — was cut along its own
  `// <name> ====` banners into 13 domain modules (`core`, `style`, `object`, `theme`,
  `text`, `media`, `shape`, `table`, `chart`, `animation`, `master`, `slide`, `pres`) and
  reduced to a re-export barrel, so `./core-interfaces.js` remains the single import site.
  Pure code motion: the published `.d.ts` surface was compared export-by-export across all
  ten entrypoints with aliases resolved, and all 1,155 pre-existing exports are identical.
  One additive change — `ImageBaseProps` was module-private but is extended by the public
  `ObjectOptions`, and is now exported by name.

- **Write-side emitter restructured into a layered `src/gen/` tree (no API change).**
  The 4,206-line `gen-xml.ts` monolith was split into focused modules mirroring the read
  side's `opc` / `pres` / `slide` / `drawingml` / `anim` layering, and reduced to a thin
  re-export barrel so `pptxgen.ts`'s `import * as genXml` is untouched. Pure code motion:
  every step was gated byte-for-byte against a full demo deck (1,437 OOXML parts, including
  recursing into embedded `.xlsx`), so emitted `.pptx` bytes are unchanged. No public types
  or exports changed.

- **Remaining write-side monoliths split into `src/gen/` too (no API change).** The same
  byte-identity-gated code motion was applied to the other four generators: `gen-tables.ts`
  → `gen/table/{autopage,html-dom}.ts`, `gen-charts.ts` → `gen/chart/{data-refs,chart-xml,embed-xlsx}.ts`,
  and the `add*Definition` layer of `gen-objects.ts` → one module per object kind under
  `gen/define/` (`group`, `chart`, `image`, `media`, `notes`, `comment`, `shape`, `connector`,
  `table`, `text`, `placeholder`, `background`, plus shared `object-name` / `hyperlinks` helpers).
  Each original file is now a thin re-export barrel, so `pptxgen.ts` / `slide.ts` import sites
  (`genCharts.*` / `genObj.*`) are untouched. Every step gated byte-for-byte against the full
  demo deck (1,437 OOXML parts, incl. embedded `.xlsx`); emitted `.pptx` bytes and all public
  types/exports are unchanged. This completes the write-side restructure begun with `gen-xml.ts`.

- **XML element builder added for the write side, and the byte-identity gate committed
  (no API change).** `src/gen/oxml/el.ts` (`el` / `voidEl` / `raw`) centralizes attribute
  and text escaping so an emitter cannot omit `encodeXmlEntities`, mirroring the read
  side's `src/read/oxml/dom.ts`. It reproduces the existing byte layout exactly —
  self-closing is decided by which function you call rather than by the child's value
  (`encodeXmlEntities(undefined)` is `''`, so a value-based rule would rewrite
  `<dc:title></dc:title>` as `<dc:title/>`), and whitespace is placed explicitly per
  element. `docProps/core.xml` and `_rels/.rels` are the first parts migrated onto it,
  byte-for-byte unchanged. The harness that proves this is now a committed script rather
  than an ad-hoc one: `pnpm run byte-identity:baseline` / `:check`.

## [10.4.0](https://github.com/shbernal/PptxGenJS/releases/tag/v10.4.0) - 2026-07-16

### Added

- **`importSlideMasters({ primary })` moves the grafted masters to the front of
  `p:sldMasterIdLst`.** By default a grafted master is appended after the ones the deck
  already had, so a brand master grafted into a generated deck trails the generator's
  stock master. List order is not part of theme resolution — a slide resolves through
  its own layout's master, so no existing slide changes appearance — but the list's
  first entry is the deck's `Designs(1)`: the theme PowerPoint's Design tab shows and
  the one Design ▸ Variants applies. With `primary`, the grafted masters lead (in import
  order) and the deck presents as their theme. Off by default: which master leads is a
  statement about what the deck *is*, so it is the caller's call, not a side effect of
  grafting. Reordering rewrites the id list only — relationships, ids, and every part
  outside `presentation.xml` are untouched, and a re-call is a no-op once the grafted
  masters already lead.

- **`importSlideMasters({ tableStyles })` carries the source deck's table styles.**
  Grafting a master shipped its layout gallery but not the deck's table styling, which
  lives in the presentation-level `ppt/tableStyles.xml`. A table then inserted on a
  grafted layout resolved against the *destination's* table styles instead — for a
  generated deck, this library's own stub, which defines zero styles and defaults to the
  standard *Medium Style 2 - Accent 1*. The result was visible and easy to mistake for a
  theme bug: the same table rendered in a different accent than it did in the source
  deck, even though the grafted master, its layouts, and its theme were all correct.

  The option carries the source's whole list. Styles union by `styleId` — a style the
  destination already defines wins, so a re-call is idempotent, matching the
  embedded-font carry's de-dupe. `a:tblStyleLst@def` (the default table style) is
  **source-wins**: carrying the styles without it does not fix the mismatch, because the
  standard default GUID is one most templates *also* define, so a new table would still
  resolve to the wrong style rather than visibly to none. A caller opting into the
  source's table styling wants its default too, and a destination `def` is typically a
  generator stub rather than a deliberate choice.

  **Off by default**, like `embedFonts`: unlike the rest of the graft it rewrites an
  existing part rather than only adding new ones. The carry is whole-deck — it copies
  *all* the source's table styles, not only those the grafted masters use, because
  `tableStyles.xml` does not record which style belongs to which master. Slides are
  untouched either way: the graft still ships a gallery without applying it.

- **`importSlideMasters({ embedFonts })` carries the source deck's embedded fonts.**
  Grafting a master brought its whole layout family, theme, and media across, but
  presentation-level embedded fonts (`p:embeddedFontLst`) were left behind — a
  documented v1 limitation. A grafted layout whose text depends on an embedded face
  therefore fell back to a substitute on any machine lacking that face locally,
  making the graft's fidelity machine-dependent. The option closes that gap, so a
  shipped layout gallery can be self-sufficient.

  This is the same carry `importSlide({ embedFonts })` already performed, and shares
  its implementation: binaries land under fresh `/ppt/fonts/` names (deduped through
  the per-source copy registry, so a re-call stays idempotent), the
  `application/x-fontdata` Default is added, and entries merge into the destination's
  `p:embeddedFontLst` de-duplicated by `typeface` + face slot. `importSlideMasters`
  was the only import path lacking it.

  **Off by default**, matching `importSlide({ embedFonts })`: fonts live on the
  presentation rather than the master, and carrying them can add megabytes, so the
  embed travels only when asked for. The carry is whole-deck — it copies *all* the
  source's embedded fonts, not only faces the grafted masters use, because
  `p:embeddedFontLst` does not record which face belongs to which master.

## [10.3.0](https://github.com/shbernal/PptxGenJS/releases/tag/v10.3.0) - 2026-07-15

### Added

- **`TextMeasurement.approximatedFaces` reports which faces `measureText()` had to
  guess at.** The resolver falls back to a conservative average-advance heuristic for
  any **named** face with no registered metrics, but a layout-time caller had no way
  to tell an exact measurement from an approximate one — the signal existed
  (`makeRegistryResolver`'s `onHeuristic`, used by the export pass to warn once) and
  `measureText` simply discarded it. It now collects those faces and returns them:
  `[]` when every run was measured exactly (and for an unmeasurable result, which
  guessed nothing), otherwise the named faces that fell back. Numbers are unchanged
  and still conservative; this only makes their provenance visible, so a caller
  needing exact values can check instead of assume. No warning is emitted —
  `measureText` is a per-layout query, so warn-once would either spam or go stale;
  the caller decides.

  **Breaking for implementers only.** `approximatedFaces` is a required field on the
  `TextMeasurement` interface. Consumers *reading* the result of `measureText()` /
  `pptx.measureText()` are unaffected — this is purely additive for them. Anyone
  *implementing* or hand-mocking `TextMeasurement` (e.g. a test double) must add
  `approximatedFaces: []` to satisfy the type.

### Fixed

- **The layout-time/export-time "never disagree" invariant is now stated accurately.**
  `measureText`'s docstring and `docs/measured-text-fit.md` claimed a layout-time
  prediction always matches what the export bakes. That holds only for a deck that
  opted into measured fit (registered ≥1 face). With an **empty** registry the two
  intentionally diverge: `applyMeasuredFit` reads "no metrics" as "not opted in" and
  bakes nothing, while `measureText` returns heuristic numbers so the API is useful
  with zero setup. Behaviour is unchanged and correct — the docs oversold it. The
  guarantee is now scoped, the asymmetry documented as deliberate, and a regression
  test pins the divergence so it cannot change silently.

## [10.2.0](https://github.com/shbernal/PptxGenJS/releases/tag/v10.2.0) - 2026-07-15

### Fixed

- **The same SVG file used more than once on a slide no longer leaves raw SVG
  bytes in a `.png` fallback part.** Each placement of one SVG *file* pushes its
  own png-fallback relationship (`isSvgPng`) that shares the svg's path. The load
  step only converted the path-unique primary; the path-duplicate fallbacks had
  the primary's data copied but never ran `createSvgPngPreview`, and the later
  catch-all step's `rel.isSvgPng && rel.data` filter ran synchronously — before
  the async load populated the duplicate's data — so the duplicate was skipped by
  both paths and kept raw SVG bytes in a `.png` part. That content-type/magic
  mismatch made PowerPoint repair (and drop) the deck. Duplicates are now
  converted right after their data is copied, so no `.png` fallback part contains
  SVG bytes and every `svgBlip` picture survives.

## [10.1.0](https://github.com/shbernal/PptxGenJS/releases/tag/v10.1.0) - 2026-07-09

### Fixed

- **Media (`addMedia`) no longer corrupts decks that pair media with other shapes.**
  A media picture's `<p:cNvPr>` id was computed as `mediaRid + 2` (the media
  relationship id), a different numbering space than the `index + 2` every other shape
  uses. When a sibling shape's slide-object index equalled the media's relationship id,
  the two collided into a **duplicate `cNvPr` id** and PowerPoint rejected the file as
  corrupt/unreadable (`0x80070570`). A second instance of the same root cause: the
  slide-level `<p:timing>` node for looping media (`loop`/`loopCount`) targeted the
  picture by `spid = mediaRid + 2`, which desynced from the shape id and pointed the
  playback timing at the wrong (or a nonexistent) shape — same corruption. Both now use
  the slide-object index (`index + 2`), consistent with animation spids, so every id is
  unique per slide and the timing target always resolves to its own media picture. This
  affected any slide mixing media with text/shapes — including the bundled Node demo,
  whose media slides previously would not open in PowerPoint.

### Added

- **`TableProps.columns` — per-column cell styling for wide colored matrices**
  (`addTable`). A `TableCellProps[]` whose entry `columns[i]` is merged as direct
  per-cell formatting onto every cell starting in column `i` (`fill`, `color`, `bold`,
  `align`, `border`, `margin`, …), so a wide (~15-column) **colored** assessment /
  scorecard / maturity grid no longer needs a fill hand-written onto every cell.
  Entries may be sparse and the whole option is optional — omit it and output is
  unchanged (degrades cleanly to text-on-white). Precedence, matching how PowerPoint
  resolves styling (direct formatting overrides a style region): explicit per-cell
  `options` > `headerRow` (row 0) > `columns[colIdx]` > `tableStyle`/defaults. Because
  the merge is property-level, a **maturity-gradient header** is just shared typography
  on `headerRow` (bold/white/centered, no fill) plus a graduated `columns[i].fill` per
  column. The column index counts each cell's `colspan` within a row. There is
  deliberately no built-in "group bracket" annotation primitive: label a span of
  columns by composing existing shapes — `addShape('rightBrace', …)` (or `'bracePair'`)
  plus `addText`, positioned from the table's `x` and `colW`.
- **`Slide.text` — flatten all of a slide's text in one read** (`pptxgenjs/read`).
  A getter on the read-model `Slide` that walks the shape tree in document order and
  concatenates every text-bearing shape's text, **recursing into groups** and
  **reading table cells** (cells tab-joined within a row, rows newline-joined);
  text-free shapes (pictures, connectors, empty boxes) contribute nothing. It is the
  slide-level counterpart to `TextFrame.text` and closes a footgun in the raw API: a
  naive `slide.shapes.map(s => s.text)` silently drops grouped and tabular text,
  because `GroupShape`/`GraphicFrame` have no text frame of their own. Extract a whole
  deck with `deck.slides.map(s => s.text)`. Scoped to the slide's own shape tree —
  chart data labels are intentionally excluded (read those via `GraphicFrame.chart`),
  and speaker notes have their own accessor (below).
- **`Slide.notesText` — read a slide's speaker notes** (`pptxgenjs/read`). Resolves
  the slide's `notesSlide` relationship and flattens the notes **body** placeholder's
  text (the `sldImg` thumbnail and `sldNum` slide-number placeholders a notes slide
  also carries are ignored). Returns `null` when the slide has no notes slide part at
  all, distinct from `''` for a notes slide whose body is empty — a distinction
  PowerPoint makes routinely, since it often attaches an empty notes slide to every
  slide. Companion to `Slide.text`.

### Fixed

- **`chartColors: ['transparent']` now yields a fully invisible series** instead of
  a black one. The series/marker *fill* already honoured `'transparent'` (→
  `<a:noFill/>`), but the *stroke* paths (the connecting line and the marker border,
  for both the line/radar and scatter/bubble renderers) passed the literal
  `'transparent'` through colour validation — which warned "not a valid scheme color
  or hex RGB" and rendered the stroke as black `000000`. All four surfaces (series
  fill + line, marker fill + border) now consistently emit `<a:noFill/>` for a
  `'transparent'` entry, with no warning. Only affects the `'transparent'` value;
  output for real colours is byte-identical.

### Changed

- **BREAKING: renamed the exported type `JSZIP_OUTPUT_TYPE` → `ZIP_OUTPUT_TYPE`.**
  The name referenced JSZip, which was fully replaced by fflate; the type is just
  the set of supported ZIP output shapes (`arraybuffer` | `base64` | `binarystring`
  | `blob` | `nodebuffer` | `uint8array`) and no longer describes a JSZip-compat
  contract. The member list is unchanged, and `WRITE_OUTPUT_TYPE` (the type
  actually referenced by `WriteProps.outputType`) is unaffected. Migration:
  consumers importing `JSZIP_OUTPUT_TYPE` should import `ZIP_OUTPUT_TYPE` instead.

## [9.2.0](https://github.com/shbernal/PptxGenJS/releases/tag/v9.2.0) - 2026-07-07

### Added

- **`FontMetricsRegistry.hasCodepoint(face, cp, opts?)` — face-keyed glyph coverage**
  (`pptxgenjs/measure`). The per-face `FontMetrics.hasCodepoint(cp)` cmap check
  already existed; this adds a convenience on the registry that resolves the face
  (exact → regular → any-variant fallback, same as `get()`) and returns
  `boolean | undefined`, where `undefined` means the face has **no** registered
  metrics — "unknown", deliberately distinct from a `false` "registered but not
  covered". A coverage audit can now register the replica face once and query
  per-codepoint coverage in-process instead of shelling out to `fc-match`.
- **`Run.resolvedBold` — placeholder-inherited bold in the read API.** Completes the
  `resolvedSizePt`/`resolvedFontFace` inheritance trio: for a placeholder run that
  sets no own `@b`, it walks the paragraph `a:defRPr` → slide `a:lstStyle` → layout
  → master placeholder → master `p:txStyles` chain (via the existing
  `inheritedRunDefRPrs` tiers) and reports the first `@b` found. `null` when the run
  sets none and inherits none — distinct from an inherited explicit `false`.

## [9.1.0](https://github.com/shbernal/PptxGenJS/releases/tag/v9.1.0) - 2026-07-07

### Added

- **Inline (in-sentence) math via `inline: true`.** A text item's `math` OMML was
  previously always emitted as its own centered display-math paragraph
  (`<a14:m><m:oMathPara><m:oMath>…`). Setting `inline: true` alongside `math` now
  emits the equation as a bare `<a14:m><m:oMath>` run flowing mid-paragraph between
  the surrounding plain text runs (no `<m:oMathPara>`), so a single `addText` call
  can mix prose and equations in one line:
  `addText([{ text: 'where ' }, { math: latexToOmml('x^2+1=y', { display: false }), inline: true }, { text: ' holds' }])`.
  Pair `inline` with the bare-`<m:oMath>` converter form (`latexToOmml(tex, { display: false })`
  or `mathmlToOmml(mathml)`). The `mc:AlternateContent` (`Requires="a14"`) envelope
  stays at the shape level, exactly as for display math. Structure pinned by the
  PowerPoint-authored `math-omml-inline.pptx` oracle; completes backlog `dn-inline-math`.

## [9.0.0](https://github.com/shbernal/PptxGenJS/releases/tag/v9.0.0) - 2026-07-07

### Added

- **LaTeX / MathML → OMML converter: new `@shbernal/pptxgenjs/math` subpath.** The
  `math:` option on a text item takes raw OMML; the new subpath lets you author the
  equation in LaTeX or MathML instead. It exports `latexToOmml(latex, { display? })`
  and `mathmlToOmml(mathml)`, composing `temml` (LaTeX → MathML) and `mathml2omml`
  (MathML → OMML) over the pipeline `LaTeX → MathML → OMML`.
  - Usage: `import { latexToOmml } from '@shbernal/pptxgenjs/math'` then
    `slide.addText([{ math: latexToOmml('x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}') }], {…})`.
  - `temml` (MIT) and `mathml2omml` (LGPL-3.0-or-later) are **optional peer
    dependencies** — the core package doesn't pull them in; install them
    (`npm install temml mathml2omml`) to use this subpath. They are never bundled
    into the package output, so the LGPL dependency stays separate, replaceable, and
    opt-in. This subpath is **Node-only** (it loads the converters via
    `node:module`'s `createRequire`, keeping the API synchronous).
  - Scope: **display (block) math** only in v1; `display: false` yields a bare
    `<m:oMath>` for inline embedding. Invalid LaTeX **throws** with temml's parse
    position (no silent coercion). No LaTeX macro packages and no `mc:Fallback` raster
    (output relies on the `Requires="a14"` envelope understood by PowerPoint 2010+).
  - Downstream: a consumer's equation authoring can drop its external LaTeX→OMML
    step and call `latexToOmml()` directly once it adds the two peer deps.
  - See `docs/math-latex.md`. Implements the LaTeX/MathML leg of upstream issue #1456.

- **`TableProps.fit` is now a declared option.** Table-level `fit` (`'none'` /
  `'shrink'` / `'resize'` / the `TextFitShrinkProps` object form) was already
  honored at export time — a `'shrink'` table cascades the shrink policy to every
  cell that doesn't set its own `fit` — but the type omitted it, so `addTable(rows,
  { fit: 'shrink' })` failed to typecheck. The property is now on the interface,
  documented, and matching the existing runtime behavior. No behavior change.

- **Read model: `Shape` return types are now a narrowable discriminated union.**
  `pptxgenjs/read` exports `AnyShape = AutoShape | Picture | Connector |
  GraphicFrame | GroupShape` (discriminated on the `shapeType` literal), and the
  shape-returning API — `Slide.shapes` / `shapeById` / `shapeByName`,
  `GroupShape.shapes`, and `Presentation.importShape` / `importShapes` — now
  returns `AnyShape` instead of the abstract base `Shape`. A consumer can reach a
  subtype's own members without a cast by narrowing on the discriminant:
  `if (shape.shapeType === 'graphicFrame') shape.chart`. Five type-guard helpers
  (`isAutoShape`, `isPicture`, `isConnector`, `isGraphicFrame`, `isGroupShape`)
  are exported for the same narrowing in expression position (`.filter(isPicture)`).
  - This is additive at runtime (the objects are unchanged) and non-breaking for
    TypeScript consumers: `AnyShape` is assignable to `Shape`, so code that
    annotated variables as `Shape` still compiles. The abstract `Shape` remains
    exported for the common-members-only case and for `instanceof` checks.

- **`addChart(data, options)` — chart type on the options object.** The canonical
  signature is now `slide.addChart(data, { type, ...options })`, matching how every
  other option is passed. This is additive: the old positional form
  `addChart(type, data, options)` still works but is now `@deprecated` and logs a
  one-time console warning; it will be removed on the fork's normal breaking-change
  cadence.
  - Migration: move the leading `type` argument onto the options object.
    - Before: `slide.addChart(pptx.ChartType.bar, data, { x: 1, y: 1 })`
    - After: `slide.addChart(data, { type: pptx.ChartType.bar, x: 1, y: 1 })`
  - Multi-type (combo) charts are unchanged: `addChart(chartsArray, options)` — each
    `IChartMulti` entry already carries its own `type`.
  - Omitting `type` on the options object of the canonical form now throws with a
    clear message rather than silently emitting a typeless chart.

- **PowerPoint-aligned border and shadow style props.** These are additive: the
  new names read verbatim off the PowerPoint UI, and the old names still work but
  are now `@deprecated` and will be removed on the fork's normal breaking-change
  cadence.
  - `BorderProps.width` (points) — the canonical name for what was `pt`, matching
    `ShapeLineProps.width` and PowerPoint's "Line > Width" field. Applies to table
    cell borders and chart `plotArea`/`chartArea`/`dataBorder`. (The legacy `pt`
    alias has since been removed — see Removed below.)
  - `BorderProps.transparency` (0–100 percent) — new; emits `<a:alpha>` on the
    line fill, matching PowerPoint's "Line > Transparency".
  - `ShadowProps.transparency` (0–100 percent) — the value PowerPoint's shadow
    dialog actually shows, as a friendlier alias of the legacy 0.0–1.0 `opacity`
    (`transparency: 25` ≡ `opacity: 0.75`). `opacity` is deprecated; when both are
    set, `transparency` wins (with a warning).

- **`AddSlideProps.masterTitle` — renamed from `masterName`.** The property that
  selects which slide master a new slide uses is now `masterTitle`, matching the
  `title` you pass to `defineSlideMaster`. This is additive: `masterName` still
  works but is now `@deprecated` and logs a one-time console warning; it will be
  removed on the fork's normal breaking-change cadence. When both are set,
  `masterTitle` wins.
  - Migration: `addSlide({ masterName: 'MyMaster' })` → `addSlide({ masterTitle: 'MyMaster' })`.

- **Radar chart `radarStyle` values renamed to match the PowerPoint UI.** The
  canonical values are now `'radar'` / `'markers'` / `'filled'` (was
  `'standard'` / `'marker'` / `'filled'`). This is additive: the old
  `'standard'` / `'marker'` spellings still work but are now `@deprecated` and log
  a one-time console warning. The emitted OOXML (`<c:radarStyle val="…"/>`,
  `ST_RadarStyle`) is unchanged — the rename is a public-API name only.
  - Migration: `radarStyle: 'standard'` → `radarStyle: 'radar'`; `radarStyle: 'marker'` → `radarStyle: 'markers'`.

### Changed

- **BREAKING: `margin` is now inches everywhere, not points.** Both kinds of `margin`
  are now interpreted as **inches**, matching `x`/`y`/`w`/`h` and the values PowerPoint's
  own dialogs show (the table cell-margin field and the text-box internal-margin field are
  both inches in PowerPoint):
  - **Table cell / table-level `margin`** previously used a magnitude heuristic — a
    component `>= 1` was read as points, `< 1` as inches — so a legitimate 1-inch margin
    silently became ~1pt with no way to express it. The heuristic is gone.
  - **Text-box / placeholder / slide-number body `margin`** (`TextProps.margin`) previously
    read as straight points (e.g. the PowerPoint "Normal" default `[3.5, 7.0, 3.5, 7.0]` pt).
    It is now inches (`[0.05, 0.1, 0.05, 0.1]`).

  A component `>= 1` is still honored as inches but logs a one-time warning that it is
  likely a legacy points value. **Migration:** if you were passing points (e.g. `margin: 10`
  or `margin: [9, 18, 9, 18]`), divide by 72 (`10pt` → `~0.139in`, `18pt` → `0.25in`).
  Fractional margins (`0.05`, `0.1`, the defaults) are unaffected. Shadow `blur`/`offset`
  and border `width` remain points — PowerPoint shows those in points, so they already
  match. Implemented via a shared `marginToEmu` helper used by the cell XML emitter, the
  text-box/slide-number insets, the autoPage row-height pass, and the measured-fit pass.

- **BREAKING: enum names rationalized to one PascalCase set.** `core-enums.ts` had
  shipped two parallel enum styles — a modern PascalCase set (`ChartType`, `ShapeType`,
  `SchemeColor`, `AlignH`, `AlignV`, `OutputType`) and a legacy `SCREAMING_CASE` set that
  partly duplicated it. The `SCREAMING_CASE` type names are gone (a clean break, no
  deprecated aliases, per the fork's API-evolution policy). Migrate as follows:

  - **Value-identical duplicates — use the PascalCase twin.** Member keys are now
    lowercase/camelCase mirroring the (unchanged) OOXML token value:
    - `CHART_TYPE.BAR` → `ChartType.bar` (likewise `BAR3D`→`bar3d`, `BUBBLE3D`→`bubble3d`, …)
    - `SHAPE_TYPE.RECTANGLE` → `ShapeType.rect`, `SHAPE_TYPE.OVAL` → `ShapeType.ellipse`,
      `SHAPE_TYPE.OVAL_CALLOUT` → `ShapeType.wedgeEllipseCallout` (keys are the OOXML
      preset tokens; the old readable SCREAMING names are gone)
    - `SCHEME_COLOR_NAMES.ACCENT1` → `SchemeColor.accent1` (`TEXT1`→`text1`, `BACKGROUND2`→`background2`, …)
    - `TEXT_HALIGN` → `AlignH` (values and keys were already identical, e.g. `AlignH.left`)
  - **SCREAMING-only enums — renamed to PascalCase, member keys unchanged:**
    `MASTER_OBJECTS`→`MasterObjectType`, `SLIDE_OBJECT_TYPES`→`SlideObjectType`,
    `PLACEHOLDER_TYPES`→`PlaceholderType`, `BULLET_TYPES`→`BulletType`,
    `TABLE_STYLE`→`TableStyle`. `PlaceholderType.title`, `BulletType.DEFAULT`, and
    `TableStyle.MEDIUM_STYLE_2_ACCENT_1` keep their existing keys (they are the readable
    handle on opaque values, so they were not lowercased).
  - **`TEXT_VALIGN` → `TextAnchor`.** This is *not* the same enum as `AlignV`: `AlignV`
    is the friendly `top`/`middle`/`bottom`, while `TextAnchor` is the OOXML bodyPr anchor
    token `b`/`ctr`/`t`. Both are kept.
  - **Removed the deprecated `pptx.charts` / `pptx.colors` / `pptx.shapes` getters.** Use
    `pptx.ChartType` / `pptx.SchemeColor` / `pptx.ShapeType` (which already existed).
    Note this is unrelated to the read-model `slide.shapes` array, which is unchanged.

- **`TableCell.text` no longer accepts `number` in its TypeScript type.** The type
  is now `string | TableCell[]` (was `string | number | TableCell[]`). Plain-JS
  callers passing a number are still coerced to a string at runtime, but
  TypeScript callers should pass `String(n)`. This removes a `number` branch that
  every consumer of the type had to account for.

- **Dropped the legacy `I` prefix from exported interface names.** The chart and
  slide interface types are now un-prefixed for consistency with the rest of the
  public surface (`IChartOpts` → `ChartOpts`, `IChartMulti` → `ChartMulti`, the
  whole `IChartProps*` family, `ISlideObject` → `SlideObject`, `ISlideRel*` →
  `SlideRel*`, `ISlideComment` → `SlideComment`). Two internal augmented shapes
  moved to the codebase's `*Internal` convention: `IOptsChartData` →
  `OptsChartDataInternal`, `IPresentationProps` → `PresentationPropsInternal`.
  This is additive: every old `I`-prefixed name is retained as a `@deprecated`
  type alias, so existing imports keep compiling; the aliases will be removed on
  the fork's normal breaking-change cadence.
  - Migration: drop the `I` prefix in type imports/annotations, e.g.
    `import type { IChartOpts } from '...'` → `import type { ChartOpts } from '...'`.

- **Renamed the internal `ChartOptsLib` shape to `ChartOptsInternal`.** This type
  is `ChartOpts` plus the internal `_type` carrier and was never meant as a public
  entry point; it now follows the same `*Internal` convention as
  `OptsChartDataInternal` / `PresentationPropsInternal`. Additive: the deprecated
  `IChartOptsLib` alias now points at `ChartOptsInternal`, so existing imports keep
  compiling.

### Removed

- **BREAKING: `BorderProps.pt` removed — use `width`.** The deprecated `pt` border
  width alias (introduced alongside `width` earlier this same cycle) is gone;
  `width` is now the single source of truth for border line width (points) on table
  cell borders and chart `plotArea`/`chartArea`/`dataBorder`. `width` is also the
  internal carrier now, so no `pt`/`width` reconciliation happens at emit time.
  - Migration: `border: { pt: 1 }` → `border: { width: 1 }`.

- **BREAKING: removed the long-deprecated compatibility aliases.** These options
  had been marked `@deprecated` (mostly since v3.3.0) but were still silently
  accepted and coerced to their modern equivalents. Per the fork's API-evolution
  policy (no external backward-compat obligation; silent coercion of legacy input
  is a footgun), they are now removed — passing them has no effect. Migrate as
  follows:

  - **Slide/master/chart background** — `bkgd` → `background`. The `slide.bkgd`
    getter/setter is gone; use `slide.background = { color: 'FF0000' }` (or
    `{ path }` / `{ data }`). On `defineSlideMaster`, replace `bkgd: '…'` with
    `background: { color: '…' }`. `AddSlideProps` never applied a background, so
    a per-slide background must be set on the returned slide object.
  - **Background color/source** — `background.fill` → `background.color`;
    `background.src` → `background.path`.
  - **Color transparency** — `alpha` → `transparency` (on solid fills, gradient
    stops, and image fills). The 8-char RGBA hex form (e.g. `'0000FF40'`) is
    unaffected.
  - **Shape/text line** — the flat `line: '<color>'`, `lineSize`, `lineDash`,
    `lineHead`, `lineTail` forms → the `line` object: `line: { color, width,
    dashType, beginArrowType, endArrowType }`. The `ShapeLineProps` inner aliases
    `pt`/`size` (→ `width`), `lineDash` (→ `dashType`), `lineHead` (→
    `beginArrowType`), `lineTail` (→ `endArrowType`) are removed. (Chart
    `lineSize`/`lineDash` options are unaffected — they were never deprecated.)
  - **Bullets** — `code` → `characterCode`, `startAt` → `numberStartAt`, `style`
    → `numberType`; the dead `marginPt` alias (→ `indent`) is removed.
  - **Text fit/inset** — `autoFit` and `shrinkText` → `fit` (`'resize'` /
    `'shrink'`); `inset` → `margin`.
  - **Underline** — the deprecated `underline: '<style>'` string form → the
    object form `underline: { style }`. (`underline: true` shorthand still works.)
  - **Table auto-paging** — `newSlideStartY` → `autoPageSlideStartY`;
    `addHeaderToEach` → `autoPageRepeatHeader`.
  - **Chart plot area** — the flat `border` / `fill` chart options →
    `plotArea.border` / `plotArea.fill`.
  - **Shape name** — the dead `ShapeProps.shapeName` alias (→ `objectName`) is
    removed (it was never wired up).
  - **Method signatures** — `write('<type>')` → `write({ outputType: '<type>' })`;
    `writeFile('<name>')` → `writeFile({ fileName: '<name>' })`;
    `addSlide('<masterName>')` → `addSlide({ masterName: '<masterName>' })`.
  - **Types** — the `ChartLineCap` alias is removed (use `LineCap`). The internal
    deprecated `DEF_CELL_MARGIN_PT` constant is removed.

  The pt-vs-inches unit reconsideration (`core-interfaces.ts` TODO) is a separate
  open design question and was intentionally left untouched.

### Fixed

- **Shadows now honor scheme colors (e.g. `accent1`) on shapes, images, and
  charts.** Shadow XML was emitted by four divergent copies of the same logic;
  three of them (`gen-charts.ts`, plus two inline blocks in `gen-xml.ts`)
  hardcoded `<a:srgbClr val="…">` and so emitted schema-invalid OOXML when a
  shadow `color` was a scheme-color constant. All four paths now route through the
  single `createShadowElement`/`createShadowEffectLst` helpers in `gen-utils.ts`,
  which build the color via `createColorElement` (scheme-aware) — a hex `color`
  is byte-identical to before, and a scheme `color` now correctly emits
  `<a:schemeClr>`. The chart copy also no longer forces `sx/sy/kx/ky/algn` onto
  `a:innerShdw` (valid only on `a:outerShdw`), and `shadow.rotateWithShape` is now
  honored on shape/image shadows, not just charts. A dead duplicate
  `correctShadowOptions` in `gen-xml.ts` (shadowed by the live one in
  `gen-utils.ts`) was removed.

### Changed

- **All library warnings now go through a single sink and carry a `PptxGenJS:`
  prefix.** A new `src/log.ts` exports `warn()` and `warnOnce()`; every
  `console.warn` call site across the source (~80, plus the ad-hoc
  `warnTextRangeOnce` deduper) now routes through them. Warnings gained a
  consistent `PptxGenJS: …` prefix (and shed the inconsistent `Warning:` /
  `[WARNING]` prefixes), so console noise is attributable to the library and
  there is now one place to later mute or redirect diagnostics via a handler.
  Message bodies are unchanged.

- **`createColorElement` is documented as the low-level color primitive it is,
  and manual `<a:solidFill>` wrapping now uses `genXmlColorSelection()`.** A stale
  TODO claimed `createColorElement` should become private with every call switched
  to `genXmlColorSelection()`; that was incorrect (the primitive is required for
  non-solid-fill color contexts — gradient stops, `<a:alpha>` on effects, line
  fills, highlight, …). The comment is corrected, and the ~34 sites that hand-wrote
  `` `<a:solidFill>${createColorElement(x)}</a:solidFill>` `` now call
  `genXmlColorSelection(x)` instead; emitted OOXML is unchanged.

- **`CHART_NAME` is now derived from the `CHART_TYPE` enum** (`` `${CHART_TYPE}` ``)
  instead of a hand-maintained duplicate string union, so the public chart-type
  name set and the enum can no longer drift — adding a member to the enum extends
  the accepted names automatically. The public type is identical (`'area' | 'bar' |
  …`) and `addChart()` signatures are unchanged, so this is not a breaking change.
  Internally, a chart's resolved `_type` and the `make*` chart-XML helpers now
  carry the `CHART_TYPE` enum, normalized once at the public boundary by the new
  `asChartType()` helper; emitted OOXML is byte-identical. This also removes the
  chart share of the suppressed `no-unsafe-enum-comparison` lint rule.

- **BREAKING: a `string` zip/inspect input is now a filesystem path, not latin1
  binary content.** `readZip`, `loadPptxPackage`/`inspectPptx`
  (`@shbernal/pptxgenjs/inspect`), and `OpcPackage.load` (`/read`) all previously
  treated a `string` input the way JSZip's `loadAsync` did — as a latin1 binary
  *content* string. That was a footgun: the natural call `loadPptxPackage("deck.pptx")`
  turned the path characters into bytes and failed with an opaque `Not a valid ZIP
  archive`. A string is now read from disk (Node, via lazily-imported `node:fs`),
  so `await loadPptxPackage("deck.pptx")` Just Works, and a missing file throws a
  clear error naming the path instead of a corrupt-archive error. In-memory
  archives are unaffected — keep passing `Uint8Array`/`ArrayBuffer`/`Blob`/`number[]`.
  Migration: to re-read a `binarystring`/`base64` write-path output, convert it to
  bytes first (e.g. `Uint8Array.from(atob(b64), c => c.charCodeAt(0))`) rather than
  passing the string. Downstream: a consumer's slide-library tooling can read
  `.pptx` parts (e.g. per-slide `_rels`) directly through the inspect API instead
  of shelling out to `unzip`.

### Added

- **`@shbernal/pptxgenjs/zip` subpath export.** The fflate-backed ZIP toolkit
  (`ZipWriter`, `readZip`, and the `ZipInput` type) is now a published entry
  point, so consumers that need to post-process a generated `.pptx` package
  (read parts, rewrite XML, re-zip) can reuse the same backend the writer uses
  instead of pulling in a second ZIP library. Downstream: `@shbernal/html2pptx`
  drops its `jszip` dependency in favour of this export.

### Fixed

- **Bubble-chart data labels now honour fractional `dataLabelFontSize`.** The
  bubble/bubble3D data-label block rounded the font size to a whole point before
  converting to the OOXML `sz` unit (hundredths of a point), so `dataLabelFontSize:
  10.5` emitted `sz="1100"` (11pt) instead of `sz="1050"` (10.5pt). It now converts
  directly like every other chart data-label site. Only fractional sizes on bubble
  charts are affected; integer sizes are unchanged. This surfaced while extracting
  the DrawingML unit factors (`60000` angle units, `100000`/`1000` percentage
  scales, `100` point-to-hundredths) into named constants/helpers in `units.ts`
  (`ANGLE_UNITS_PER_DEGREE`, `PERCENT_SCALE`, `FIXED_PCT_PER_PERCENT`,
  `HUNDREDTHS_PER_POINT`, `ptToHundredths()`); that extraction is otherwise
  byte-identical.

## [8.1.0](https://github.com/shbernal/PptxGenJS/releases/tag/v8.1.0) - 2026-06-26

### Added

- **Slide transitions & preset build animations (`docs/animations-and-transitions.md`, Phase 1):**
  two complementary subsystems, faithful to how PowerPoint authors the XML.
  - **Transitions — full typed model, both ways.** Write side: `slide.transition = { type,
    durationMs?, speed?, advanceOnClick?, advanceAfterMs?, variant? }` emits `p:transition`
    between `p:clrMapOvr` and `p:timing` — the bare `<p:transition>` form for a coarse `spd`
    speed bucket, or PowerPoint's `mc:AlternateContent` form (a `p14` Choice carrying the exact
    `p14:dur` plus a base `mc:Fallback`) when `durationMs` is set. Read side (`pptxgenjs/read`):
    `slide.transition` is a typed get/set accessor (`TransitionInfo`) handling both forms and
    preferring the `p14` Choice so `durationMs` round-trips. Types `TransitionProps` /
    `TransitionType` (the 21 base ECMA-376 types: `fade`, `push`, `wipe`, `cut`, `dissolve`, …),
    with type-specific variants (e.g. `{ dir: 'd' }`) via `variant`.
  - **Animations — opaque, spid-aware preservation + preset-template authoring.** The
    `p:timing`/`p:bldLst` build tree is modeled opaquely (no semantic AST): an unmodified slide
    round-trips it byte-identically, and the read model exposes only `slide.hasAnimations` plus
    the internal `spid` operations needed to keep references coherent — enumerate
    (`animationSpids()`), remap (`remapAnimationSpids()`), and prune (`pruneAnimationSpids()`)
    over `<p:spTgt @spid>`/`<p:bldP @spid>`. Authoring uses a fixed preset set (no general
    timing builder): `slide.addAnimation({ preset: 'fadeIn'|'flyIn'|'grow'|'fadeOut', shapeIndex
    |objectName, trigger?, durationMs? })` emits verbatim PowerPoint templates assembled into a
    `mainSeq`, grouped into click steps by `trigger` (`onClick`/`withPrevious`/`afterPrevious`),
    with one `<p:bldP>` per animated shape. `slideTimingToXml` was **extended** (not replaced) so
    a slide carries either the looping-media tree or a build-animation tree; the media-only path
    is byte-unchanged. The write emitters reproduce the PowerPoint-authored oracle decks
    byte-for-byte (asserted in tests). Implements backlog `gitbrent/PptxGenJS#1431` (Phase 1;
    `importShape` timing carry-through, an expanded preset set, and transition sounds remain
    Phase 2).
  - **Phase 2 capability B — expanded preset set (2026-06-26).** `PresetEffect` now covers
    eight presets: the entrance set adds `appear` and `wipe`, emphasis adds `spin`, and exit
    adds `flyOut` (alongside the Phase 1 `fadeIn`/`flyIn`/`grow`/`fadeOut`). Each is a verbatim
    PowerPoint template captured from `slide-animation-presets.pptx`; the writer reproduces all
    eight byte-for-byte (regression + schema fixtures). `flyOut` faithfully mirrors PowerPoint's
    exit-fly serialization (bare `ppt_x`/`ppt_y` run-time variables with no leading `#`, no
    `fill="hold"` on the motion node, and a trailing visibility `set`).
  - **Phase 2 capability C — transition sounds (2026-06-26).** `slide.transition.sound`
    (`TransitionSoundProps`) adds a `p:sndAc` to the transition: an embedded start sound
    (`{ data | path, name?, loop? }` → `p:stSnd`/`p:snd r:embed`, optionally `loop="1"`) or the
    stop-previous form (`{ stopPrevious: true }` → `p:endSnd`). A start sound registers an ECMA
    `audio` relationship, embeds the WAV as a `ppt/media/*` part, and adds a
    `wav=audio/x-wav` Default content type; identical sound bytes are deduped to a single part
    across slides. Read side (`pptxgenjs/read`): `slide.transition.sound` decodes the `sndAc`
    into a `TransitionSoundInfo` (`form`/`loop`/`embedRid`/`name`).
  - **`avContentType('wav')` now returns `audio/x-wav`** (was `audio/wav`), matching what
    PowerPoint authors for embedded audio; affects both `addMedia` WAV audio and transition
    sounds.
  - **Phase 2 capability A — `importShape` build-animation carry (2026-06-26).**
    `Presentation.importShape`/`importShapes` gained an opt-in `{ carryAnimation: true }`
    (`ImportShapeOptions`). By default a lifted shape still lands static (its slide-scoped build
    is dropped); when set, the shape's effect click-group(s) and `<p:bldP>` are copied into the
    destination `p:timing` (created from scratch when the host has none), their `spid` references
    remapped to the shape's new id and their `<p:cTn>` ids renumbered to stay collision-free, and
    appended after any existing build — the programmatic analogue of PowerPoint's
    copy/paste-with-animation merge.
  - **Whole-slide animation flatten — `slide.flattenAnimations()` (2026-06-26).** Removes the
    slide's `<p:timing>` block (the `<p:bldLst>` and effect tree), flattening the slide to its
    final static state with every shape shown at once. The whole-slide counterpart to the
    per-shape `pruneAnimationSpids()`. Gated like `hasAnimations`: a `<p:timing>` that is purely a
    media loop (no `<p:bldP>` / `presetID`) is left untouched so media playback survives. Marks
    the part dirty and returns `true` only when a timing block is removed; idempotent; never
    deletes shapes. Implements backlog `dn-flatten-slide-animations`.

- **`addChart(type, data, { metadata })` — custom chart-level metadata via a schema-valid
  extension:** pass a `Record<string, string>` of annotations (e.g. a source-data id, a
  generator tag, a semantic role) that should travel with the chart. They are emitted as the
  last child of the chart space (`c:chartSpace/c:extLst`) under a stable PptxGenJS vendor GUID
  (`{094A432E-1F6C-499B-95B8-B57DC9536949}`) in a foreign namespace
  (`http://pptxgenjs.com/schema/chart/metadata`), riding the `CT_Extension` lax `xsd:any`
  wildcard so PowerPoint **preserves** the data untouched and ignores it for rendering. This is
  the OOXML-valid form of the rejected upstream #894 `c:meta` injection, which proposed an
  invalid sibling element PowerPoint would strip/repair. It is a **validated** primitive, not a
  raw-XML escape hatch: non-string/empty keys and non-string values are dropped with a console
  warning (no silent coercion), keys and values are XML-escaped, and metadata that is absent or
  has no valid entries emits no `extLst` at all. Lands in `src/gen-charts.ts`
  (`genXmlChartMetadata`) with the type on `IChartOpts` in `src/core-interfaces.ts`. Implements
  backlog `chart-metadata-extlst`.

- **`importShape(target, source, i, { rescale })` — lift shapes across decks of different
  slide sizes:** by default `importShape`/`importShapes` still throws on a slide-size
  mismatch (now with a hint pointing at the option); set `rescale` to scale the lifted
  shape's geometry onto this deck's canvas instead. Mirrors `importSlide`'s option:
  `'fit'` (alias `true`) scales by `min(sx, sy)` and centers the slack, holding aspect
  ratio (matches "Ensure Fit"); `'stretch'` scales each axis independently (matches
  "Maximize"). Only **geometry** is rewritten — the shape/group/`graphicFrame` transform
  (`a:off`/`a:ext`) and any table grid (`a:gridCol@w`, `a:tr@h`); font sizes and line
  widths are left as authored. The transform is applied after the `preserve`-flatten pass
  (so a placeholder's just-baked inherited `a:xfrm` scales too), and explicit
  `left`/`top`/`width`/`height` overrides are applied last and win. Reuses the stateless
  rescale toolkit added for `importSlide`. Lands in `src/read/api/presentation.ts`.
  Implements the geometry-rescale gap of backlog `dn-importshape-v1-limits` (its
  timing-drop and best-effort-placeholder limits remain deferred).

- **`importShape(..., { theme: 'preserve' })` now lifts placeholders robustly:** a lifted
  placeholder is baked **self-contained** and **demoted** to a plain shape (its `p:ph`
  stripped) once everything it inherited is materialized, so it no longer re-resolves
  against the *host* deck's layout/master placeholder of the same `type`/`idx` (wrong
  inheritance, or a fallback when the host has none) and can no longer collide with the
  host slide's own placeholder of that type. Two new bakes back the demotion, alongside the
  existing geometry/colour/run-size passes: the placeholder-inherited **vertical anchor**
  (`a:bodyPr/@anchor`, so a centred title doesn't jump to top-anchored) and the inherited
  **list style** (per-level `a:lstStyle` paragraph defaults — indent, bullets, alignment,
  `a:defRPr` — resolved layout placeholder → master placeholder → master `p:txStyles`
  category, most-specific tier per level; explicit paragraph `a:pPr` on the slide's own runs
  still wins). Scoped to `flattenShape` (the `importShape` `preserve` path) — `flattenSlide`
  keeps placeholders as placeholders by design, and `restyle`/`copy` keep `p:ph` so the
  shape re-brands. Lands in `src/read/oxml/theme.ts`. Closes the best-effort-placeholder
  limit of backlog `dn-importshape-v1-limits`; only the animation/timing drop remains
  deferred (tracked with `dn-flatten-slide-animations`).

## [8.0.0](https://github.com/shbernal/PptxGenJS/releases/tag/v8.0.0) - 2026-06-25

### Added

- **`importSlide(source, i, { importNotes: true })` — carry the source slide's speaker
  notes across decks:** by default `importSlide` still drops the slide's `notesSlide`
  (the prior behaviour, so an import doesn't drag a notes master across); set `importNotes`
  to copy the notes onto the imported slide. The source `notesSlide` is copied under a
  fresh partname and wired to the new slide; its `slide` back-relationship is repointed at
  the imported slide (the source slide is **not** copied); any notes media travels along.
  Because a presentation may have **at most one** `notesMaster` (`CT_NotesMasterIdList`
  holds 0..1 `p:notesMasterId`), the imported notes **reuse this deck's existing notes
  master** when it has one — the source notes master and its theme are not copied, so the
  destination's notes styling wins — and only when the deck has none is the source notes
  master (plus its theme) copied and registered (`p:notesMasterIdLst`, after
  `p:sldMasterIdLst` in `CT_Presentation`). Works across all three `theme` modes. Lands in
  `src/read/api/presentation.ts` (`#carryNotes` / `#ensureNotesMaster`). Implements backlog
  `dn-importslide-v1-limits` gap #3 (the last one — the item is now fully implemented).

- **`importSlide(source, i, { rescale })` — import a slide across decks of different
  slide sizes:** by default `importSlide` still throws on a `p:sldSz` mismatch (now with
  a hint pointing at the option); set `rescale` to remap the imported geometry onto this
  deck's canvas instead. `'fit'` (alias `true`) scales by `min(sx, sy)` and centers the
  slack, preserving aspect ratio (circles stay circles, rotations hold), matching
  PowerPoint's "Ensure Fit"; `'stretch'` scales each axis independently to fill the
  canvas (distorts shapes), matching "Maximize". Only **geometry** is rewritten — every
  top-level shape/group/`graphicFrame` transform (`a:off` scale+translate, `a:ext` scale;
  groups are not recursed, so children remap via the untouched `chOff`/`chExt`) plus table
  grids (`a:gridCol@w` by `sx`, `a:tr@h` by `sy`). Font sizes and line widths are left as
  authored, so heavy down-scaling can leave text overflowing its (now smaller) box. In
  `copy` mode the imported layout and master shape trees are rescaled too (idempotently,
  so a master shared across repeated imports is scaled once), keeping inherited
  placeholder/background geometry aligned; `preserve`/`restyle` rebind to this deck's own
  master/layout (already the destination size) so only the slide is touched. Lands in
  `src/read/api/presentation.ts`. Implements backlog `dn-importslide-v1-limits` gap #2.

- **Embedded fonts — author-side `pptx.embedFont()` and import-carry
  `importSlide(..., { embedFonts: true })`:** embed font faces so a deck renders with
  them on machines that lack the font, mirroring PowerPoint's "Embed fonts in the file".
  Two independent entry points sharing one OOXML model (`src/embedded-fonts.ts`):
  - **Author-side** — `await pptx.embedFont({ path | data, typeface, style })` embeds a
    **whole** face (not glyph-subset) from a file path/URL or in-memory bytes
    (`Uint8Array` / `ArrayBuffer` / base64). Repeated calls with the same `typeface` and
    different `style` (`'regular'` | `'bold'` | `'italic'` | `'boldItalic'`) accumulate
    into one `p:embeddedFont` entry. The declared `typeface` MUST match the family name
    used in `fontFace`/run typefaces or PowerPoint won't bind it. At write time this emits
    `/ppt/fonts/fontN.fntdata` parts (raw bytes, STORE-compressed), a single
    `application/x-fontdata` content-type Default, one `font` relationship per face, and a
    `p:embeddedFontLst` at `CT_Presentation` index 7, and sets `embedTrueTypeFonts="1"` +
    `saveSubsetFonts="0"` on `p:presentation`. Font licensing (`OS/2.fsType` permission
    bits) is the caller's responsibility — not enforced. Lands in `src/pptxgen.ts`,
    `src/gen-xml.ts`, `src/core-interfaces.ts`.
  - **Import-carry** — `importSlide(source, i, { embedFonts: true })` brings the source
    deck's presentation-level embedded fonts across: copies the `.fntdata` parts (deduped
    via the per-source copy registry), adds the `fntdata` Default, rebuilds the font
    relationships, and merges entries into this deck's `p:embeddedFontLst` (de-duped by
    `typeface` + face slot, so repeated imports carry each face once). Default off — the
    deck is unchanged without the flag. Lands in `src/read/api/presentation.ts`.
  - **Append-carry** — `appendSlides(generator, …)` now carries the **generator's**
    author-side embedded fonts (`pptx.embedFont`) into the destination deck instead of
    silently dropping them: `extractSlides()` surfaces them on `ExtractedSlides`, and the
    append path writes the `.fntdata` parts, adds the `fntdata` Default + font
    relationships, and merges into the deck's `p:embeddedFontLst` (de-duped by `typeface`
    + face slot, so appending the same generator twice — or onto a template that already
    embeds the face — carries each face once). So a generator that calls `embedFont()`
    keeps its embedded fonts when grafted onto a `fromTemplate` deck. Shares one merge
    core with import-carry in `src/read/api/presentation.ts`.

  Note: when no fonts are embedded, output is unchanged — the historical inert
  `saveSubsetFonts="1"` and absent `embedTrueTypeFonts` are preserved. Implements backlog
  `dn-importslide-v1-limits` gap #1; supersedes the dismissed `upstream-pr-1302`.

- **`addTable(rows, { fitColumns: 'shrink' })` — shrink columns to fit the slide:**
  when a table's total column width exceeds the space between its `x` and the right
  slide margin, scale every column down by the same factor so the whole table fits.
  Applies to an explicit `colW` array (the common "too many columns" case) and to a `w`
  wider than the slide; it is **shrink-only** (never grows columns) with **no
  minimum-width floor** (a very high column count can still become thin). Opt-in and
  off by default — without the flag, explicit widths are emitted as-is and may run off
  the slide, so existing decks are unchanged. The scaling runs once in
  `addTableDefinition` after width resolution, so both the XML emitter and the
  measured-fit pass see the fitted grid. Lands in `src/core-interfaces.ts`
  (`TableProps.fitColumns`) and `src/gen-objects.ts` (`addTableDefinition`). Implements
  backlog `upstream-issue-1451`.

- **`addTable(rows, { headerRow })` — inline header-row styling:** style a table's
  first row distinctly from the body **without** first registering a custom style via
  `pptx.defineTableStyle({ firstRow })`. `headerRow` (a `TableCellProps`) is applied as
  direct per-cell formatting on row 0 — `fill`, `color`, `bold`, `align`, `border`, etc.
  Precedence matches how PowerPoint resolves styling (direct formatting overrides a style
  region): explicit per-cell `options` on a row-0 cell win over `headerRow`, which wins
  over a `tableStyle`'s `firstRow` region. Setting `headerRow` also implies `hasHeader:
  true` (emits `firstRow="1"` for the accessibility "table header" marker) unless
  `hasHeader` is explicitly `false`. The caller's `rows` array is not mutated. Additive
  (non-breaking); lands in `src/core-interfaces.ts` (`TableProps.headerRow`) and
  `src/gen-objects.ts` (`addTableDefinition`). Implements backlog `upstream-issue-1256`.

- **`defineSlideMaster({ textStyles })` — configurable per-level master text styles:**
  configure the shared slide master's `<p:txStyles>` instead of accepting the fixed
  Office defaults. `textStyles` (`MasterTextStyleProps`) carries `title` (single level),
  `body[]`, and `other[]` (each up to the nine list levels; index `0` is `lvl1`). Each
  level (`MasterTextStyleLevel`) sets `fontSize` (pt), `fontFace`, `color` (hex or theme
  slot), `bold`, `italic`, `align`, `marginLeft`/`indent` (**inches**), and `bullet`
  (`false` → `<a:buNone/>`, or `MasterBulletProps` for a character/auto-number bullet) —
  exactly the nested-bullet character, size, and colour control that was previously
  impossible (gitbrent/PptxGenJS#1360). Any unset field keeps that level's built-in
  default, and a deck that does not pass `textStyles` emits the **byte-identical** default
  master as before. Because a deck has a single shared master, `textStyles` is **deck-wide**:
  set across multiple `defineSlideMaster()` calls, the last value for each group
  (`title`/`body`/`other`) wins. Levels past nine warn and are dropped; an invalid
  `fontSize` warns and keeps the default. Lands in `src/gen-xml.ts` (`makeXmlMasterTxStyles`)
  and `src/pptxgen.ts` (`defineSlideMaster`). Implements backlog `upstream-issue-1360`.

- **`importSlide(src, i, { theme: 'restyle', remapLiterals: true })` — force-remap
  literals and copy table styles:** an opt-in flag that pushes a `restyle` re-brand
  past what symbolic theme references reach, for slides whose palette is partly
  hardcoded. Plain `restyle` can only recolour what is symbolic (a literal `a:srgbClr`
  has no theme reference, so it stays its authored RGB; a restyled table whose
  `@tableStyleId` is absent in the destination silently falls back). `remapLiterals`
  closes both gaps: (1) every literal `a:srgbClr` equal to a **source**-theme
  `clrScheme` slot is rewritten back to a symbolic `a:schemeClr` — routed through the
  source `clrMap`, transforms carried, a literal matching no slot left untouched — so
  it re-resolves against the destination theme; and (2) any `<a:tableStyleId>` the
  slide references is copied from the source `tableStyles.xml` into this deck's under
  the **same** id (idempotent; an id the deck already defines is left alone; a missing
  `tableStyles.xml` part is created and wired), and since the copied `<a:tblStyle>` is
  itself symbolic it re-brands to the destination theme. Off by default, preserving the
  prior byte-identical-literal guarantee; deliberately reinterprets authored literals
  as theme colours, so output wants visual QA. Lands in `src/read/oxml/theme.ts`
  (`remapLiteralColors`) and `src/read/api/presentation.ts`
  (`#copySourceTableStyles`/`#ensureTableStylesPart`). Verified against
  PowerPoint-authored `test/read/fixtures/multi-theme.pptx` slide 3. Implements backlog
  `dn-importslide-restyle-literals`.

- **`slide.addComment({ author, text, … })` — native PowerPoint review comments:**
  attach legacy (ISO/IEC 29500 §13) comments to a slide. The writer emits a per-slide
  `/ppt/comments/comment{N}.xml` (`<p:cmLst>` of `<p:cm authorId dt? idx><p:pos x y/>
  <p:text>…</p:text></p:cm>`, `pos` in EMU) plus the shared presentation-level
  `/ppt/commentAuthors.xml` (`<p:cmAuthorLst>` of `<p:cmAuthor id name initials lastIdx
  clrIdx/>`), and wires the `slide→comments`, `presentation→commentAuthors` relationships
  and both `[Content_Types].xml` Overrides. Authors are de-duplicated deck-wide by
  `name`+`initials`; each author's comments are numbered with a per-author 1-based `idx`
  (and `lastIdx`). Options: `author` (required), `text` (required), `initials` (defaults
  to letters derived from `author`), `x`/`y` marker position in inches (default `0.5`),
  and `date` (a `Date` or ISO-8601 string; omitted when absent). Missing `author`/`text`
  warns and skips rather than emitting a degenerate comment. Comment-free decks are
  byte-identical to before. Modern (PowerPoint 2021+, MS-PPTX 2.16) comments are not yet
  supported. Schema-validated against the ECMA-376 XSD via two fixtures in
  `test/schema.test.js`. Lands in `src/slide.ts`, `src/gen-objects.ts`, `src/gen-xml.ts`,
  `src/pptxgen.ts`, `src/core-interfaces.ts`. Implements backlog `upstream-pr-1447`.

- **`appendSlides` carries online (external-link) video, and `addMedia({type:'online'})`
  now emits the rel graph PowerPoint authors:** a source slide with an online video
  appends onto an existing deck instead of dropping the link (previously the appended
  slide kept the poster but left a **dangling** `<a:videoFile r:link>`). The generator
  side was also brought up to PowerPoint's shape — `addMedia({type:'online', link})`
  now writes **two external relationships** sharing the link Target (the ECMA
  `…/relationships/video` rel and the Microsoft-2007 `…/relationships/media` rel, both
  `TargetMode="External"`), the body references the second via `<p14:media r:link>`
  (inside the `p:nvPr` extLst) alongside `<a:videoFile r:link>` and the poster
  `<a:blip r:embed>`, and the `p:cNvPr` gains `<a:hlinkClick action="ppaction://media"/>`.
  There is **no** media binary part and **no** A/V content-type entry — only the poster
  image part is added. `extractSlides()` exposes these via a new `onlineMedia` descriptor
  array on each extracted slide (**breaking**: `ExtractedSlide` gains a required
  `onlineMedia` field). Verified against a PowerPoint-authored oracle fixture
  (`test/read/fixtures/online-video.pptx`). Lands in `src/gen-objects.ts`,
  `src/gen-xml.ts`, `src/pptxgen.ts`, `src/read/api/presentation.ts`. Implements backlog
  `dn-append-online-video`.

- **`Presentation.fromTemplate(input, options?)` — author a fresh deck on a real
  PowerPoint template (`pptxgenjs/read`):** open a `.pptx` *or* `.potx` template,
  get back an empty deck shell whose slide masters, layouts, and theme are kept
  **byte-identical**, then discover bindable layouts with `layouts()` and graft
  generator-produced slides on with `appendSlides()` — reusing the template's
  authored chrome verbatim instead of rebuilding it with `defineSlideMaster()`.
  Any sample slides the template carried are stripped to a master/layout/theme-only
  shell (the existing `removeSlide` path, which never prunes shared chrome). A
  `.potx` package's main-part content type (`…presentationml.template.main+xml`) is
  normalized to the editable `…presentationml.presentation.main+xml` by default so
  the saved output opens as a normal deck; pass `{ keepTemplateContentType: true }`
  to preserve the template type. `FromTemplateOptions` is re-exported from
  `pptxgenjs/read`. Verified against a PowerPoint-authored `.potx` oracle fixture
  (`test/read/fixtures/template.potx`). Lands in `src/read/api/presentation.ts`,
  `src/read.ts`. Implements backlog `dn-import-template-masters`.

- **`appendSlides` works on a template with zero slides:** `Presentation.appendSlides`
  (and `cloneSlide`) previously threw `presentation.xml has no p:sldIdLst to append
  a slide to` when the deck omitted `p:sldIdLst` — which real PowerPoint templates do
  when they carry no sample slides. It now creates `p:sldIdLst` in `CT_Presentation`
  document order (before `p:sldSz`) on demand, so authoring onto a freshly-loaded
  template shell works. Lands in `src/read/api/presentation.ts` (`#insertSlidePart`).

- **`appendSlides` carries embedded audio/video media:** a source slide produced
  with `addMedia({ type: 'audio' | 'video', … })` now appends onto an existing
  deck, reproducing the rel graph PowerPoint authors — one media part backing two
  relationships that share its Target (the ECMA `audio`/`video` rel and the
  Microsoft-2007 `media` rel referenced by `<p14:media>`) plus a separate preview
  image part — with the media part's content type registered as a `Default`
  extension entry (e.g. `video/mp4`, `audio/mpeg`), not a per-part `Override`.
  `extractSlides()` now exposes these via an `avMedia` descriptor array on each
  extracted slide (**breaking**: replaces the prior `hasAvMedia: boolean` flag),
  and `appendSlides` no longer throws on A/V slides. Verified against a
  PowerPoint-authored oracle fixture (`test/read/fixtures/av-media.pptx`).
  Lands in `src/pptxgen.ts`, `src/read/api/presentation.ts`, `src/gen-utils.ts`
  (`avContentType`). Implements backlog `dn-append-av-media`.

- **`Presentation.appendSlides(source, { layout })` — generate slides onto an
  existing deck (`pptxgenjs/read`):** open a real `.pptx`, keep its
  masters/layouts/theme (and every other untouched part) **byte-identical**, and
  append generator-produced slides bound to one of the deck's existing layouts —
  the hybrid "generate-onto-existing" path. `source` is any slide producer (a
  `PptxGenJS` instance); its slides are authored, serialized, and spliced in under
  fresh partnames with each slide's `slideLayout` relationship pointed at the
  chosen layout (relationship ids in the slide body are preserved, only their
  targets are repointed). Only `presentation.xml`, its `.rels`,
  `[Content_Types].xml`, and the new slide/media parts change. Companion additions:
  **`Presentation.layouts()`** returns the deck's layout gallery as addressable
  `LayoutHandle`s, and **`PptxGenJS.extractSlides()`** exposes the generator's
  per-slide artifacts (body XML + image media + hyperlinks + charts) without
  producing a package. **Charts** (chart XML + `.rels` + embedded workbook) and
  **internal slide-to-slide hyperlinks** (`slide:N` repointed at the Nth appended
  slide) are carried across; appendSlides injects in two passes so a forward link
  resolves. Limitations: an internal link to a source
  slide outside the appended batch throws; appended slides are concrete absolute-positioned content (no
  placeholder inheritance — `schemeClr` re-resolves against the destination theme);
  source/destination slide sizes must match; notes are not generated. Lands in
  `src/pptxgen.ts` (`extractSlides`), `src/read/api/presentation.ts` (`layouts`,
  `appendSlides`, reusing `#insertSlidePart`), `src/gen-charts.ts`
  (`buildEmbeddedWorksheet`/`buildChartRelsXml` split), and `src/read.ts` (type
  exports). Implements backlog `dn-append-onto-existing-deck`.
- **`FontMetrics.hasCodepoint(cp)` — cmap glyph coverage on the `measure` API:**
  the `measure` subpath's `FontMetrics` interface gains
  `hasCodepoint(cp: number): boolean`, reporting whether a face's cmap maps a code
  point to a real glyph (not `.notdef`). File-backed metrics
  (`parseFontMetrics`, `pptx.registerFontMetrics`) answer authoritatively via the
  parsed font; the unregistered-font heuristic has no cmap and reports every code
  point as covered. Lets a consumer flag source code points the replica face
  cannot render (e.g. U+2011 non-breaking hyphen, absent from Aptos) without
  shelling out to fontconfig. Lands in `src/font-metrics.ts`.
- **`slide.addGroup(children, options?)` — group slide objects (issue #307):**
  wraps child objects in a PowerPoint group (`<p:grpSp>`) so they become one
  selectable/movable group. Children are key-tagged descriptors reusing the slide-master
  `objects` shape (`{ text }`, `{ image }`, `{ shape }`, `{ rect }`, `{ roundRect }`,
  `{ line }`), plus `{ group: { children, options? } }` to **nest a group inside a group**
  to any depth. The child coordinate space is the identity transform (`chOff/chExt` ==
  `off/ext`) at every level, so children keep their slide-absolute `x/y/w/h` and grouping
  is visually a no-op. When `options.x/y/w/h` are omitted the group's bounds are
  auto-computed as the bounding box of its children (recursing into nested groups). Not yet
  supported (each skipped with a warning): charts, media, tables, and placeholders as group
  children; true child-space scaling (`chOff/chExt` != `off/ext`) is also deferred. Lands in
  `src/gen-objects.ts` (`addGroupDefinition`/`buildGroupObject`), `src/gen-xml.ts`
  (`<p:grpSp>` serialization, recursive `resolveObjBounds`), `src/slide.ts` (`addGroup`),
  with `GroupProps`/`GroupChildProps` (recursive) in `src/core-interfaces.ts`.
- **`pptx.tableLayout(rows, opts)` — computed table-cell geometry (issue #1169):**
  a layout-time accessor that returns each cell's `{ row, col, rowSpan, colSpan,
  xIn, yIn, wIn, hIn, heightExact }` (inches) plus overall `widthIn`/`heightIn`,
  without adding the table to a slide — so a consumer can place images or shapes
  precisely over cells. Takes the same `rows`/`opts` as `slide.addTable`. Column
  widths (cell `x`/`w`) are **exact**, derived from the same
  `resolveTableColWidthsEmu` the writer uses, and resolve colspan/rowspan via the
  same grid walk as the measured-fit pass (now shared, so placement cannot drift).
  Row heights are exact when pinned by `rowH` (array or scalar) or table `h`; an
  auto-height row is estimated with the same conservative (tall) text model as
  `measureText` and flagged `heightExact: false` (register the cell font via
  `registerFontMetrics` for an exact estimate). Single-slide only — `autoPage`
  paging is not modeled. New exported types: `TableCellLayout`, `TableLayoutResult`.

- **Read-model: gradient geometry, line arrowheads, outer shadow, and text-body
  properties.** The `pptxgenjs/read` shape proxies gain four getters that close
  the gaps a faithful slide replica kept hand-reading from raw XML: `gradientFill`
  (`GradientFill` — the stops **plus** the linear `angleDeg` / path shape that the
  bare `gradientStops` omits), `lineEnds` (`LineEnds` — connector `headEnd`/`tailEnd`
  arrowhead type/width/length), and `shadow` (`OuterShadow` — `a:outerShdw` colour
  (theme-resolved), `alpha`, `blurPt`, `offsetPt`, `angleDeg`). `TextFrame` gains
  `bodyProperties` (`BodyProperties` — `a:bodyPr` `vert`/`anchor`/`wrap` and
  explicitly-set insets in points). Angles use the OOXML clockwise-from-3-o'clock
  degree convention and distances are in points, so each round-trips to the
  matching write-side prop. New exported types: `GradientFill`, `LineEnd`,
  `LineEnds`, `OuterShadow`, `BodyProperties`.

- **Read-model: line dash, explicit no-line, and resolved (inherited) text
  anchor.** Three more getters that close gaps a faithful replica kept hand-reading
  from raw `slide.xml`. On the shape proxies: `lineDash` (`spPr/a:ln/a:prstDash/@val`
  — `'dash'`/`'lgDashDot'`/`'sysDot'`/…, or `null` when solid/unset) and `lineNoFill`
  (`true` when the shape sets an explicit `<a:ln><a:noFill/>`, distinguishing a
  deliberately border-less shape from one with an inherited line — both of which
  `resolvedLine` reports as `null`). On `TextFrame`: `resolvedAnchor`, the effective
  vertical anchor (`t`/`ctr`/`b`) resolving placeholder inheritance — the frame's own
  `a:bodyPr/@anchor` when set, else the anchor inherited from the layout → master
  placeholder `a:bodyPr` — where `bodyProperties.anchor` reports only the own
  attribute (so an inherited top-anchored title reads `null` there). New helper
  `placeholderInheritedAnchor` (sibling of `placeholderInheritedXfrm`) +
  `resolveInheritedAnchor` wrapper.

- **`widestLineIn` on `measureText()`/layout-time measurement:** `TextMeasurement`
  (from `pptx.measureText()` and the `pptxgenjs/measure` subpath) now reports the
  width in inches of the widest laid-out line, alongside `heightIn`/`lineCount`.
  With an unconstrained `wIn` it is the natural single-line width (deciding a box
  width / whether to wrap); constrained, it is the widest wrapped line (tightening
  a box to the actual text extent). It carries the same conservative
  `WIDTH_SAFETY` inflation as the wrap decision, so a box set to this width will
  not re-wrap. The underlying `measureLayout`/`LayoutResult` gains
  `widestLineWidthPt`. No drift: the value comes from the same wrap model the
  export-time bake uses. See `docs/measured-text-fit.md`.

- **Per-text-frame autofit mode and body insets on inspect elements:** each
  `PptxSlideElement` from `inspectPptx()`/`extractSlides()` now exposes `autofit`
  (`'none'` | `'normAutofit'` | `'spAutoFit'`, or `null` for elements without a
  text frame) and `bodyInsets` (`{ left, top, right, bottom }` in inches, with
  PowerPoint defaults applied when an `a:bodyPr` inset attribute is absent — 0.1in
  left/right, 0.05in top/bottom). Together these let a consumer distinguish a
  bounded text box (a genuine overflow candidate) from an auto-growing
  (`spAutoFit`) or text-shrinking (`normAutofit`) one, and compute the inner text
  box (`box` minus `bodyInsets`) for overflow detection. See
  `docs/reference/pptx-inspection.md`.

### Fixed

- **`addMedia` audio/video now embeds with PowerPoint-correct `Default` content
  types:** the write path derived the media `[Content_Types].xml` `Default` from
  the rel's internal `mtype/extn` string, so `mp3` audio emitted `audio/mp3`
  rather than the `audio/mpeg` PowerPoint authors. A new `avContentType(extn,
  mtype)` helper in `src/gen-utils.ts` resolves the standard mapping
  (`mp3 → audio/mpeg`, `m4a → audio/mp4`, `mov → video/quicktime`, …), used by
  both the generator's content-type builder (`src/gen-xml.ts`) and the
  `appendSlides` injection path so a fresh write and an append agree.
- **EMF/WMF images now embed with OOXML-correct content types
  (`image/x-emf` / `image/x-wmf`):** the write path previously built the image
  content type inline as `'image/' + extn`, so `emf`/`wmf` extensions emitted
  `image/emf` / `image/wmf` — values the library's own read side would not
  recognize as EMF/WMF (`IMAGE_EXTENSION_BY_CONTENT_TYPE` expects the `x-`
  forms). A new `imageContentType(extn)` helper in `src/gen-utils.ts` (the
  inverse of the read-side map) is now used at every image-rel `type:`
  assignment and duplicate-guard in `src/gen-objects.ts`. Two latent bugs are
  fixed along the way: the slide-background image push was emitting the literal
  string `"image"` as a content type, and JPEG images now correctly emit
  `image/jpeg` (previously `image/jpg`). File extensions and Target filenames
  are unchanged. **Downstream impact:** decks embedding EMF/WMF now open
  cleanly in stricter consumers than LibreOffice; consumers asserting the old
  `image/emf` / `image/wmf` / `image/jpg` content-type strings must update.

## [7.0.0](https://github.com/shbernal/PptxGenJS/releases/tag/v7.0.0) - 2026-06-21

This major release adds **measured text fit** — a calibrated, font-metrics-driven
layout engine that bakes autofit results at export time so overflowing text
self-corrects in headless renders and on plain file-open. It introduces the new
public methods `pptx.registerFontMetrics`, `pptx.measureText`, and
`pptx.overflowsBox`, plus the new `pptxgenjs/measure` subpath export. The ZIP
backend is also replaced (JSZip → fflate) for faster builds, reads, and writes.

Two behavior changes motivate the major version bump: (1) when font metrics are
registered, `fit:'shrink'`/`'resize'` now bake the computed result (font scale or
shape height) instead of emitting a bare autofit flag — with no metrics registered
the previous bare-flag behavior is unchanged; and (2) a registered-but-mismatched
named face now bakes a conservative heuristic fit rather than keeping the bare flag.
This release also expands the `pptxgenjs/read` subsystem (freeform custom geometry,
placeholder-inherited run size/typeface/colour, table-cell style accessors, binary
part access, DrawingML colour transforms, group-aware absolute frames) and adds
generic preset shapes in masters, native math (OMML), table placeholders, slide/part
removal, and external slide-master grafting. See `docs/measured-text-fit.md`.

### Added

- **Measured text fit for `fit:'shrink'` and `fit:'resize'` (`pptx.registerFontMetrics`):**
  the library can now compute and bake the autofit result so overflowing text
  self-corrects in headless renders (and on plain file-open) without a manual
  edit/resize. Register a face's font file once —
  `await pptx.registerFontMetrics('Aptos', '/path/Aptos.ttf')` (path/URL or raw
  `Uint8Array`/`ArrayBuffer`; pass `{ bold }`/`{ italic }` per variant) — and any
  `fit:'shrink'` or `fit:'resize'` text box in that face is measured at export time.
  The box's text is wrapped with the font's `hmtx` advances (raw, no kerning/GSUB —
  the conservative direction). For `'shrink'` the largest fitting `fontScale` is found
  on PowerPoint's 2.5% grid and `<a:normAutofit fontScale=…/>` is emitted; for
  `'resize'` the height the text needs is computed and baked into the shape's
  `a:ext/@cy` (with `a:off/@y` shifted per vertical anchor — 0 / half / full of the
  delta for `t` / `ctr` / `b`), leaving the `<a:spAutoFit/>` marker in place.
  Measurement uses `opentype.js` (new dependency, lazily imported, Node/web only). The
  model is calibrated against PowerPoint-authored fixtures (`autofit-calibration.json`)
  and is conservative — the `'shrink'` `fontScale` is ≤ the value PowerPoint bakes and
  the `'resize'` `cy` is ≥ both PowerPoint's and LibreOffice's rendered height — so
  text never overflows in PowerPoint or LibreOffice (regression:
  `autofit-calibration-oracle`). **Behavior change:** when metrics are registered,
  `fit:'shrink'`/`'resize'` now bake the result instead of emitting a bare flag. With
  no metrics registered the previous behavior is unchanged (bare flag); a box whose
  face lacks metrics keeps the bare flag and warns once.
  Bold/italic/charSpacing/line-spacing/space-before-after, multi-run paragraphs, hard
  breaks, and `wrap=none` are handled. New public method `PptxGenJS.registerFontMetrics`.
  (see `docs/measured-text-fit.md`)
- **Measured fit for table cells (`TableCellProps.fit:'shrink'`) + unregistered-font
  heuristic (see `docs/measured-text-fit.md`):**
  - `addTable` cells now accept `fit:'shrink'` (also cascades from a table-level
    `fit:'shrink'`). PowerPoint has **no** text-autofit for table cells (`a:tcPr`
    carries no autofit element and the app ignores `normAutofit` inside a cell — rows
    auto-grow instead), so a cell's shrink is honored by baking a **reduced literal
    font size** onto its runs (computed with the same calibrated wrap simulator + shrink
    solver), which both PowerPoint and LibreOffice render identically with no edit/resize.
    Only triggers for cells in a **fixed-height** row (`rowH`/table `h`) whose text
    overflows; an auto-height row is left alone (it simply grows). `'resize'` and the
    object form are intentionally ignored for cells (a row already auto-grows — the cell
    equivalent of `spAutoFit`).
  - When a deck has registered *some* metrics, a `fit:'shrink'`/`'resize'` box or cell
    whose **named** face has no exact metrics now falls back to a conservative
    average-advance heuristic and still bakes an approximate result (it warns once that
    the estimate was used) instead of degrading to the bare flag. A deck that registers
    no metrics at all is unaffected (measured fit stays off), and an unnamed
    (theme-default) face still stays unmeasurable (the face cannot be guessed).
    **Behavior change:** previously a registered-but-mismatched named face kept the bare
    flag; it now bakes a heuristic fit.
- **Layout-time text measurement (`pptx.measureText` / `pptx.overflowsBox` + the
  `pptxgenjs/measure` subpath):** the calibrated wrap model that powers the export-time
  autofit bake is now a public API, so a consumer can size its own geometry **before
  export** — grow a card to fit its text, reflow a grid, or detect overflow — using the
  *same* model the bake uses (a layout-time prediction never disagrees with the baked
  result). With metrics registered, `pptx.measureText(text, { wIn, fontSize, fontFace, … })`
  returns `{ heightIn, lineCount, measurable, fitsBox(hIn), shrinkScaleFor(hIn) }`
  (inches in, conservative/tall height out — matching the `'resize'` bake); a named face
  without exact metrics uses the conservative heuristic, an unnamed theme-default face is
  `measurable:false`. `pptx.overflowsBox(text, { wIn, hIn, … })` is a thin
  (slightly over-reporting) overflow check for a build-time warning. The pure primitives
  are also re-exported from a dedicated subpath so a consumer can measure standalone
  without a `PptxGenJS` instance: `measureLayout`/`measureHeightPt`/`solveShrink`/
  `solveResize`, the `FitParagraph`/`FitBox`/`MetricsResolver`/… types, the calibration
  constants, and `parseFontMetrics`/`getHeuristicFontMetrics`/`FontMetricsRegistry`
  (`opentype.js` stays lazily imported). New public methods `PptxGenJS.measureText` /
  `PptxGenJS.overflowsBox`, types `MeasureTextOptions`/`TextMeasurement`/`OverflowBoxOptions`,
  and the `@shbernal/pptxgenjs/measure` entry point. (see `docs/measured-text-fit.md`)
- **Freeform custom-geometry reads (`Shape.customGeometry`):** the `pptxgenjs/read`
  model now exposes a shape's `spPr/a:custGeom/a:pathLst` path geometry, or `null`
  when the shape uses preset geometry / none (the freeform counterpart of
  `Shape.presetGeometry`). It returns a faithful `CustomGeometry { paths:
  CustomGeometryPath[] }`: each `a:path` keeps its own path-unit viewport (`w`/`h`)
  and `fill`/`stroke` (schema defaults `norm`/`true` applied), plus an ordered
  `GeometryCommand[]` whose verbs — `moveTo`/`lnTo`/`cubicBezTo`/`quadBezTo`/`arcTo`/
  `close` — mirror the write-side `GeometryPoint` DSL so a consumer maps one-to-one
  (coordinates are raw path-unit integers in the path's `0..w`/`0..h` space; `arcTo`
  angles are exposed in degrees). The multi-path array is chosen over flattening to
  the single-path write DSL because `a:pathLst` is repeatable with independent
  per-path `fill`/`stroke`. Previously a consumer replicating a native freeform glyph
  had to reverse-engineer each `a:path` with a one-off extractor. New types
  `CustomGeometry`, `CustomGeometryPath`, `GeometryCommand` are exported from
  `pptxgenjs/read`. Pinned against `custgeom.pptx` (PowerPoint-authored
  freeform-lines / freeform-cubic / freeform-hole); note PowerPoint's own Merge
  Shapes emits a hole as a single `a:path` with two `moveTo`…`close` contours, so
  `paths.length` is 1 for PowerPoint-authored freeforms.
- **`readPptxBinaryPart(pptxPackage, path)` on the `pptxgenjs/inspect` package-access
  surface:** the binary sibling of `readPptxTextPart`, returning a part's raw bytes as a
  `Uint8Array` (or `null` when absent) instead of UTF-8 decoding them. The `PptxPackage.file()`
  accessor now also accepts `async('uint8array')` alongside `async('string')`. This lets a
  consumer pull embedded media (SVG/PNG/EMF blobs, fonts) out of a `.pptx` without a second
  zip library — previously the inspect surface only exposed text parts, so callers extracting
  media bytes had to keep their own JSZip dependency. Used by a consumer's icon extractor
  to drop JSZip entirely.
- **`pptxgenjs/read` resolves placeholder-inherited run size + typeface
  (`Run.resolvedSizePt`/`Run.resolvedFontFace`):** the size/face sibling of the
  existing `Run.resolvedColor`. When a placeholder run sets no own `@sz`/`a:latin`,
  these getters walk the same inheritance chain the colour resolver does —
  paragraph `a:defRPr` → slide `a:lstStyle` → layout/master placeholder `a:lstStyle`
  → master `p:txStyles` — and return the inherited value as a literal. `resolvedFontFace`
  additionally resolves a `+mj-*`/`+mn-*` major/minor theme-font token (whether on the
  run itself or reached through the chain) to its concrete face via the theme
  `fontScheme`. The run's own `@sz`/`a:latin` still wins when set (`resolvedFontFace`
  resolves a token there too). Previously a consumer transcribing a placeholder
  title/eyebrow had to eyeball the point size and assume the house typeface, since
  neither was emitted by the read model. `Slide.themeContext()` now also carries the
  theme `fontScheme`. Pinned against `multi-theme.pptx` slide 2 (inherited-title → 42pt
  / Century Gothic via titleStyle + `+mj-lt`; explicit-body → 20pt / Century Gothic).
- **Table-cell style reads (`TableCell.resolvedFill`/`fillSchemeColor`/`verticalText`/`anchor`/`marginsEmu`):**
  the `pptxgenjs/read` model now profiles table cells beyond their text. `resolvedFill`
  resolves the cell's `a:tcPr/a:solidFill` against the slide theme to a literal hex
  (with `effectiveHex` after `lumMod`/`lumOff`/… transforms — the same resolver as
  `Shape.resolvedFill`); `fillSchemeColor` reports the raw token; `verticalText`
  (`@vert`, e.g. `vert270`), `anchor` (`@anchor`), and `marginsEmu` (`@marL`/`@marR`/
  `@marT`/`@marB`) expose the cell's layout. This lets a consumer reconstruct a table's
  per-cell appearance (fills, vertical labels, insets) without hand-parsing `a:tbl` XML.
- **Generic preset shapes in slide masters (`{ shape: { type, options } }`):** a
  `defineSlideMaster({ objects })` entry can now be any preset shape addressed by
  `SHAPE_NAME`, e.g. `{ shape: { type: 'ellipse', options: { x, y, w, h, fill } } }`.
  This generalizes the existing hard-coded `line`/`rect`/`roundRect` master-object
  shortcuts (which remain) to every preset the `addShape()` serializer already
  supports (ellipse, triangle, chevron, …), so masters are no longer limited to
  rectangles and lines for decorative geometry (gitbrent/PptxGenJS#776).
- **Native math equations (raw OMML) in text (`TextProps.math`):** a text item can
  now carry a native, editable PowerPoint equation. `addText([{ math: '<raw OMML>' }])`
  emits a display-math paragraph (`<a14:m><m:oMathPara><m:oMath>…`) and wraps the
  equation shape in `<mc:AlternateContent><mc:Choice Requires="a14">` exactly as
  PowerPoint authors it, so the package validates and opens clean. `math` accepts
  inner OMML, a full `<m:oMath>`, or a full `<m:oMathPara>` (the `m`/`a14` namespaces
  are supplied by the wrapper). Raw OMML is the first deliverable; LaTeX/MathML→OMML
  conversion and an `mc:Fallback` raster are future work. Pinned against the
  PowerPoint-authored `math-omml.pptx` oracle (gitbrent/PptxGenJS#1456).
- **Table placeholders (`TableProps.placeholder`):** a table can bind to a
  layout/master content placeholder by name. The table `<p:graphicFrame>` then emits
  that placeholder's `<p:ph>` on its `<p:nvPr>`, and the table inherits the
  placeholder's position/size for any omitted `x`/`y`/`w`/`h` — mirroring the
  existing image (#1258) and text (#640) placeholder inheritance. Pinned against the
  `table-placeholder.pptx` oracle (gitbrent/PptxGenJS#1151).
- **Remove slides / parts (`Presentation.removeSlide`, `OpcPackage.removePart`):**
  the `pptxgenjs/read` model can now delete content. `removeSlide(index)` drops
  the `p:sldId` entry, the presentation→slide relationship, the slide part and its
  `.rels`, and recursively prunes any part the slide *privately* owned (its notes
  slide, slide-only media, charts/embeddings) that no remaining part references —
  while never pruning shared chrome (layout/master/theme), so the deck stays
  renderable and removing every slide leaves a valid master/layout-only template
  shell. The low-level `OpcPackage.removePart(partName)` deletes one part and
  unregisters its `Override` content type (supporting primitives:
  `Relationships.remove(id)`, `ContentTypes.removeOverride(partName)`). Untouched
  parts stay byte-identical; the shell is schema-valid
  (`test/read/remove-slide.test.js`, 6 cases). Motivated downstream by stripping a
  brand template to a master-only graft source (see `importSlideMasters`).
- **Graft external slide masters into a deck (`Presentation.importSlideMasters`):**
  a new `pptxgenjs/read` method copies slide master(s) from another open package
  together with their **whole** layout family and attaches them to no slide, so a
  brand template's layouts land in a generated deck's layout gallery
  (PowerPoint's *New Slide* / *Layout* picker) without touching existing slides.
  It complements `importSlide`, which only brings a master across as a slide's
  dependency and prunes it to the one used layout. Each grafted master is
  registered in `p:sldMasterIdLst`, its `p:sldLayoutIdLst` is rebuilt to exactly
  the copied layouts, and the connected theme/media/tag parts come across under
  fresh partnames (re-calls are idempotent via the copy registry). `options`:
  `masters`/`layouts` predicates narrow what is grafted (default: all);
  `requireEqualSize` (default `true`) guards against mis-scaled layouts. Reuses
  the existing cross-package copy engine; untouched parts stay byte-identical and
  the result is schema-valid (`test/read/import-slide-masters.test.js`, 9 cases).
  Brand-agnostic: the caller supplies the source `.pptx` (a consumer points it
  at its own brand template to ship a layout gallery).
- **Read model applies DrawingML colour transforms (`effectiveHex`):** the
  `pptxgenjs/read` colour resolver now computes the colour a renderer actually
  paints, not just the base token. `ResolvedColor` (from `Shape.resolvedFill` /
  `Shape.resolvedLine`, `Run.resolvedColor`, and each `Shape.gradientStops`
  entry) gains an `effectiveHex` field — the base `hex` with its ordered child
  transforms (`lumMod`/`lumOff`/`shade`/`tint`/`satMod`/…) applied — plus an
  optional `alpha` (0–1) when an `alpha*` transform sets opacity. The base `hex`
  and the raw `transforms` list are unchanged, so the `theme: 'preserve'`
  flatten path still re-emits transforms verbatim (byte-for-byte identical
  output). A new pure helper is exported for direct use:
  `applyColorTransforms(baseHex, transforms): EffectiveColor`. Additive, no
  signature breaks. Verified against an oracle table of PowerPoint/LibreOffice
  source→effective mappings (`test/read/color-transform.test.js`, 17 cases).
- **Group-aware absolute shape frames:** `Shape.absoluteFrame` now composes
  enclosing group rotation and flips, including scaled rotated groups, children
  with their own rotation/flip, and nested rotated groups. The returned frame
  adds effective `rotation`, `flipH`, and `flipV` fields; `left`/`top`/`width`/
  `height` remain the PowerPoint-style unrotated placement box in slide EMU.
  Verified against an expanded desktop PowerPoint-authored
  `group-transform.pptx` fixture whose second slide is PowerPoint's own
  ungrouped ground truth.

### Performance

- **fflate ZIP backend (replaces JSZip):** the whole library now builds and reads `.pptx`
  archives with [fflate](https://github.com/101arrowz/fflate) instead of JSZip, and the
  `jszip` runtime dependency is dropped. fflate's DEFLATE is several times faster than
  JSZip's pure-JS pako, so all output speeds up, not just media-heavy decks. All backend
  contact is isolated in `src/zip.ts`: a `ZipWriter` seam on the write path and a `readZip`
  reader on the read path. The public output types
  (`nodebuffer`/`arraybuffer`/`blob`/`base64`/`binarystring`/`uint8array`/`STREAM`) and the
  `compression` option are unchanged, and `OpcPackage.load`/`Presentation.load`/`inspect`
  accept the same inputs as before (`string`/`number[]`/`Uint8Array`/`ArrayBuffer`/`Blob`),
  normalized to `Uint8Array` for fflate. Phase 1 (write) covered `write()`/`writeFile()`/
  `stream()` and chart generation; phase 2 (read) covered `pptxgenjs/read` and `inspect`.
  Two incidental wins: archives no longer carry empty directory entries (fflate keys on full
  paths), and the embedded chart XLSX is now stored as raw bytes instead of a base64
  round-trip. The #1006 per-entry STORE-of-media behaviour is preserved, and archive bytes
  are reproducible across runs (entries carry a fixed timestamp). `jszip` is retained as a
  dev dependency only, where the round-trip tests use it as an independent zip oracle
  (backlog `dn-fflate-zip-backend`). The seam currently finalizes synchronously
  (`zipSync`/`unzipSync`); moving it to fflate's worker-parallel async `zip()`/`unzip()`
  for an additional large-deck speedup is deferred behind the same `src/zip.ts` boundary
  (backlog `dn-fflate-async-zip`).
- **Skip DEFLATE on already-compressed media (gitbrent/PptxGenJS#1006):** image and
  video parts (`jpg`/`jpeg`/`png`/`gif`/`webp`/`heic`/`heif`/`avif`/`mp4`/`m4v`/`mov`/
  `avi`/`mpg`/`mpeg`/`wmv`/`webm`/`mkv`/`mp3`/`m4a`/`aac`/`ogg`/`oga`) are now written
  to the package with per-entry `STORE` instead of `DEFLATE`. Their bytes are already
  entropy-coded, so DEFLATE-ing them burned CPU in JSZip's `generateAsync` for a
  negligible size gain — the dominant cost when exporting large, media-heavy decks.
  XML parts still DEFLATE, and `compression: false` still stores everything. Formats
  that genuinely benefit from DEFLATE (`bmp`/`wav`/`tiff`/`emf`/`wmf`/`svg`) are
  excluded and keep inheriting the global compression. No change to OOXML validity or
  output that PowerPoint opens.

### Fixed

- **Auto-width table columns collapsed to a sliver:** an `addTable` sized with `w`
  (or with neither `w` nor `colW`, the default full-slide width) emitted
  `<a:gridCol w="…">` values computed by dividing the **raw inches** width and writing
  the result as **EMU** — e.g. `w: 9` over 3 columns produced `gridCol w="3"` (≈0 in),
  so PowerPoint/LibreOffice drew near-zero-width columns and wrapped every cell to one
  glyph per line. The grid now divides the table's resolved EMU width. Explicit `colW`
  arrays were already correct and are unchanged. Column-width resolution is now a single
  shared helper (`resolveTableColWidthsEmu`) used by both the table emitter and the
  measured-fit pass. (regression: `table-colwidth-distribution`)
- **`addSection()` ignores duplicate and invalid sections (gitbrent/PptxGenJS#1152):**
  `addSection({ title })` now skips (with a `console.warn`) any title that already
  exists, instead of appending a second section with the same name — duplicate titles
  silently broke section-by-title lookups (`addSlide({ sectionTitle })` and autoPage
  continuation resolved to the first match). The pre-existing missing-argument and
  missing-title warnings now also early-return rather than falling through and pushing
  a section with an `undefined` title. Pinned by `test/regression/add-section-duplicate.test.js`.
- **Master/layout placeholder body properties (margin/valign) (gitbrent/PptxGenJS#1247,
  #1208):** a placeholder authored on a slide master/layout (via `defineSlideMaster`)
  with `margin` and/or `valign` now emits those in its `<a:bodyPr>` (insets + `anchor`)
  instead of degrading to the default, so a slide inserted from that layout inherits
  them. `genXmlBodyProperties` previously applied configured body properties only to
  ordinary text objects, not placeholder objects. Pinned against the
  `layout-placeholder-bodypr.pptx` oracle.
- **`pptxgenjs/read` `Run.resolvedColor` resolves placeholder-inherited colour:** a
  run inside a placeholder that sets no colour of its own now resolves the colour it
  inherits through the paragraph `a:defRPr` → slide `a:lstStyle` → placeholder
  layout/master `a:lstStyle` → master `p:txStyles` chain, instead of returning
  `null`. Pinned against `multi-theme.pptx` slide 2.
- **Master slide numbers no longer disappear on slides inserted in PowerPoint
  (gitbrent/PptxGenJS#1159):** when `defineSlideMaster({ slideNumber })` defined a
  slide-number placeholder, the master still emitted `<p:hf sldNum="0" .../>`.
  Because `CT_HeaderFooter/@sldNum` defaults to `true` (ECMA-376), the explicit
  `"0"` disabled the slide-number field for any slide PowerPoint inserts/inherits
  from that master — so the number rendered on PptxGenJS-generated slides but
  vanished on newly inserted ones. The master now omits `@sldNum` when a slide
  number is defined (letting it default to `true`); masters without a slide number
  still emit `sldNum="0"`. Covered by
  `test/regression/master-slide-number-hf.test.js`.
- **Auto-paged tables: a `rowH` *array* now follows its original rows across
  pages (gitbrent/PptxGenJS#1145):** `rowH` as an array is keyed by the original
  row index, but auto-paging passed the whole array to every overflow slide where
  it was re-applied by *physical* row index (which restarts per slide). A tall
  first row therefore reappeared as the first row of every overflow slide, and
  with `autoPageRepeatHeader` the body rows inherited the wrong heights. The
  auto-pager now resolves a per-slide height list aligned 1:1 with each generated
  slide's rows (repeated header rows keep their configured height; rows with no
  configured height auto-size), so a configured height stays with its source row.
  Single-number `rowH` is unchanged (it already applied uniformly). Covered by
  `test/regression/table-autopage-rowh-array.test.js`.
- **Read model: `Shape.resolvedFill`/`resolvedLine` now follow a `p:style`
  `fillRef`/`lnRef` (theme style matrix):** a shape whose fill or line comes only
  from its `p:style` style-matrix reference — no explicit `spPr` fill/line — read
  back as `null` from `resolvedFill`/`resolvedLine`, even though PowerPoint renders
  it with the referenced theme colour. The getters now fall back to the indexed
  theme `fmtScheme` entry with its `phClr` substituted by the ref colour (carrying
  the ref's colour transforms), matching how the `importSlide({ theme: 'preserve' })`
  flatten path already bakes it — the shared logic is factored into
  `styleRefFill`/`styleRefLine`, and `Slide.themeContext()` now also carries the
  theme `fmtScheme`. An explicit `spPr` fill/line still wins. Pinned against a real
  PowerPoint-authored fixture (`test/read/fixtures/multi-theme.pptx`). The
  placeholder-inherited *run* colour leg of the same resolver remains deferred
  (see `docs/backlog.yml` `dn-readmodel-style-followups`).
- **Images targeting a placeholder now inherit the placeholder's geometry (#1258):**
  `addImage({ placeholder: 'name' })` referencing a picture placeholder defined on a
  slide master/layout previously ignored the placeholder's position/size — the image
  collapsed to its natural pixel size (or the 1in fallback) unless explicit `w`/`h`
  were supplied. `addImageDefinition` now fills any of `x`/`y`/`w`/`h` the caller
  omits from the matching layout placeholder (mirroring the existing text-placeholder
  inheritance, #640); explicit `opt` values still win. Schema fixture added asserting
  the slide picture's `<a:ext>` matches the placeholder geometry.
- **`textDirection` now serializes to `<a:bodyPr vert="…">` on text boxes:** the
  documented `textDirection` option (`'horz' | 'vert' | 'vert270' | 'wordArtVert'`)
  was typed and documented but never emitted for text boxes — only the
  undocumented `vert` alias was read — so `textDirection: 'vert270'` was silently
  dropped and the text rendered horizontal. (Tables already honored
  `textDirection`.) Text-box body properties now set the `vert` attribute from
  `textDirection`, falling back to `vert`, which is retained as a legacy escape
  hatch for the full `ST_TextVerticalType` range (`eaVert`, `mongolianVert`,
  `wordArtVertRtl`) not listed by `textDirection`'s type. Schema fixture added.
- **Hyperlinks now inherit the theme hyperlink color (#1165):** a text
  hyperlink created without an explicit `color` no longer renders in the default
  black body color. Previously every run's color was defaulted to `000000`,
  which made the hyperlink emit an explicit `<a:solidFill>` plus an
  `ahyp:hlinkClr val="tx"` override — pinning the link to black and suppressing
  the theme `hlink` color (and the followed-link `folHlink` color after a click).
  Such runs now emit a bare `<a:hlinkClick/>` with no fill, so PowerPoint applies
  the theme hyperlink and visited colors automatically. Hyperlinks that set an
  explicit `color` keep their existing behavior (explicit fill + `hlinkClr="tx"`).

## [6.0.0](https://github.com/shbernal/PptxGenJS/releases/tag/v6.0.0) - 2026-06-14

This major release introduces the new `pptxgenjs/read` subsystem — a separate,
lossless read/edit/round-trip layer for existing decks — alongside radial
gradient fills and the image-in-shape composition. The new public `./read`
subpath export and its substantial API surface motivate the major version bump.

### Added

- **New `pptxgenjs/read` subsystem (`@shbernal/pptxgenjs/read`):** open an
  existing `.pptx`, navigate and edit it, and save it back losslessly (untouched
  parts stay byte-for-byte identical). It keeps the package's own XML as the
  source of truth, unlike the one-way/lossy generator and inspector subpaths.
  - OPC layer (`OpcPackage`): load, enumerate parts, content types, and
    relationships, writable parts, and a lossless `save()`.
  - Navigable read model: `Presentation → slides → shapes → text frame →
    paragraphs → runs`, including tables, charts (read-only), connectors, and
    nested groups.
  - Typed edits over the live DOM: run text and character formatting, shape
    position/size, shape fill/line colour, `Slide.hidden`, `Picture.setImage`
    to swap a picture's image, and `Picture`/`Slide.addPicture`. Setting a
    property mutates the DOM in place and marks only the affected part(s) dirty.
  - Structural edits: add/remove shapes, add pictures, edit table cell text, and
    slide cloning.
  - Cross-package composition: `Presentation.importSlide` (with
    `theme: 'preserve' | 'restyle'`, `carryMasterGraphics`, placeholder geometry
    and run-size baking, and placeholder-inherited run-colour preservation) and
    `importShape`/`importShapes` for cross-slide shape composition.
  - Loader hardening: PowerPoint `[trash]` parts are dropped on load.
  - Docs: `docs/reference/pptx-read.md`.
- **Radial gradient fills:** `RadialGradientFillProps` (`kind: 'radial'`, with
  optional `center` and `rotateWithShape`) joins the `GradientFillProps` union.
  It serializes as `<a:gradFill>` with `<a:path path="circle">` and a
  `<a:fillToRect>` focus derived from `center`, while the linear path is
  unchanged.
- Documented and tested "image embedded in a shape": `addImage({ points })` clips
  a picture to a freeform `custGeom` path (or `shape`/`rounding` for a preset),
  and pairing it with `sizing: { type: 'cover' }` fills the clip with an
  aspect-correct center-cropped source — the picture-placeholder form (`<p:pic>`
  with a clip in `<p:spPr>` and a source crop in `<p:blipFill>`). New regression
  + schema fixtures cover the `points` + `sizing` composition (incl. an `arcTo`
  half-disc clip) and the correct `blipFill`-before-`spPr` order. New docs
  (`docs/image-in-shape.md`) and a demo slide.

### Changed

- Image `sizing` (`cover`/`contain`/`crop`) now emits an explicit `<a:fillRect/>`
  inside `<a:stretch>` (the canonical form PowerPoint authors, ECMA-376
  §L.4.8.4.3) instead of an empty self-closing `<a:stretch/>`. Semantically
  identical (an absent `fillRect` already defaults to the full shape bounds), but
  it removes any rendering ambiguity when a source crop is composed with a
  `custGeom` clip.

## [5.4.0](https://github.com/shbernal/PptxGenJS/releases/tag/v5.4.0) - 2026-06-13

### Added

- `slide.addConnector({ type, x1, y1, x2, y2, ...line })` emits a real
  PowerPoint connector shape (`<p:cxnSp>`) with straight/elbow/curved preset
  geometries, min-corner box plus `flipH`/`flipV` derived from the endpoints, and
  line styling/arrowheads (upstream #1059).
- `ThemeProps.colorScheme` to configure a presentation's theme color scheme
  (upstream #1243).
- Table `autoPagePlaceholder` option to carry placeholders onto auto-paged
  overflow slides (upstream #1136).
- Image `line?: ShapeLineProps` for a picture border outline, emitted as `<a:ln>`
  in the picture `<p:spPr>` (reuses the shape-outline vocabulary; pairs with the
  existing `shadow`) (upstream #986).
- Pie/doughnut data-label leader-line styling via `leaderLineColor` /
  `leaderLineSize`, emitting `<c:leaderLines>` only when leader lines are enabled
  and styled (otherwise PowerPoint's automatic color is kept) (upstream #1376).
- `bullet.fontFace` (emits `<a:buFont/>`) and `bullet.size` (percent 25–400,
  mapped to `<a:buSzPct/>`, warns and falls back to 100% when out of range) for
  custom symbol/numbered bullet glyphs (upstream #800, #743).
- SVG source images for picture bullets (`bullet.image` with a `.svg` path or
  `image/svg+xml` data). SVG bullets now embed a PNG preview plus the SVG using the
  same dual-rel handling as `addImage()`: `<a:buBlip>` references the PNG preview via
  `<a:blip r:embed>` and the SVG via the `<asvg:svgBlip>` extension. Raster picture
  bullets are unchanged (follow-up to upstream #898).
- Actionable media-load errors with an opt-in placeholder fallback (upstream
  #1310).

### Changed

- **BREAKING:** exported `.pptx` packages are now DEFLATE-compressed by default
  on every export path (previously STORE, and the typed-output `write()` branch
  ignored the compression option entirely), producing packages comparable in size
  to a deck re-saved by PowerPoint. Pass `compression: false` to restore the old
  uncompressed STORE behavior (upstream #1268).

### Fixed

- `sizing: 'cover' | 'contain'` is now aspect-correct for SVG images:
  `getImageSizeFromBase64` reads an SVG's intrinsic size from the root `<svg>`
  (absolute `width`/`height`, else `viewBox`), so the letterbox/crop is computed
  from the real aspect ratio instead of stretching the SVG to fill the box.
- `defineSlideMaster` now passes rich-text arrays (`TextProps[]`) through master
  text objects unchanged, instead of wrapping them so the runs were lost
  (upstream #962).
- Table `autoPage` no longer crashes with `addTable: Array expected` when an
  explicit `h` plus `y`/margins leaves no usable vertical height: a non-positive
  usable height is clamped to the slide height (warning once) and empty overflow
  pages are no longer emitted.
- Tables honor `data-pptx-width` and no longer compute `NaN` column widths for
  hidden tables (upstream #1157).

## [5.3.0](https://github.com/shbernal/PptxGenJS/releases/tag/v5.3.0) - 2026-06-11

### Added

- `addImage` infers an image's natural size when `w`/`h` are omitted. For base64
  `data` images the intrinsic pixel size is read synchronously from the header
  (PNG/JPEG/GIF/BMP/WebP) and applied at 96 DPI; when only one of `w`/`h` is
  given, the other is derived from the natural aspect ratio. Previously a
  dimensionless image collapsed to a 1in square. `path` and SVG images (not
  synchronously measurable) keep the 1in fallback.
- Explicit coordinate unit suffixes on any `Coord` (`x`/`y`/`w`/`h`): `"<n>in"`
  (inches), `"<n>pt"` (points), and `"<n>emu"` (raw EMU) — alongside the existing
  bare number (inches) and `"<n>%"`. Example: `{ x: '72pt', w: '914400emu' }`.
- Exported branded `Emu` type and `coordToEmu` / `percentToEmu` converters from
  the units module (joining the existing `inchesToEmu` / `pointsToEmu` /
  `emuToInches` helpers).
- Run-level text shadow: a `shadow` (or `glow`) set on a text run now emits an
  `<a:effectLst>` inside `<a:rPr>`, so text in table cells — which have no shape
  `spPr` — can finally carry a shadow (upstream #1011).
- `OptsChartData.customLabels?: string[]` for per-data-point data label text
  overrides on BAR/LINE/AREA/RADAR and PIE/DOUGHNUT charts; empty entries fall
  back to chart-level settings (upstream #1337).
- `OptsChartData.pointStyles?: ChartDataPointStyle[]` for typed per-data-point
  border/fill styling (`{ border?: BorderProps; fill?: HexColor }`), index-
  aligned with `values[]`, on BAR/BAR3D/LINE/AREA/SCATTER/PIE/DOUGHNUT
  (upstream #1343).
- Text-box columns: `columns` (1–16) and `columnSpacing` (points) on
  `TextPropsOptions`, emitting `numCol`/`spcCol` on `<a:bodyPr>` (upstream #1320).
- Line `cap` (`'flat'|'round'|'square'`) on `ShapeLineProps` and `BorderProps`,
  emitted on `<a:ln>` for shapes, table cell borders, and charts via a shared
  `LineCap` type (upstream #782).
- `objectLock` (`ObjectLockProps`) on shapes, text boxes, images, media, and
  tables, serializing DrawingML `a:spLocks`/`a:picLocks`/`a:graphicFrameLocks`
  (noGrp, noMove, noResize, noRot, noCrop, …) (upstream #438).
- `shapeAdjust` ({ name, value }, single or array) on `ShapeProps` and
  `ImageBaseProps`, emitting preset-shape adjustment guides in `<a:avLst>`
  (upstream #1300).
- Chart title `titleItalic` / `titleUnderline` props, mirroring `titleBold`
  (upstream #1188).
- Partial chart-title manual layout: `titlePos` now accepts a partial
  `{ x?, y? }`, applying a manual offset on one axis while leaving the other on
  automatic layout (upstream #1363).
- Shrink-autofit tuning: `fit` accepts `{ type: 'shrink', fontScale?,
  lnSpcReduction? }` (percent 0–100) emitted on `<a:normAutofit>`; bare
  `fit: 'shrink'` is unchanged (upstream #1199).
- `barSeriesLine` on bar charts (`true` or an `OptsChartGridLine` object) emits
  `<c:serLines>` for stacked bars (upstream #1329).
- `showBubbleSize` option for bubble-chart data labels (upstream #744).

### Changed

- **Behavior change:** A bare-number coordinate is now **always inches**. The library no
  longer guesses units by magnitude — previously a number `>= 100` was silently
  treated as raw EMU (and `inch2Emu`/coordinate parsing carried a matching
  `> 100` passthrough), which mis-rendered any legitimately large value and made
  values near the threshold ambiguous.
  - *Migration:* if you were passing raw EMU as a large number (e.g. `914400`),
    pass it explicitly as a string instead (`'914400emu'`), or convert with the
    `emuToInches` helper. Bare numbers, `'%'`, and the new unit suffixes need no
    change.
  - Non-finite coordinates now throw with a descriptive message instead of
    collapsing the object to zero size; an implausibly large bare number (> 1000
    inches) is interpreted as inches but warns, pointing at the `'<n>emu'` form.
  - Internally, user coordinates are resolved to EMU exactly once at the
    emission boundary (no in-place pre-conversion / double-parse), and resolved
    values carry a branded `Emu` type so they cannot be silently re-converted.
- **Behavior change:** Removed the invalid `LINE_CALLOUT_4*` shape presets
  (`borderCallout4`, `accentCallout3=4`, `accentBorderCallout4`, `callout4`) —
  no callout-4 exists in ECMA-376 `ST_ShapeType`, so they only ever produced
  corrupt packages. `FOLDED_CORNER` is also corrected from the invalid
  `folderCorner` to the spec spelling `foldedCorner` (upstream #1449).
- Chart values now carry their number format into each series'
  `<c:numCache><c:formatCode>` (resolved from `valLabelFormatCode` /
  `dataTableFormatCode` / `dataLabelFormatCode`, default `#,##0`) instead of a
  hard-coded `General`, so PowerPoint and Google Slides honor `formatCode` the
  way LibreOffice already did. This deliberately changes default cached output
  to match the data-label format (upstream #1309).
- Identical media is now deduplicated: inline base64 `data:` media is reused
  per slide, and a deck-wide export pass collapses repeated images (including
  background images and SVG) to a single package part instead of embedding one
  copy per use (upstream #1339).
- `ChartLineCap` is now a deprecated alias for the shared `LineCap` type.

### Fixed

- Table merged cells (colspan/rowspan covered cells) now render the span's outer
  borders and fill instead of emitting an empty `<a:tcPr/>`; the origin cell's
  border tuple and resolved fill are applied to the covered edges to match
  PowerPoint-authored output (upstream #680).
- RGBA effect colors no longer emit a duplicate `<a:alpha>`: when a shadow/glow
  caller supplies an explicit alpha (notably on table-cell paths that skip
  `correctShadowOptions`), it wins and the RGBA byte is dropped, fixing
  schema-invalid double-`<a:alpha>` output that triggered PowerPoint repair.
- Text-box `margin` arrays are now mapped as `[top, right, bottom, left]`,
  matching table cells and slide numbers; previously Top and Left were
  transposed, mis-rendering asymmetric margins (upstream #1248).
- Out-of-range fill/line/gradient transparency, glow opacity, and line widths
  are clamped to schema-valid `<a:alpha>` / `<a:ln w>` ranges (warning on
  coercion) instead of emitting values PowerPoint rejects.
- Out-of-range `fontSize`, `charSpacing`, and `lineSpacing` are clamped to their
  schema ranges at run/paragraph emission (covering text boxes, table cells, and
  the slide-number placeholder).
- Chart `gapWidth`/`gapDepth`, `overlap`, `holeSize`, and `firstSliceAng` are
  clamped to their schema ranges via a shared helper (upstream #1233).
- Chart `lineDataSymbolSize` is rounded and clamped into the valid
  `ST_MarkerSize` range 2–72 (upstream #1233).
- Non-finite (`NaN`/`Infinity`) chart data values are dropped (with a warning)
  rather than emitting an invalid `<c:numCache>` that PowerPoint flags for
  repair; `null`/`undefined` remain valid sparse gaps (upstream #1357).
- Chart text (title, legend, axis labels, data labels) now stamps the requested
  typeface onto the `<a:latin>`/`<a:ea>`/`<a:cs>` trio so East-Asian and
  complex-script glyphs honor the chosen font (most visibly on PowerPoint for
  Mac) (upstream #1420).
- Scatter/bubble X axes in combo charts now emit `<c:valAx>` instead of
  `<c:catAx>`, fixing packages PowerPoint flagged for repair; an unsatisfiable
  shared-axis configuration now warns (upstream #1355).
- `addShape` (and the `shape` option on `addText`/`addImage`) now rejects
  unknown presets with a clear error at the `genXmlPresetGeom` chokepoint rather
  than emitting an invalid `<a:prstGeom>` that corrupts the package
  (upstream #1449).
- HTML-table conversion preserves fractional border widths (e.g. 0.5px hairlines)
  instead of rounding them to 0pt; a zero/non-finite computed width now yields
  `{ type: 'none' }` (upstream #1235).

## [5.2.0](https://github.com/shbernal/PptxGenJS/releases/tag/v5.2.0) - 2026-06-10

### Added

- `textRun(text, options?)` / `textRuns(runs)` factory helpers for building
  typed inline-run arrays without `as never` casts.
- Native pattern fills for shapes via `fill: { type: 'pattern', pattern: {
  preset, fgColor?, bgColor? } }`, covering the full OOXML `ST_PresetPatternVal`
  preset set.
- `defineTableStyle()` registers a custom reusable table style, and a
  `TABLE_STYLE` enum plus `tblPr` style flags expose the built-in styles.
- `hasHeader` table option emits `firstRow="1"` on `tblPr`.
- Slide masters accept a `roundRect` object and placeholder shapes.
- Chart `seriesOptions` sets per-series color and data-label overrides.
- Combo charts can suppress subchart series from the shared legend.
- Image `duotone` recolor option maps shadows/highlights to two colors.
- `firstSlideNum` sets a custom starting slide number for the presentation.
- `setCustomProperty` writes OOXML custom document properties.

### Fixed

- `textRun` / `textRuns` are now exported from every runtime entry
  (`node`, `browser`, `standalone`, `core`); previously only `index` shipped
  them, so `import { textRun }` type-checked but threw at runtime under the
  Node export condition.
- Image `cover` / `contain` crop is computed from the natural pixel ratio
  instead of the display ratio, fixing incorrect crop windows.
- Out-of-bounds image crop windows now throw instead of emitting a negative
  `srcRect`.
- Multi-level category charts use the correct embedded-workbook cell and
  shared-string-table indices.
- `round2SameRect` and `round2DiagRect` preset shapes emit `adj1`/`adj2`.
- Table `autoPage` shares line-wrap state across styled runs in a cell,
  preserves the originating slide section across overflow slides, and no
  longer breaks inside an active rowspan group.
- `breakLine: false` is preserved on the last piece of a CRLF-split run.
- SVG PNG previews use a transparent placeholder instead of the broken-image
  icon.
- Image hyperlink URLs are XML-entity encoded.

## [5.1.0](https://github.com/shbernal/PptxGenJS/releases/tag/v5.1.0) - 2026-06-09

### Added

- `catAxisLabelFormatCode` on scatter and bubble charts sets an independent
  number format for the X (horizontal) axis, decoupled from
  `valAxisLabelFormatCode` which controls the Y axis.
- `lineDashValues?: ChartLineDash[]` on line, scatter, and bubble charts sets
  a per-series dash pattern; entries fall back to the chart-level `lineDash`
  default.
- `addImage({ shape })` clips a picture to any preset geometry (e.g.
  `'hexagon'`, `'roundRect'`). `rounding: true` remains a shorthand for
  `shape: 'ellipse'`. `shape` takes precedence when both are set.
- `addImage({ points })` clips a picture to an arbitrary freeform path
  (`custGeom`) using the same path DSL as freeform shapes (`moveTo`/`lnTo`/
  `cubicBezTo`/`quadBezTo`/`arcTo`/`close`). Takes precedence over
  `shape`/`rounding`.
- `addImage({ svg })` accepts raw SVG markup directly, converting it to a
  base64 data URI internally. `data`/`path` still win when also supplied.
- `altText` prop extended to text boxes, shapes, tables, and media objects.
  Previously only images and charts emitted `p:cNvPr descr`.
- Object name validation: warns (without throwing) on names that cannot
  provide a stable Selection Pane identity — empty/whitespace, control
  characters, names over 255 chars, or duplicates on the same slide.
- `bullet.color` (HexColor) colors a bullet glyph independently of the
  text run color via `<a:buClr>`.
- `TextBaseProps.caps` (`'none'` | `'small'` | `'all'`) applies all-caps or
  small-caps styling to a text run.
- `valAxisCrossBetween` (`'between'` | `'midCat'`) exposes the OOXML
  `crossBetween` setting on the value axis.
- `STANDARD_LAYOUTS.*` now expose `.width` / `.height` inch aliases.
  `pptx.layout` accepts a preset object directly (e.g.
  `STANDARD_LAYOUTS.LAYOUT_16x9`). `slide.width` / `slide.height` getters
  return the active layout size in inches.
- `displayBlanksAs: 'zero'` added as a valid chart option value.

### Fixed

- `getSmartParseNumber` now throws on `NaN`/`Infinity` instead of silently
  collapsing objects to zero size or position.
- XML 1.0 illegal control characters (U+0000–U+0008, U+000B, U+000C,
  U+000E–U+001F, U+007F) are stripped before serialization, preventing
  PowerPoint repair dialogs.
- `createColorElement` guards against non-string input, preventing a
  `TypeError` when an object is passed via `chartColors`.
- Table `write()` / `writeFile()` is now idempotent on merged-cell tables;
  the internal row expansion no longer mutates the caller's array.
- Scatter/bubble chart data labels now apply `dataLabelFontSize`,
  `dataLabelFontBold`, `dataLabelFontItalic`, `dataLabelColor`, and
  `dataLabelFontFace` to custom label `rPr` elements.
- Slide master and layout media targets are namespaced so they no longer
  collide with regular slide targets in large decks.
- Line charts now emit `c:grouping` (required, defaults to `'standard'`) and
  respect `barGrouping: 'stacked'`.
- Single-level category labels now emit `c:strRef/c:strCache` instead of
  `c:multiLvlStrRef`, improving Google Slides and other importer
  compatibility.
- Chart zero values are preserved in embedded workbook cells; the previous
  `||` guard treated `0` as blank.
- Pie and doughnut parent `dLbls` now use `dataLabelFontSize`,
  `dataLabelColor`, and `dataLabelFontFace` instead of hard-coded defaults.
- Transparent `chartColors` entries on line/radar charts now emit
  `<a:noFill/>` on markers instead of a solid fill.
- Chart null values now omit `<c:pt>` entirely (correct OOXML gap encoding)
  rather than emitting empty `<c:v/>`.
- Stray apostrophe removed from embedded workbook table-ref attribute,
  fixing chart rendering in Apple Keynote.
- Pie/doughnut `dataLabelPosition` is now applied to the parent `dLbls`
  block instead of being hard-coded to `'ctr'`.
- Shadow `blur`, `angle`, and `opacity` zero values are now honored instead
  of being replaced by defaults.
- `barOverlapPct` is respected on stacked bar charts; previously the
  stacked-bar path forced `100` before the user value was checked.
- Scatter/bubble cat-axis now reads `catAxisLabelPos` instead of
  hard-coding `'nextTo'`.
- `catAxisOrientation` and `valAxisOrientation` type unions now include
  `'maxMin'`; XML emission was already correct.

### Changed

- `displayBlanksAs` default changed from `'span'` to `'gap'`.

## [5.0.2](https://github.com/shbernal/PptxGenJS/releases/tag/v5.0.2) - 2026-06-08

### Added

- Native linear gradient fills for shapes.
- Public slide-layout unit helpers and package inspection primitives.
- Generated documentation site and object identity reference documentation.

### Fixed

- Zero chart axis crossing values are preserved instead of being treated as
  absent.
- Company metadata XML is escaped before serialization.
- Inner shadow XML is closed correctly.

### Changed

- Regression tests were reorganized into the current suite layout.

## [5.0.1](https://github.com/shbernal/PptxGenJS/releases/tag/v5.0.1) - 2026-06-07

### Added

- GitHub Actions npm publishing workflow for `@shbernal/pptxgenjs`, using npm
  trusted publishing and provenance on published GitHub releases.

### Changed

- Release documentation now describes the automated `publish.yml` workflow, tag
  guard, manual retry path, and post-publish checks.

## [5.0.0](https://github.com/shbernal/PptxGenJS/releases/tag/v5.0.0) - 2026-06-07

### Added

- Scoped package release target: `@shbernal/pptxgenjs`.
- Package-boundary validation for the scoped default import and subpath imports:
  `@shbernal/pptxgenjs/core`, `@shbernal/pptxgenjs/node`,
  `@shbernal/pptxgenjs/browser`, and `@shbernal/pptxgenjs/standalone`.

### Fixed

- Multiple `<a:pPr>` elements emitted per `<a:p>` cause "needs repair" — paragraph properties were re-emitted for every text run [\#1322](https://github.com/gitbrent/PptxGenJS/issues/1322)
- `[Content_Types].xml` emits a slideMaster `Override` per slide instead of a single Override matching `slideMaster1.xml` [\#1444](https://github.com/gitbrent/PptxGenJS/issues/1444) [\#1449](https://github.com/gitbrent/PptxGenJS/issues/1449)
- `addShape()` with bare-string aliases (`"oval"`, `"rectangle"`, `"roundedRectangle"`) emits invalid OOXML preset names that PowerPoint strips during repair
- Solid-color slide background omits `<a:effectLst/>` inside `<p:bgPr>`, triggering the "needs repair" dialog [\#1442](https://github.com/gitbrent/PptxGenJS/issues/1442)
- Shapes added without text emit `<p:sp>` with no `<p:txBody>`, triggering the "needs repair" dialog [\#1441](https://github.com/gitbrent/PptxGenJS/issues/1441)
- Non-numeric table cell `margin` values leak `NaN` into `<a:tcPr>` `marL/marR/marT/marB` attributes
- `notesMaster` rel resolves to `theme1.xml` (the slideMaster theme) instead of its own `theme2.xml` part [\#1443](https://github.com/gitbrent/PptxGenJS/issues/1443) [\#1449](https://github.com/gitbrent/PptxGenJS/issues/1449)
- Calling `writeFile()`/`stream()`/`write()` more than once on the same Presentation mutates `options.shadow` and produces invalid EMU values on subsequent writes
- `addShape()` with a `#`-prefixed shadow color emits invalid `<a:srgbClr val="#...">`
- 8-character hex (RGBA) color values silently fall back to black and discard alpha; shadow colors emit invalid 8-char `val` attributes
- Unpopulated layout placeholders render a "Click to add text" hint over populated content because the empty stub was stored as text rather than placeholder
- `bullet:{type:"bullet"}` emits no bullet markup; `characterCode` was unreachable when combined with `type`
- Leading bullet glyphs in user text (e.g. `addText("• item", {bullet:true})`) render alongside the paragraph-level bullet, producing double bullets
- `[Content_Types].xml` emits Default `Extension` entries for media types not present in the deck (and a `vml` entry with no corresponding part)
- Every output `.pptx` contains stray empty `ppt/charts/`, `ppt/charts/_rels/`, and `ppt/embeddings/` directories even when the deck has no chart
- Combo charts with `secondaryValAxis`/`secondaryCatAxis` flags emit dangling axis-ID references; 2D bar/line/area/radar charts emit a series-axis reference with no matching definition
- `<p:presentation>` child elements emitted in non-canonical order — `<p:notesMasterIdLst>` now appears before `<p:sldIdLst>` to match the OOXML CT_Presentation child sequence (ECMA-376 Part 1 §19.2.1.26)

### Changed

- Package version is now `5.0.0` because this release intentionally narrows the
  package contract to modern ESM consumers and Node.js `>=24`.
- CommonJS, IIFE/global browser bundles, direct CDN script-tag workflows, and
  legacy generated artifact names are not maintained package targets.
- Release documentation now lives under `docs/RELEASING.md` with a manual npm
  publishing path for the first scoped release.
- `npm test` now runs both the regression suite and the OOXML schema-validation suite. Schema validation requires a one-time `./tools/ooxml-validator/install.sh` to download the validator binary.

## Older releases (v4.0.1 and earlier)

Pre-fork upstream history (gitbrent/PptxGenJS, v4.0.1 back to v1.0.0) has been
moved to [`docs/changelog-archive/pre-fork-v4.0.1-and-earlier.md`](docs/changelog-archive/pre-fork-v4.0.1-and-earlier.md)
to keep this file focused on the fork's own releases (v5.0.0 onward).
