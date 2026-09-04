# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **`Presentation.appProperties` reads `docProps/app.xml`.** The write API has had a
  `pptx.company` setter since the beginning and the read model could not see the part it
  writes to, so `ts-pptx/script` had to declare `company` unreadable and a converter could
  carry the other four `docProps` and never that one. The new accessor reports four fields
  -- `application`, `appVersion`, `company` and `titlesOfParts` -- and `readExtendedProperties`
  plus the `ExtendedProperties` type are published alongside the core/custom pair.

  The part's *statistics* (`Slides`, `Words`, `Paragraphs`, `HiddenSlides`, …) are
  deliberately not reported: they are numbers the producing application computed for the file
  it wrote, and this read model can hand back an edited deck, so reporting them would state a
  fact about a document that no longer exists. `titlesOfParts` is the flat `vt:lpstr` vector
  as written -- fonts, themes and slide titles in one list -- because `<HeadingPairs>` holds
  the counts that partition it and is not read; a caller who wants the slide titles alone
  pairs the two itself.

  `ts-pptx/script` carries `company` again as a result: `DeckPropsIr` has the field back, a
  standalone script emits `pptx.company = …` when the source states one, and the
  `deck.docProps` note no longer claims a property is lost that now round-trips. Five of the
  thirteen document properties survive a standalone conversion, against four before.

- **`StrokeProps`: one stroke, one vocabulary.** A stroke is width, dash, colour, cap and
  transparency, and the public type surface spelled it five ways -- `width` on a border,
  `size` on a gridline, `catAxisLineSize` on an axis; `type` + `dashType` on a border,
  `style` on a gridline, `catAxisLineStyle` on an axis; `HexColor`, `Color` and bare `string`
  for the same paint. `StrokeProps` is that vocabulary stated once, `BorderProps` is now an
  alias of it, and the chart option bag takes it wherever it used to spell a stroke by hand:

  - `OptsChartGridLine` and `ChartErrorBarOptions` extend it (both were already nested
    objects, so nothing about the flat bag changes).
  - `catAxisLine`, `valAxisLine` and `serAxisLine` are new nested `StrokeProps` on
    `ChartOpts`, superseding the twelve flat `*AxisLineColor` / `*AxisLineShow` /
    `*AxisLineSize` / `*AxisLineStyle` keys.

  Three capabilities come with it, all of which the old spellings could not express:
  **the full `ST_PresetLineDashVal` set** on an axis line, a gridline and an error bar (they
  took `solid | dash | dot` and nothing else, so a caller replicating a source deck could
  carry `lgDashDot` on a shape and not on an axis); a **`cap`** and a **`transparency`** on
  an axis line and an error bar; and a **`dashType`** on an error bar, which had only a
  colour and a width.

  Every pre-existing spelling still works and emits the same bytes. Twenty-five legacy
  option combinations were diffed part-for-part against the previous tree, because the
  byte-identity showcase decks reach only `*AxisLineShow` and gridlines -- `catAxisLineStyle`,
  `*AxisLineSize`, `*AxisLineColor`, error bars and `barSeriesLine` have no coverage there,
  and a green gate would have proved nothing about them.

- **Every published subpath now exports `setDiagnosticHandler` and `resetDiagnosticState`.**
  The library's warnings — a chart point cache out of range, a picture whose relationship
  does not resolve, a table span the auto-pager refuses — are reported through a handler a
  consumer installs, but that handler was published only by the three authoring entries
  (`.`, and the node/browser conditions of it). A consumer of `ts-pptx/read`,
  `ts-pptx/measure`, `ts-pptx/script`, `ts-pptx/inspect`, `ts-pptx/html`, `ts-pptx/math` or
  `ts-pptx/zip` got `console.warn` output from those paths with no supported way to
  intercept it.

  Importing the handler from `.` did happen to work — bundling puts the diagnostics module
  in a shared chunk — but that is an artifact of chunking rather than a promise, and it
  pulls the whole write path in for a three-line function. The surface is now republished
  from one module the way the error taxonomy already was, so one installed handler serves
  every subpath, and a test pins that they are literally the same function.

- **The series axis line can be sized and dashed.** The third axis of a 3-D chart could be
  shown and coloured but not sized or dashed: the emitter hardcoded one point and `solid`
  where the other two axes read the caller. Comparing the three axes' option spellings side
  by side is how that was found, which is also what produced `StrokeProps` below. Spell it
  `serAxisLine: { width, dashType }`; the flat `serAxisLineSize` / `serAxisLineStyle` exist
  too, deprecated from birth, so that a caller already using the flat form on the other two
  axes is not left reaching for a key that does not exist. A chart that sets none of them
  emits what it always did.

- **`Shape.absoluteFrameFailure` names why `absoluteFrame` is `null`.** That one `null`
  stands for three different situations — the shape states no transform of its own
  (`'no-own-transform'`, an ordinary placeholder inheriting its box from the layout), an
  enclosing `p:grpSp` states no complete `a:off`/`a:ext` **and** `a:chOff`/`a:chExt` pair
  (`'group-transform-missing'`), or an enclosing group's `a:chExt` is zero on an axis
  (`'group-transform-degenerate'`) — and only a caller that wants to *report* the
  unresolvable shape needs them apart: the first is normal, the other two say the deck is
  malformed. The new getter is `null` when the frame resolved; both project from a single
  ancestry walk, so they cannot disagree. A missing group transform outranks a degenerate
  one wherever the two meet in one chain, since composing needs every group's mapping.

  `ts-pptx/inspect` was the caller that needed the distinction, and it re-derived it by
  walking the same ancestry a second time over the raw DOM. It now reads the getter,
  which drops its last three DOM helpers and its only `element_` escape hatch — the
  surface is now the pure projection over `ts-pptx/read` its own header claims. No
  diagnostic message or dropped element changes.

- **`importSlides` takes `embedFonts` and `rescale`, so the batch path is no longer the
  one missing two options.** The batch is the import with the all-or-nothing guarantee, so
  it is the one a caller should reach for, and it was also the one that could not carry a
  source deck's embedded fonts or put a differently-sized source on this canvas. Both are
  now per request, spelled as they are on `importSlide`.

  Both are really whole-*deck* decisions wearing a per-page spelling, so the batch
  reconciles them before it moves anything:

  - `embedFonts` carries the source deck's **whole** face list, once, when any of that
    source's requests asks. `p:embeddedFontLst` does not record which page uses which
    face, so there is nothing finer to carry. Merging de-dupes by typeface and face slot
    as before. The font parts are part of the up-front dry run
    (`checkEmbeddedFontsCopyable`) for the same reason the notes are: the carry runs after
    the pages are copied, so a source missing a font binary would otherwise throw with
    parts already added and no way back.
  - `rescale` must **agree** across the requests naming one source, or the batch is
    refused with the new `import/rescale-conflict` before anything is copied. A batch
    import is `'copy'` themed, so a rescale rewrites the imported layout and master shape
    trees alongside the page, and those are shared between that source's pages: rescaling
    one and not another would leave the second aligned against a master that had moved
    under it. `true` and `'fit'` are the same answer, not a disagreement, and different
    sources are independent.

  A size mismatch without the option is still `import/slide-size-mismatch`, and its
  message now names the spelling that answers it.

- **A native picture fill takes a source `crop`, so `fill: { type: 'image' }` is no longer
  the one picture that cannot be cropped.** `ImageFillProps.crop` is the same
  `{ l, t, r, b }` percentage inset that `addImage` has always taken, validated the same
  way and serialized to the same `<a:srcRect>` — a shape, a table cell or a background
  filled with an image can now show a sub-region of its source instead of the whole thing
  stretched into the box:

  ```js
  slide.addShape(ShapeType.triangle, {
    x: 1, y: 1, w: 4, h: 3,
    fill: { type: 'image', image: { path: 'portrait.png', crop: { t: 25, b: 25 } } },
  })
  ```

  The shape geometry stays the mask; `crop` only decides which pixels arrive inside it.
  Note that the remaining region is still *stretched* to the fill's box, so a cover-style
  placement means computing insets that match the box's aspect ratio yourself. Where the
  output does not have to be a filled shape, `addImage({ shape, sizing: { type: 'cover' } })`
  is the better answer to that: it reads the source's own dimensions and derives the
  aspect-correct crop for you (see `docs/image-in-shape.md`).

  Omitting `crop` emits the same empty `<a:srcRect/>` as before, byte for byte. An inset
  outside 0–100, or a pair that would leave no source area, throws
  `image/crop-inset-out-of-range` / `image/crop-insets-exceed-extent` as it does on
  `addImage`.

  Thanks to [@flyisland](https://github.com/flyisland) ([#28](https://github.com/shbernal/ts-pptx/issues/28), [#29](https://github.com/shbernal/ts-pptx/pull/29)).

- **`Slide` declares `addSlideZoom`, `addSectionZoom` and `addSummaryZoom`.** `SlideBuilder`
  has implemented all three since zooms shipped, `ZoomBaseProps` and its three option types
  are exported, and the doc comments describe them — but the published `Slide` interface did
  not list them, so a TypeScript consumer calling `slide.addSlideZoom(...)` got
  `Property 'addSlideZoom' does not exist on type 'Slide'` and had to cast. Purely additive:
  the runtime is unchanged, and the gate deck that carried an `any` cast for exactly this
  no longer needs one.

### Changed

- **`radarStyle` accepts `ST_RadarStyle`'s own spellings.** The option had one vocabulary --
  `radar` / `markers` / `filled`, the PowerPoint UI's words -- which the emitter mapped onto
  the schema's `standard` / `marker` / `filled`. Since `filled` is spelled the same in both,
  the other two read as typos rather than as a second vocabulary, and writing `marker` (what
  the chart part itself shows) warned and silently fell back to a plain radar. Both
  vocabularies are now accepted and normalized to the wire member at definition time, the
  `RadarStyle` and `RadarStyleAlias` types are published, and an unknown value still warns
  and falls back. No emitted bytes change for any spelling that already worked.

- **A text box with no stated height gets 0.3in of one.** `addText('hi', { x, y, w })` used
  to emit `<a:ext cy="0">` -- the degenerate zero-size object this project's own API rules
  refuse -- because `addTextDefinition` defaulted no axis where `addShapeDefinition` defaults
  all four. A rescue for it existed in the text serializer, guarded by `!itemOpts.line`, but
  the definer has written `line = line || {}` onto every text object since before that guard
  was added, so it had never once fired.

  The default now lives beside the other definer defaults, where the three statements stay
  distinct: silence takes 0.3in, a stated `h: 0` is kept (as `addShapeDefinition` keeps one),
  and a `line` shape is exempt because a horizontal rule is a zero-height shape on purpose.
  The serializer could not have drawn those distinctions -- an omitted height and a stated
  zero both reach it as `cy === 0`.

  **Migration:** a caller relying on a zero-height text box now states `h: 0` for it. All 47
  byte-identity showcase decks are unchanged, so no deck in the corpus was relying on the old
  behaviour.

- **`ShapeLineProps` is deliberately not a `StrokeProps`.** A shape stroke is a *paint*, so
  its `type` names a fill kind (`gradient`, `pattern`, `inherit`, …) rather than the
  three-way dash switch a border's does. The four keys the two share -- `width`, `dashType`,
  `cap`, `transparency` -- already agreed and are unchanged. This is written down because
  "make every stroke one type" is the obvious next step and it is the wrong one.

- **`seriesOptions` now warns on the chart types that do not read it.** It is documented
  chart-wide and honoured by one plot family: `bar`, `bar3d`, `line`, `area` and `radar`
  resolve it per series, while `scatter`, `bubble`, `bubble3d`, `pie`, `doughnut`, `stock` and
  `surface` build their series colours straight from the palette and never look at it. A
  documented public option that does nothing on six of eleven types, with no warning and no
  type-level signal, is the state the library's option rules forbid: the caller said it and
  nothing happened.

  Setting it on one of those now raises `chart/option-not-supported` and the type says which
  families read it. Nothing about the emitted chart changes. A pie has no referent for a
  *series* override even in principle — it colours points — while the others are a gap whose
  fix needs a question settled first: a scatter's `data[0]` is the shared X row, so it is not
  obvious which series `seriesOptions[0]` names.

- **A chart option that is not a number now throws instead of silently taking the default.**
  BREAKING for a caller passing `NaN` or a non-number where a number is declared.
  `clampRangedInput` states the library's one policy for a value outside a schema range, and
  `docs/diagnostics.md` describes it: a finite value has a nearest legal neighbour, so it
  clamps and warns; a value that is not a number at all has none, so the request is discarded
  and that throws. The chart clamp answered `undefined` for the second case — discarding the
  request and reporting nothing — so `holeSize: NaN` quietly took the default while
  `holeSize: 200` warned. Same option, same class of mistake, two behaviours.

  Affects `holeSize`, `firstSliceAng`, `barGapWidthPct`, `barGapDepthPct`, `barOverlapPct` and
  `lineDataSymbolSize`, on the chart and on a combo subchart's overrides, under the new
  `chart/option-non-finite` code. An omitted option still means omitted. A fractional in-range
  value is still rounded and now says so — these are integer schema types, so `holeSize: 42.5`
  is as much a correction as `holeSize: 200`.

- **An object's own coordinates now beat the frame of the placeholder it names.** BREAKING
  for a deck that states both. `addText('own coords', { placeholder: 'body', x: 5, y: 3, w: 2,
  h: 1 })` used to emit the layout placeholder's box and discard all four stated values, with
  no diagnostic — while the same object with a *partial* frame and no placeholder warns
  loudly. Three states (inherit the box, override it, override part of it) had one spelling
  between them, and the decision is the one the rest of the library makes everywhere: an
  explicit option beats an inherited one, and a placeholder's frame is an inherited one.

  Each axis resolves independently, so `{ placeholder: 'body', x: 5, w: 2 }` takes `x` and `w`
  from the object and `y` and `h` from the placeholder. A stated `0` counts as stated. The
  resolution also moved *before* the negative-extent normalization: the placeholder's extents
  used to skip it while the flip flags were derived from the object's own signs, so a negative
  extent on either side composed wrong.

  Every other option a placeholder states is still imposed on the object, unchanged.

- **`SectionProps.order` counts from 1, as it has always documented.** BREAKING for a deck
  passing `order`. It was spliced in raw, so `order: 1` inserted the section *second* and
  `order: 0` — falsy — appended it with nothing said. `order: 1` is now the first position.
  An order past the end appends. Anything that is not a whole number of at least 1 warns with
  the new `section/invalid-order` code and appends, rather than splicing from somewhere
  unpredictable.

  Migration: subtract one from any `order` you pass today, or drop it where you were relying
  on `0` to append.

- **A hole in a table-level `border` tuple is left alone, exactly as it is on a cell.**
  `[rule, null, rule, null]` meant two different things depending on where it was written: on a
  cell a `null` side is *omitted*, so the edge keeps inheriting from the built-in table style,
  while at table level the definer filled the hole with an explicit `{ type: 'none' }` — which
  is direct formatting and erases the style's rule. And because the same array object reached
  every cell by reference, the filling propagated into cells that had already captured it. Both
  levels now read a hole the same way.

  **Migration:** a table that was relying on a `null` side to erase the style's rule spells it
  `{ type: 'none' }`, which is what the cell path has always required.

- **`Picture.imagePartName`, `.svgPartName` and `.mediaPartName` return `null` where they used
  to throw.** `Relationships.resolveTarget` throws on a dangling id and on an external target,
  which is right for an accessor asking a question a malformed package cannot answer — and
  wrong for one that already reports "absent" as `null`. The *same* `a:blip/@r:embed` reached
  through `AutoShape.pictureFill`, `TableCell.pictureFill` or `Slide.background` degraded to
  `null` and through `Picture.imagePartName` threw, so one broken embed on an imported deck
  took out a whole `slide.shapes` walk; `mediaPartName`, the "just give me the bytes"
  accessor, inherited the throw from both halves. A `p:pic` carrying a *linked* image — which
  `mediaKind`'s own documentation names as its `'none'` case — threw rather than reporting it.

  **Migration:** a caller that was catching `relationship/not-found` or
  `relationship/external-has-no-partname` around these three getters checks for `null`
  instead. The `GraphicFrame` accessors that deliberately let it throw are unchanged.

- **Enumerated attributes are checked against their `ST_` union on every write path, not on
  four of them.** `check-enum.ts` is the declared policy for this and had four call sites
  against roughly thirty hand-rolled ones, which is how the *same type* reaching the *same*
  attribute got two answers: `addTable(..., { border: { dashType: 'bogusDash' } })` warned
  and fell back to solid, while `addShape('rect', { line: { dashType: 'bogusDash' } })`
  wrote `<a:prstDash val="bogusDash"/>` into the part. Nine dash sites now share one
  `resolveDash`; `a:headEnd`/`a:tailEnd` (`ST_LineEndType`), `a:bodyPr/@vert`
  (`ST_TextVerticalType`) and `a:prstTxWarp/@prst` (`ST_TextShapeType`) are checked for the
  first time, the last two reporting under the new `text/invalid-vertical` and
  `text/invalid-warp`; an unrecognized arrowhead reports under the new
  `line/invalid-arrow-type`.

  The chart definer's own twenty-odd inline `Array.includes` tests — three of the lists
  written out twice, verbatim — are now one `chartEnum` wrapper over tuples in
  `ooxml/st-enums.ts`, and report under the new `chart/invalid-option-value` instead of
  correcting in silence. The tuples are the library's accepted set rather than the schema's
  where the two differ, which is documented on each: `lineDataSymbol` deliberately omits the
  `ST_MarkerStyle` members PowerPoint does not draw, and `radarStyle` is not `ST_RadarStyle`
  at all but a library vocabulary mapped at emit time.

- **`ShapeLineProps.beginArrowType` / `.endArrowType` and `TextPropsOptions.textWarp` are
  narrowed to their `ST_` unions.** The two arrow options repeated their six members inline
  and `textWarp` was `string`; they now take `LineEndType` and `TextShapeType`, derived from
  the same tuples the runtime check reads, so the type and the validator cannot disagree.
  **Migration:** none for a value that was already legal.

- **A bare colour is now the documented shorthand for a solid fill, at every fill option.**
  `fill: 'FF0000'` says exactly what `fill: { color: 'FF0000' }` says and emits the same
  bytes, and the runtime had accepted it that way for a long time — but the types said
  `ShapeFillProps`, so only JavaScript callers could reach the spelling and only by writing
  something TypeScript rejected. `ShapeProps.fill`, `TextPropsOptions.fill`,
  `TableCellProps.fill` (so also `headerRow` and `columns[i]`), `TableProps.fill`,
  `TableProps.tableFill` and a chart's `plotArea.fill` / `chartArea.fill` now take the new
  `FillOption` (`Color | ShapeFillProps`); `Slide.background`, `SlideLayout.background` and
  `SlideMasterProps.background` take `BackgroundOption` (`BackgroundProps | Color`) for the
  same reason.

  Settled as one decision for the whole surface rather than per key: every one of these
  keys hands its option to the same `genXmlColorSelection`, so a rule that held at some of
  them and not others would make the API less predictable, not more. The shorthand is
  lossless — it is the object form minus the keys it has no way to say — which is why it is
  documented rather than deprecated: unlike `coord/bare-number-is-inches` or
  `margin/legacy-points`, there is no changed meaning to warn about.

  **A stroke is deliberately not included.** `line` carries width and dash alongside its
  paint, and those defaults come from rebuilding the line object at definition time, so a
  bare string would paint the colour and silently ship an `<a:ln>` with no `w` and no
  `prstDash`. `ShapeLineProps` says so.

- **`PresLayout.width` and `.height` say which unit they are in, because it depends on the
  direction.** `defineLayout` reads them as INCHES — that is what its own example passes —
  while `pptx.presLayout` returns them in EMU. Both are public and both are the same type,
  so `defineLayout({ ...pptx.presLayout, name: 'Copy' })`, the obvious way to derive a
  layout from the current one, states a width of nine million inches. Every value is
  finite, so nothing used to warn.

  The fields are documented rather than split into a separate `LayoutDefinition`. A split
  would be a breaking change to the most-used type on the surface, and it is not what
  actually stops the bad deck: the bound below is. What the doc does is tell a reader which
  direction they are travelling in before they hit it.

- **`AlignV` is deprecated.** It is the enum form of the `VAlign` string union, and no
  option in the library is typed as it — unlike its sibling `AlignH`, which
  `TextPropsOptions.align` really does take. It was a third public spelling of a
  two-spelling fact.

  **Migration:** use the `VAlign` union (`'top' | 'middle' | 'bottom'`) for a `valign`
  option, or `TextAnchor` for a text body's own `a:bodyPr/@anchor`.

- **`exactOptionalPropertyTypes` is on, and a few published types widen to say what they
  always did.** The flag distinguishes a missing key from one present and holding
  `undefined`; the library now draws that line deliberately rather than by accident, with
  the rule written down in `docs/development.md` ("Absent versus present-but-`undefined`")
  and its two helpers in `src/options-internal.ts`.

  Nothing about the emitted OOXML changes: every step of the series was gated on
  byte-identity across all 1172 parts of the showcase decks. What a consumer can see is a
  handful of `.d.ts` declarations that now read `?: T | undefined` — `Slide.transition`,
  `Slide.background`, `Slide.color`, `Slide.slideNumber` and `SlideLayout.background`,
  which `SlideBuilder` implements as accessor pairs and so always has present. That is a
  widening, so no call that type-checked before stops doing so, and it means
  `slide.transition = undefined` is now expressible in a project that has the flag on
  itself.

  **If you turn the flag on downstream:** the option bags the write API takes are still
  declared `?: T`, which is deliberate. They are spread over one another inside the
  library, where an absent key inherits and a present `undefined` overrides, so the two
  are not interchangeable. Omit an option you have no value for rather than passing
  `undefined` for it.

- **An out-of-range percentage has one answer, and three options stop having their own.**
  `bullet.size`, `fit.fontScale` and `fit.lnSpcReduction` used to *reject* a value outside
  their range: warn, and emit no attribute at all. That is the combination
  `docs/diagnostics.md` ("Warn or throw?") exists to rule out, because the request is
  discarded and reported as a warning, and the caller reads the warning while getting a
  deck whose bullet is silently back at its inherited size. All three now go through the
  same clamp every other percentage option already used:

  - **Finite and out of range: clamp to the nearest bound and warn.** `bullet.size: 500`
    emits `<a:buSzPct val="400000"/>`; `fit.fontScale: 150` emits `fontScale="100000"`.
    The diagnostic codes are unchanged (`bullet/size-out-of-range`,
    `text/invalid-fit-percentage`) and still fire; only what happens next differs.
  - **Not a number at all: throw `percent/non-finite`.** There is no nearest legal value
    for `NaN`, and clamping it puts `val="NaN"` in the package. This is the same throw
    `transparency`, `opacity` and the other clamped percentages have always raised.

  **Migration:** if you were relying on an out-of-range `bullet.size` or `fit.*` being
  dropped, omit the option instead — that is still how you leave the value inherited. A
  `NaN` reaching any of the three is now an `InvalidOptionError` rather than a warning; so
  is a numeric *string* (`bullet.size: '80'`), which the coercing `Number()` in front of
  that option used to accept. The rule is written down in `docs/diagnostics.md` under
  "The rule applied: an out-of-range number".

- **`import/slide-size-unknown` now means what its name says, at every import entry point.**
  Five methods enforce equal slide sizes — `importSlide`, `importSlides`,
  `importSlideMasters`, `appendSlides` and `importShape`/`importShapes` — and they carried
  five copies of the same precondition that disagreed about one thing: what an *unknown*
  size is. `importSlide` reported a deck with no `p:sldSz` as `import/slide-size-unknown`,
  but only when a rescale had been requested; the other four folded it into
  `import/slide-size-mismatch` and printed the word `unknown` inside the message. So a
  consumer branching on `err.code` could not tell "these decks are different sizes" from "I
  could not read a size" anywhere but one method, and the `unknown` code existed for a case
  one of five call sites could raise.

  The two conditions are separate everywhere now. `import/slide-size-unknown` is raised
  whenever a deck does not declare a size, and `import/slide-size-mismatch` only when both
  are known and differ. The split is worth having because only the second one is
  *answerable*: a mismatch is what `{ rescale }` and `{ requireEqualSize: false }` exist
  for, while a size that is not there cannot be compared or rescaled onto no matter what
  the caller passes.

  **Migration:** a `catch` matching `import/slide-size-mismatch` will stop seeing the
  missing-`p:sldSz` case at the four methods that used to report it that way; match both
  codes to keep the old breadth. Nothing changes for two decks that both declare a size,
  which is every deck PowerPoint writes.

- **The pptx→script converter carries a picture fill's source crop, instead of noting it
  as lost.** `a:srcRect` on an `a:blipFill` was one of the things
  `fill.picture.geometry` / `table.cell.fill.picture.geometry` declared uncarried, for the
  good reason that the write path had no option to put it in. It does now, so the note was
  claiming a loss that no longer happens — and a note that excuses a difference and a note
  that never fires look identical to `script:census`, which is what makes a stale one worse
  than none. The crop now comes back as `image.crop`, exact rather than to within a
  rounding step: the reader's division by 100000 is undone by the same factor, so
  `l="33333"` returns as `33.333` and re-serializes to the integer it came from.

  The note stays for the two rects `image.crop` cannot hold, and now says which they are: a
  **negative** inset — how a `contain`-style fill bleeds its source past the surface — and a
  pair of opposite insets summing to 100% or more. Carrying either would emit a script that
  throws when it is run. Tiling, the destination `a:fillRect` inset, a non-zero `@dpi` and
  `rotWithShape="0"` are unchanged: still uncarried, still noted.

  No published census figure moves — no corpus deck crops a *fill*, as opposed to a picture
  — so this shows up as a note that no longer fires where a deck of your own uses one.

- **Default `objectName` indices are 1-based for every kind, and `addChart` now takes its
  index from the shared counter.** Six definers numbered their default Selection Pane name
  from 0 (`Shape 0`, `Text 0`, `Image 0`, `Connector 0`, `Media 0`, `Table 0`) and four from
  1 (`Group 1`, `Object 1`, `3D Model 1`, and the zoom tiles), while `addChart` did neither:
  it derived `Chart 0` by counting the chart objects *currently* on the slide, which is the
  numbering the per-slide counter exists to replace — group children are spliced out of that
  list, so a count taken from it can hand a later object a name already in use.

  All eleven are 1-based now, which is the base PowerPoint itself uses: it names an inserted
  rectangle `Rectangle 1`, and nothing it authors is ever suffixed `0`. `resolveObjectName`
  has no `base` parameter left to pass, so a definer added later cannot pick the other
  convention by accident.

  **Migration:** this renames defaulted objects only. Anything authored with an explicit
  `objectName` is untouched, as are master and layout placeholders, which default to their
  declared name. A consumer that matched on a generated name — `'Shape 0'`, `'Chart 0'` —
  should add 1, or better, set `objectName` explicitly: `docs/reference/object-identity.md`
  has always said generated names are not a stable identity.

- **BREAKING: the 20 `_`-prefixed internal fields are off the public type surface.**
  `TransitionProps._sndRId`, `HyperlinkProps._rId`, `ShapeFillProps._imgRid`,
  `TextPropsOptions._bodyProp` / `._lineIdx`, the picture-bullet `_rId` / `_rIdSvg`,
  `ObjectOptions`' six (`_placeholderIdx`, `_placeholderType`, `_connectorAdj`, `_startCxn`,
  `_endCxn`, `_szAuto`), `TableCell`'s seven (`_type`, `_lines`, `_lineHeight`, `_hmerge`,
  `_vmerge`, `_rowContinue`, `_spanOrigin`) and `_arrObjTabHeadRows` on both `TableProps` and
  `TableToSlidesProps` all landed in `dist/index.d.ts` and in consumer IntelliSense. Every one
  is **produced** by the write path rather than authored — a resolved relationship id, a
  normalized `a:bodyPr`, the auto-pager's own row split — so a caller who set one was feeding
  a pass its own output. `_bodyProp` in particular was a documented option key on a published
  interface, i.e. a de-facto API nobody intended.

  They now live on `…Internal` counterparts in `src/types/internal.ts`, which
  `src/types/index.ts` deliberately does not barrel — the pattern `ShadowPropsInternal`
  already set. Because every added member is optional, a public value is still assignable to
  its internal counterpart, so the *carriers* stay public: `hyperlink?: HyperlinkProps` on a
  shape, a run and a table cell is unchanged, and only the emitters that read or stamp the id
  name the internal type.

  **Migration:** none exists, and none is wanted — reading one told you nothing the public
  API does not, and writing one was undefined behaviour. `startShape`/`startShapeIdx` is the
  supported way to bind a connector (`_startCxn`/`_endCxn` were its resolved output), and
  `autoPageRepeatHeader` + `autoPageHeaderRows` is the supported way to repeat header rows
  (`_arrObjTabHeadRows` was theirs). No emitted byte changes: the byte-identity gate is clean
  across 1256 parts.

  Two things came out of the surface as a side effect, both improvements. **`TextBulletProps`
  is now a named public interface** rather than an anonymous arm of the `bullet` union, so a
  caller can build one as a variable. And **`ShapeLineProps` no longer `Omit`s `_imgRid`** by
  hand: it omits `image` alone, which is the rule it was really stating — a stroke cannot be
  a picture fill.

  One place still surfaces them, unchanged from before: `ts-pptx/measure`'s
  `buildFitParagraphs` takes the internal option shape, so that entry's `.d.ts` names
  `ObjectOptionsInternal`. That is now visible in the type's name instead of inlined.

- **Three oversized read modules are split along the seams they already named.** Pure moves,
  byte-identical, and the public export surface is unchanged — `package:lint`, the bundle
  ratchet and `test:package` all pass without a budget edit.

  `read/api/table.ts` (1064 lines) held `Table`, `TableRow` and `TableCell` with no shared
  state; it is now `table.ts` / `table-row.ts` / `table-cell.ts`, chained rather than circular,
  with `table.ts` re-exporting the other two and the cell value types. `read/api/text.ts` (1034)
  becomes `text/run.ts` / `text/paragraph.ts` / `text/frame.ts` / `text/edit.ts` behind a
  `text.ts` barrel. `read/api/shapes/base.ts` (959 → 880) hands its `a:effectLst` block to a new
  `shapes/effects.ts` as free functions over `(effectLst, ctx)`, which is all those five getters
  and two private helpers ever needed — their only reference to the class was `themeContext()`.
  And `script/from-read/shape.ts` (955 → 701) hands its shared paint mappers — transform, fill,
  line, arrows, shadow, glow, `p:style` — to `shape-paint.ts`, leaving the five per-kind call
  builders behind. `script/verify/diff.ts` was on the same list and is not split: the construct
  table it was mostly made of is now one table, and 565 lines is no longer oversized.

### Deprecated

- **The chart option bag's older stroke spellings.** All of them still work, are still read,
  and emit exactly what they did; they carry `@deprecated` so an editor points at the
  replacement.

  | Deprecated | Use |
  |---|---|
  | `OptsChartGridLine.size`, `ChartErrorBarOptions.size` | `width` |
  | `OptsChartGridLine.style: 'solid' \| 'dash' \| 'dot'` | `dashType` (same values, plus the other eight presets) |
  | `OptsChartGridLine.style: 'none'` | `type: 'none'` |
  | `catAxisLineColor` / `valAxisLineColor` / `serAxisLineColor` | `catAxisLine.color`, … |
  | `catAxisLineSize` / `valAxisLineSize` / `serAxisLineSize` | `catAxisLine.width`, … |
  | `catAxisLineStyle` / `valAxisLineStyle` / `serAxisLineStyle` | `catAxisLine.dashType`, … |
  | `catAxisLineShow: false` / `valAxisLineShow: false` / `serAxisLineShow: false` | `catAxisLine.type: 'none'`, … |

  **Precedence, where a caller sets both:** the new key wins, and an explicit `type` wins
  over one inferred from `style: 'none'` or `*AxisLineShow: false`. The one place the two
  halves fold *independently* is a hidden axis line: `*AxisLineShow: false` with a
  `*AxisLineStyle` still writes that dash, because the axis' `<c:spPr>` has always carried
  an `<a:prstDash>` alongside its `<a:noFill/>` and swallowing it would move bytes.

### Removed

- **`scripts/docs-init.mjs` and the `docs:init` script are gone.** The scaffolder planted a
  fresh repo's docs kit -- the `project-documentation` skill, `docs/docs.json`, four starter
  pages and the `docs:*` script block -- and in this repo every file it wrote already
  existed, so a run reported `preserve:` for all of them. What it still carried was two
  liabilities: inline copies of four docs pages that had long since diverged from the real
  ones with nothing checking, and a `PACKAGE_SCRIPTS` constant that, unlike the pages, was
  *not* preserve-on-conflict -- `addPackageScripts` rewrote any `docs:*` script whose
  command differed, so a stale entry there silently downgraded a real repo's docs pipeline,
  which had already happened once. Nothing in this repo ran it. **Migration:** none; the
  kit stays in git history if it is ever wanted as the seed of a template repo.

- **`ChartOpts` no longer extends `OptsChartGridLine`.** That inheritance put `color`,
  `size`, `style` and `cap` on the chart option bag itself, where nothing read them: a
  caller who wrote `addChart(data, { style: 'dash' })` typechecked and got nothing. They
  were never a chart-level concept -- the three real ones are `catGridLine`, `valGridLine`
  and `serGridLine`. **Migration:** move the value to whichever of those three you meant.
  The keys were dead in every version that had them, so no output changes.

- **The three `*AxisLineShow` defaults in `normalizeChartPlotAreaOptions`.** They wrote
  `true` over an absent flag, which was the last thing keeping "the caller said nothing" and
  "the caller said yes" distinguishable, and nothing read the distinction. The axis emitter
  now folds the flag into a stroke where only an explicit `false` says anything. No emitted
  byte moves.

- **`correctShadowOptions` is now `normalizeShadowOptions` and no longer mutates its
  argument.** It is not exported from any entrypoint, so this reaches no consumer; the rename
  is recorded because the old name promised a correction *of* the caller's object and that is
  what four internal call sites had come to depend on unevenly.

- **The five series-axis unit options are gone.** `serAxisMajorUnit`, `serAxisMinorUnit`,
  `serAxisBaseTimeUnit`, `serAxisMajorTimeUnit` and `serAxisMinorTimeUnit` each wrote an
  element `CT_SerAx` has no slot for, so every value any of them could take produced a
  chart part PowerPoint refuses to open. There is no correct value to keep them for.
  **Migration:** none is possible on the series axis; the category axis takes the same five
  through `catAxis*` when it is a date axis, and the value axis takes the numeric pair.

- **`PresLayout._sizeW` / `._sizeH` are gone.** Every writer set them equal to `width` /
  `height`, and the only two readers were spelled `presLayout._sizeW || presLayout.width` —
  two names for one fact, with the fallback proving the second was always enough. Because
  they were public and optional, a caller *could* set them differently, and only the table
  auto-width path would have honoured it. **Migration:** use `width` and `height`.

- **`TableCell._tableCells` is gone.** Zero reads and zero writes anywhere in the library.

- **`SlideMasterChartProps.opts` is gone; use `options`.** Both were declared, neither was
  documented, and the definer read `chart.opts || chart.options`, so the undocumented alias
  *won* when a caller had both — setting the documented one and leaving a stray `opts`
  behind got the stray. `options` is the spelling every sibling descriptor uses.

- **`CONNECTOR_PRESETS` is no longer exported.** Its only consumer was
  `connectorPresetFor`, eleven lines below it in the same file, and
  `docs/architecture.md` puts internal OOXML generators off the published surface unless
  deliberately exposed. **Migration:** name a connector by its `ConnectorType`
  (`'straight' | 'elbow' | 'curved'`), or by the `CONNECTOR_PRESET_NAME` union, which is
  still public.

- **`image` is gone from `ShapeLineProps`: a picture stroke is not expressible in OOXML.**
  `ShapeLineProps extends ShapeFillProps` carried `image` and `type: 'image'` along with
  everything else, and `<a:ln>`'s paint child is `EG_LineFillProperties` — `a:noFill`,
  `a:solidFill`, `a:gradFill`, `a:pattFill` and nothing else. There is no `a:blipFill`
  slot on a stroke to put a bitmap in. The option had therefore never worked: no call site
  registered the media for it, so `line: { image: { path } }` reached the emitter with no
  relationship, warned `image-fill/unresolved-media` about a rel it was never going to be
  given, and painted nothing.

  `ShapeLineProps` now subtracts `image`, `_imgRid` and the `'image'` member of `type`
  from the fill props it inherits, so TypeScript rejects the spelling outright, and the
  new `resolveLineKind` refuses it at run time (`line/image-fill-unsupported`, an
  `UnsupportedFeatureError`) for the JS caller types cannot stop. Refusing beats painting
  nothing: had the media ever been registered, the emitted `a:blipFill` inside `a:ln`
  would have been a package PowerPoint reports as needing repair.

  **Migration:** use `fill: { image }` for a picture interior. A stroke can still be
  `solid`, `gradient` or `pattern`.

### Fixed

- **`addText('hi', { shadow })` emitted two shadows.** One `<a:effectLst>` landed in the
  shape's `p:spPr` and another in the run's `a:rPr`, because the bare-string overload handed a
  single options object to the shape *and* to its lone run, and `shadow` was on the list of
  options a run inherits from its shape.

  `shadow` now means the shape's shadow on a shape's bag and the text shadow on a run's, and a
  run does not inherit its shape's. That is PowerPoint's own division: Shape Effects and Text
  Effects are separate gestures writing to separate elements, and only applying both gives both
  (`shadow-shape-vs-text.pptx` is one text box per state). Stating it in both places still
  works, and now takes two statements here as it takes two actions there.

  The aliasing behind it is gone: the bare-string overload wraps its text into a run that
  states nothing and inherits what a run inherits. Two things followed from that being the only
  thing keeping them alive:

  - **A shape-level hyperlink emitted `r:id="rIdundefined"`.** A shape and its runs carry two
    `hlinkClick` elements resolving two relationship ids, and only the runs' were registered;
    the shape read its id back off the object the run had minted one on. `addText([{ text }],
    { hyperlink })` was already broken this way.
  - **A line-shaped text box drew its 1pt default only through the string overload.** The
    defaults were computed on a pass that had no `line` bag to assign them to and were only
    seen by a *second* pass over the same object. The bag is installed first now, so one pass is
    enough and `shape: 'line'` draws the same line through either overload.

  `rtlMode` and `tabStops` gained the inherit list they never had -- the paragraph builder reads
  them off a run's bag, and they used to arrive only through the shared object.

- **An HTML row covered end to end by a rowspan was dropped, taking the table's shape with
  it.** A source row states only the cells it *starts*, so a row every column of which is held
  by a `rowspan` from above states none at all -- what `<tr></tr>` between two spanned rows
  means. The auto-pager dropped it as an empty buffer while the table emitter went on
  synthesizing that grid row's `vMerge` continuations, so they landed in the next row's
  `<a:tr>`: a two-column table came out with a four-cell row, which is the malformation
  PowerPoint offers to repair. An empty row that nothing covers is still dropped, for the same
  reason it always was.

  Found by a new end-to-end fixture set for ragged HTML tables
  (`test/regression/html/html-ragged-tables.test.js`). The column-measuring helper was
  unit-tested, but nothing pinned the end of the pipe -- whether its answer composes with the
  emitter's own merge synthesis into a rectangular table -- and the byte-identity corpus holds
  no imported HTML table at all.

- **A layout placeholder no longer overrides an option the caller stated.**
  `addText(text, { placeholder: 'body', ... })` spread the layout placeholder's options over
  the caller's, so the placeholder won on every key it stated: `valign`, `margin`, `bullet`
  and the text-style keys were all discarded without a word. The frame was fixed first
  (3.4.0); this is the rest of the same shape.

  The rule, now written on `TextPropsOptions.placeholder`: **a placeholder supplies an option
  the caller left out, and never imposes one over a stated value.** It is PowerPoint's own
  model. A layout placeholder given `anchor="b"` and `lIns="914400"`, with the slide's
  placeholder then re-anchored to the top, is saved as `<a:bodyPr lIns="914400" anchor="b"/>`
  on the layout and `<a:bodyPr anchor="t"/>` on the slide: only the overridden property is
  stated there, the inset is absent because it was never overridden, and the stated anchor is
  the one that applies. `placeholder-override.pptx` is that deck.

  "Stated" means what the *caller* wrote, captured before this library defaults anything onto
  the same bag. A placeholder-targeting object has its `bullet` defaulted to `false` on the
  way past, so reading statedness off the bag would have let that default beat the layout's
  bullet, which is a different claim from letting a caller beat it.

- **`seriesOptions` reaches the XY and stock plots, and says which fields it cannot.** The
  option was read by the category-axis family alone, so `scatter`, `bubble`, `bubble3d` and
  `stock` took their series colours straight from the palette. All four now read it:
  `color` and `lineSize` on the XY plots, plus the per-series data-label font on `scatter`,
  whose label builders run inside the series loop; `color` on `stock`, where it paints the
  volume bar and the close marker.

  The index is stated rather than left to inference. `seriesOptions[N]` is the **Nth series
  of the chart** -- the number the series carries in `<c:idx>`/`<c:order>` -- so a scatter's
  `seriesOptions[0]` styles `data[1]`, the first Y series, because `data[0]` is the shared X
  row and is not a series at all.

  The `chart/option-not-supported` warning is now per *field* rather than per chart type,
  because the gaps are per field: `lineSize` reached a stroke only on a series that draws
  one, so it was accepted and dropped on `bar`, `bar3d` and `area` while every other field on
  the same bag worked. `pie`, `doughnut` and `surface` still warn on everything -- a pie
  colours points and a surface colours bands, so a per-series override has no referent.

- **A combo chart restarted the colour palette in every subchart.** Each subchart's series
  builder looked its colour up by the series' position *within that subchart*, so a
  bar(2) + line(1) combo painted the line with palette entry 0 -- the same colour as the
  first bar -- while `<c:idx>`, `<c:order>` and `seriesOptions` were all keyed on the
  position across the whole chart. Only the colour restarted.

  Desktop PowerPoint does not restart it: a three-series clustered-column chart whose third
  series is switched to a line keeps that series' third colour and merely moves it from
  `Format.Fill` to `Format.Line`, read back over COM. Every per-series lookup in the
  category-axis family now keys on the overall position, which also fixes `lineDashValues`,
  whose entries past the longest subchart were unreachable. A single-type chart is
  unaffected -- the two indices are the same number there -- so only combo decks move bytes.

  `seriesOptions` and `lineDashValues` now document that index as "the Nth series of the
  chart", the same number the series carries in `<c:idx>`/`<c:order>`, rather than as a
  position in `data`, which a combo splits across subcharts.

- **Three percentage accessors reported `55.00000000000001` for a `55%` attribute.**
  `TextFrame.autofitFontScale`, `TextFrame.autofitLineSpaceReduction`, `Run.baselinePct`
  and a paragraph's percent line spacing each read the attribute through `parsePercent` —
  which divides the fixed-point form by 100000 to give a fraction — and then multiplied by
  100 to get back to a percent, re-introducing the rounding the division had just taken on.

  They now divide once, exactly, through a `parsePercentPoints` that reads the same union
  (`55000` and `"55%"` both give `55`). The values move by a floating-point ulp, which
  matters to an equality check and to a serialized snapshot: `inspect-surface.snapshot.json`
  had drifted from what its own generator produces, so the sidecar could not be regenerated
  without an unrelated diff.

- **A slide's own picture or theme background is no longer dropped.** `ts-pptx/script` had
  two background mappers and only one of them could carry a picture. The layout/master arm
  handled `solid`, `image` and `themeRef`; the slide arm handled `solid` and `none` and
  recorded everything else as

  > a {type} slide background is not expressible through the write API's background option

  which was not true for `image`: `SlideProps.background` takes the same `BackgroundProps`
  the layout arm authors, four files away, and `BackgroundIr.data` was a declared and
  documented IR field nothing could produce. There is now **one** mapper for both tiers, so a
  slide-scoped picture background carries its bytes and a slide-scoped `p:bgRef` bakes its
  resolved colour exactly as a layout's does. The note constructs stay tier-scoped
  (`slide.background` vs `master.background`), so an existing declared loss keeps its key.

  `BackgroundIr.transparency` gains its first producer at the same time, on both tiers: a
  background colour carrying an `a:alpha` now reaches `BackgroundProps.transparency`. The
  read model reports opacity as a 0-1 fraction and the write option takes transparency as a
  0-100 percent, and that conversion is the one place the two conventions meet.

  `slide.background`'s note-field list claimed `image`, which was never a key on either side
  — the IR spells a picture background `data` — so that entry excused nothing. It is
  `data`/`$asset` now, matching `image.data`.

  **What blocked this was the fixture, and it is now in the corpus.** No deck under
  `test/read/fixtures/` had a slide-scoped background that was not `solid` or `none`, so the
  slide arm's claim had nothing to contradict it — and the round trip could not catch it
  either, since it excludes exactly the *declared* losses and this loss was declared. The new
  `slide-background.pptx` is three PowerPoint-authored slides: a picture background, a
  slide-scoped `p:bgRef`, and a solid with an `a:alpha`.

- **A combo chart's line markers took a different palette entry from the line they sit on.**
  Two lookups for one series: the line body read the series' position *within its own
  subchart* and the marker read its position *across all of them*. Those are the same number
  for a single-type chart and not for a combo — a bar(2) + line(1) combo drew a `C0504D` line
  with `9BBB59` dots — and a `seriesOptions.color` override moved the body while leaving the
  dots on the palette. The marker now takes the series colour that was already resolved for
  the body, so the two cannot disagree. Emitted bytes change for a combo whose line or radar
  subchart is not the first one.

- **`dataLabelPosition: 'outEnd'` on a clustered bar chart is no longer silently dropped.**
  The two rules deciding which `ST_DLblPos` values a bar accepts ran as sequential `if`s, each
  keyed to the grouping it did *not* apply to, so the outcome was exactly inverted: the list
  containing `outEnd` was applied when the grouping was not clustered and the list without it
  when the grouping was not stacked. `{ barGrouping: 'clustered', dataLabelPosition: 'outEnd' }`
  — the most ordinary combination there is — emitted no `<c:dLblPos>` at all, while the stacked
  form that PowerPoint itself refuses was kept.

  Settled against PowerPoint over COM: setting `DataLabels.Position` to
  `xlLabelPositionOutsideEnd` on a clustered column chart is accepted and reads back, and on a
  stacked or 100%-stacked one it raises 0x80004005; inside-end, inside-base and centre are
  accepted on all three. Both rules are now one table from chart type (plus bar grouping) to
  the positions that plot allows.

- **A table-level `italic` (and the other text options the fitter reads) now reaches the
  cells it was measured against.** The emitter and the measured-fit pass each had their own
  list of what a cell inherits from its table, and the measure one's docstring claimed to
  mirror the emitter's while naming four keys — `italic`, `charSpacing`, `lineSpacing`,
  `lineSpacingMultiple` — that the emitter did not carry. So a table-level `italic` was laid
  out with italic metrics and emitted upright; on `fit: 'shrink'` a table-level line spacing
  made the solver compute a taller layout and bake a smaller font that the cell then rendered
  at default spacing, and `pptx.tableLayout()` reported geometry the file disagreed with.

  The text keys are now one shared list. The emitter's extra keys (`border`, `color`, `fill`,
  `textDirection`, `underline`) stay its own and say why: they paint, or the fitter does not
  model them. Emitted bytes change for a table that states one of the four keys at table level
  and has cells that do not.

- **A table's or a table cell's fill lost its transparency when converted to a script.**
  An `a:solidFill` carrying an `a:alphaModFix` is read on any surface, but only the shape's
  copy of the fill ladder passed the alpha on: a table background or a cell fill came back
  fully opaque, with no fidelity note, because a key neither the source nor the copy
  produces is invisible to the round trip. All three surfaces read one ladder now, so a
  generated script states `transparency` wherever the source deck did.

- **Eight blind spots in the round-trip verifier, which is the thing that decides whether a
  generated script rebuilt the deck it came from.** Each is a way it could answer *yes*
  without having compared, or report a defect where the loss was already declared.

  - **Seven note constructs were emitted with no entry in the coverage table** — `chart.data`,
    `line.align`, `fill.gradient`, `line.gradient`, `line.gradient.path`,
    `text.autofit.fontScale` and `text.autofit.lnSpcReduction`. A construct with no entry
    declares nothing, so the difference the note predicted came back as a defect. The table is
    now the source of the `NoteConstruct` union a note is recorded under, so an unmapped
    construct is a **compile** error; that alone found two more (`table.fill.gradient.schemeToken`
    and its cell twin). The corpus check that was supposed to catch this only sees constructs
    the fixtures happen to produce, which is why it stayed green through all nine.
  - **A note or a write-path default matched the last key of a path, at any depth.** So `type`,
    written about a fill's solid default, excused an added `bullet.type` — a character bullet
    that came back as a numbered list, waved through; `title`, written about `docProps`, excused
    an added chart title; and a note about `line.width` excused a table cell border's width in
    the same call. Both mechanisms match a dotted path SUFFIX now.
  - **A gradient stop's scheme-token note ignored the surface it was on.** The two notes beside
    it are scoped by surface and this one hardcoded the shape spelling, so a table background's
    gradient recorded `fill.gradient.schemeToken` while its difference landed on `tableFill` —
    a declared, genuine loss reported as an undeclared defect, on exactly the distinction the
    table mapper goes to trouble to keep.
  - **The aligner let an unkeyed item claim a slot a later keyed one would match.** With an
    unnamed shape ahead of a named one on one side and the reverse on the other, both pairs
    cross-matched and the report described two shapes that each had an exact counterpart as a
    full set of differences. Keyed matches are reserved before the positional walk now, which
    also puts a lost item's `null` beside the item that went missing rather than at the end.
  - **An unresolved asset reference compared equal to itself.** It fell back to
    `unknown:<name>` — the asset's own name — so if neither side resolved one, both
    canonicalised to the same string and the round trip came back clean on bytes it never
    compared. That is the one case where names carry no information, which is what the content
    digest exists to avoid. It throws `script/unresolved-asset-reference` now, matching the
    print path.
  - **`DeckPropsIr.company` was a field nothing populated**, so every standalone script carried
    a note asserting "the source deck declares no company" — a claim the converter never
    checked and one that is false for most real decks. `Company` lives in `docProps/app.xml`,
    which the read model does not open. The field is gone and the `deck.docProps` note now says
    what is actually true of it.
  - **A chart with no cached series claimed its workbook had been rebuilt.** The
    `chart.workbook` note was recorded before the guard that drops such a chart, so it emitted
    two contradictory notes — and `chart.workbook` maps to `['*']`, the widest exclusion in the
    table, applied to the case it least describes.
  - **The standalone printer collected one note and dropped it.** `printDocProps` ran after the
    note list had been snapshotted, so `deck.docPropsDefault` — which says the output declares
    document properties the source did not — reached neither the emitted script's header nor
    the returned notes. It now fires on every fixture, which is where the census gains an entry.

  Also: the asset-identifier deduplicator re-derived its collision candidate from the raw
  name rather than the sanitised one, so a base starting with a digit would have produced
  `1image_2` — not a legal JavaScript identifier. Unreachable today, and one line.

- **Building a table no longer writes into the caller's cell objects.** `addTableDefinition`
  takes ownership of the *table* options and left the per-cell ones aliased, so the border
  completion, the hyperlink rel id and the table-level inheritance all wrote through to the
  objects the caller passed in. One `rows` array reused across two tables — the obvious thing
  to do with a shared style fixture — came out styled by the *first* table both times, and the
  caller's cell came back holding `{ border, bold, color, fontSize, margin, _lineIdx }`. The
  emitter resolves the inheritance into a local bag now instead of stamping the stored model,
  which also makes two `write()` calls on one deck independent.

  The file's own invariant, three hundred lines above the offending line, has said "never
  mutates: `arrTabRows` holds the caller's own cell objects" throughout.

- **A shadow literal shared between two shapes comes back as written.**
  `correctShadowOptions` normalized its argument *in place* and returned it, and its four
  callers split on which of the two contracts they used — two assigned the result, two
  discarded it and relied on the mutation. It is `normalizeShadowOptions` now, pure, with all
  four callers assigning: the derived `_alpha` reaches the emitter as a value rather than as a
  side effect. Emitted bytes are unchanged, which is the point — the function was a trap
  rather than a live defect, because the obvious "make it pure" cleanup silently drops
  `transparency` and RGBA alpha on every text run unless all four move together.

- **A percentage stated in the string form is read as the value it states.** `ST_Percentage`
  and its relatives are *unions*: the fixed-point integer PowerPoint writes (`100%` →
  `100000`) and a decimal string carrying a literal `%` (`"62.5%"`) — and the string form is
  the **only** one the Strict profile has. `read/oxml/dom.ts` has carried the union parser and
  a note saying exactly this; four getters read the attribute as a bare number instead, so
  `Run.baselinePct`, `Paragraph.lineSpacing`, `TextFrame.autofitFontScale` and
  `TextFrame.autofitLineSpaceReduction` reported `null` for a value that was present, and the
  model's contract for those nulls says "absent". A Strict-profile or non-PowerPoint deck
  therefore came back with no baseline shift, no line spacing and no autofit scale.

  `a:buSzPct/@val` is genuinely `ST_TextBulletSizePercent`, which has no string form, so it
  stays a bare fixed-point read; `pctFromThousandths` now documents that it is for those.

- **`p:spTree/p:extLst` is the shape tree's own child, not a shape.** Three helpers over a
  shape tree asked "is this a shape" and only one of them excluded `p:extLst`, so
  `carriedDecorations` reported it as a decoration to carry — and `importSlide`'s
  `carryMasterGraphics` inserts each of those *before* the destination's own shapes, which
  puts an `extLst` where `CT_GroupShape` sequences `nvGrpSpPr, grpSpPr, (shape)*, extLst?` and
  makes the part invalid. The read corpus holds no direct `p:spTree/p:extLst` across 776 shape
  trees, so nothing existing could have caught it.

- **A table cell's `fillSchemeColor` reads the same element its siblings do.** It fell back to
  `a:tc` when the cell had no `a:tcPr` — a location `CT_TableCell` does not permit — so on a
  malformed deck it reported a scheme token that `resolvedFill` and `hasOwnFill` both denied,
  and on a well-formed one the arm was unreachable code that read as a deliberate rule.

- **Two read sites hand-rolled `xsd:boolean`.** `TableCell.anchorCtr` spelled the test a fourth
  way (and read the attribute twice to do it) and the `vt:bool` document-property decoder
  accepted `'TRUE'`, which `xsd:boolean` does not. Both go through the module that exists to
  make that judgement unnecessary. Neither was a live defect — both attributes default false —
  but a hand-rolled `=== '1'` passes every deck this library and PowerPoint produce and
  misreads the rest, which is the shape of bug the module was extracted to prevent.

- **A new slide id stays inside `ST_SlideId`.** The allocator named and enforced the type's
  minimum and not its maximum, so a deck near the ceiling got an out-of-range `p:sldId/@id`
  written with no diagnostic. Past the ceiling it now falls back to the lowest id in range the
  deck is not already using, and a deck holding every one of the two billion throws
  `slide/id-space-exhausted` instead of writing a number the format has no room for.

- **A table's usable width read the wrong two slide margins.** `resolveSlideMarginsInches`
  returns `[top, right, bottom, left]`, and `usableTableWidthEmu` took index 1 (the right
  margin) as its left-edge fallback and subtracted index 3 (the left margin) as the right
  one — both one index off, so the two errors cancel for a symmetric margin and are wrong by
  their difference for anything else. On a 10in slide at `x: 1` with margins
  `[0.5, 0.5, 0.5, 2]`, an auto-paged table came out 7.0in wide where 8.5in was available.
  The sibling site in the table definer had already been found and fixed, with a comment
  naming this exact swap; this was the copy that was missed.

- **`slideMargin` steers a table whether or not `autoPage` is on.** The definer never read
  it — only the pager did — so a table's default width, its `fitColumns: 'shrink'` target
  and its per-page reset all silently depended on an unrelated option. It is also declared
  on `TableProps` now rather than on `TableToSlidesProps` alone, which is where the pager was
  reading it from on both paths.

  The two width formulas behind it are one: `defaultTableWidthIn()` was
  `slide − right − left` and ignored `x`, while the pager's was `slide − x − right`, and the
  pager's own doc claimed they were "the same reading". Both now go through
  `usableTableWidthEmu`, so a table at a stated `x` is no longer sized as though it began at
  the left margin.

- **Every series' worksheet reference addresses its own column.** The embedded workbook lays
  every series out behind the FIRST series' label columns, one row per that series'
  categories; the chart XML derived both numbers from *each* series' own labels. Label only
  the first series — the shape the plot builders' own worked example shows — and from series
  1 on the part carried `Sheet1!$A$2:$$1` for the categories, a `<c:val>` range running
  backwards over series A's column, and a `<c:tx>` naming series A's header. The workbook was
  right throughout. One `sheetLayout(data)` now serves the chart XML and the workbook
  builder, so the two sides cannot drift; a series carrying no labels of its own states no
  `<c:cat>` rather than an empty one. Reaches the category-axis, stock and surface families.

- **`headerRow` styling is no longer reapplied to a body row on every auto-paged page.** The
  sugar is applied to row 0 at definition time, and `headerRow` was carried onto each
  continuation page, where the recursive `addTable` re-ran it against that page's row 0. With
  the default repeat-header off, that painted rows 11, 22 and 33 of a 40-row table bold and
  filled, and every page emitted `firstRow="1"`. `hasHeader` now follows the same rule: true
  on the first page, and on a later page only when the header row is genuinely repeated onto
  it. `columns` still carries, because it is positional and re-applying it per page is
  correct.

- **A repeated header row costs the page budget it occupies.** The pager priced it from
  `cell._lineHeight`, which is written only onto its own working cells — the repeated rows are
  the definer's plain cells, so the accumulation was always zero and each continuation page
  took the header for free, then packed the same number of body rows the first page fits. A
  40-row table at 18pt paged 15/15/10 without a repeat header and 15/16/11 with one, the last
  row hanging off the slide. The row is now priced the way a body row is, once rather than per
  page.

- **Every page of a paged table gets the same usable height.** There were three rules for two
  cases: the "after the initial slide" block was gated on `tableRowSlides.length > 1`, so
  despite its own comment it began on the *third* page, its first arm recomputed what had just
  been computed, and the explicit-`h` floor reached page one and pages three and up. Page two
  was the one page that got neither, so a 60-row table with `h: 4` and `y: 0.2` paged
  9/**7**/9/9/9/9/8. `calcSlideTabH` is now one rule: this page's start-Y, less the bottom
  margin, floored at `h`.

- **A `colspan` or a cell `margin` is read the same way everywhere.** Six sites read a span in
  six spellings and five read a cell margin in five, agreeing only on values that were already
  valid — and the auto-pager gated each margin side on truthiness, so a cell asking for
  `margin: [0, …]` fell through to the *table's* margin instead of taking its own zero, and a
  scalar `margin: 0.2` was not seen at all where the emitter broadcasts it to four sides. Spans
  now read through `resolveSpan` (already the declared rule) and margins through the new
  `resolveCellMarginsInches`, which reports an unusable one under `table/invalid-margin`
  instead of swapping in the default in silence.

- **`tableToSlides` checks its spans.** `MAX_TABLE_SPAN` guards the two allocations a span
  decides and its own module says it covers both paths, but `tableToSlides` calls
  `getSlidesForTableRows` directly, without the definer that applies the check, and its own
  attribute reader has no ceiling. So `colspan="4000000000"` in a source table reached the
  per-column depth array — the allocation V8 aborts on, with no exception to catch. The check
  now runs at the top of `getSlidesForTableRows`, so the one entry point is guarded whoever
  calls it.

- **Axis units are emitted per axis TYPE, so a chart carrying them opens.** `<c:catAx>`
  (`CT_CatAx`) has no `majorUnit`/`minorUnit` slot at all and `<c:serAx>` (`CT_SerAx`) has
  none of the five units, yet the emitters appended whichever the caller named as the last
  child of all three axes. The validator reported
  `Sch_InvalidElementContentExpectingComplex` on the category and series axes, which
  PowerPoint reports as a file needing repair. The date axis was invalid too, for a
  different reason: `CT_DateAx` does hold all five, but interleaved as `baseTimeUnit,
  majorUnit, majorTimeUnit, minorUnit, minorTimeUnit`, and the code wrote the three time
  units before the two numeric ones.

  The three tags the category builder emits now each get what their content model has:
  none on `<c:catAx>`, the numeric pair on `<c:valAx>` (the tag a scatter or bubble X axis
  takes, where `catAxisMajorUnit` is meaningful and was previously legal by accident), all
  five in schema order on `<c:dateAx>`. `<c:serAx>` emits none. A unit option named on an
  axis with no slot for it now warns under `chart/option-not-on-axis` instead of being
  silently written or silently dropped.

  A regression test had pinned the invalid bytes as correct, which is why this survived; the
  line is now held by `test/schema-cases.js`, where the validator rather than a hand-read
  assertion decides.

- **Every `sz` in the package is bounded by `ST_TextFontSize`.** `clampFontSizeSz` bounds a
  font size to 100..400000 hundredths, and six of the nine sites that write `sz` called the
  bare converter instead — so `addText({ fontSize: 99999 })` warned and emitted `400000`
  while `defineSlideMaster({ textStyles: { body: [{ fontSize: 99999 }] } })` silently emitted
  `9999900`, and one chart could carry both readings (`dataLabelFontSize: 5000` corrected,
  `catAxisLabelFontSize: 5000` raw). All nine now go through the clamp, and the warning names
  the option the caller actually spelled. The master's `textStyles` margins are clamped the
  same way, into `ST_TextMargin` and `ST_TextIndent`.

- **A rejected `indentLevel` is rejected for the bullet margin too.** The level was validated
  for `a:p/@lvl` and then multiplied into the bullet arm's default `marL` regardless, so
  `indentLevel: 1e6` warned that the value was being ignored and emitted
  `marL="342900342900"` — about 6700x the `ST_TextMargin` ceiling. The level is resolved once
  now, and both arms of the paragraph margin go through the clamp rather than only the
  explicitly stated one.

- **A preset-geometry adjustment is not computed against a zero extent.** `rectRadius` is
  emitted as a fraction of the shape's shorter side, and a text object with no stated height
  reaches the emitter with `cy === 0` — the renderer rescues a height only for a *line-less*
  text shape, so any line at all put the division back on zero and wrote
  `<a:gd name="adj" fmla="val Infinity"/>`. A guide formula is a plain string in the schema,
  so nothing downstream refused it. A zero or negative divisor now warns under
  `shape/degenerate-extent` and leaves the preset's own handle in place; a non-finite
  `rectRadius` or `arcThicknessRatio` warns under `geometry/invalid-shape-adjust` and is
  dropped.

- **An animation's `shapeIndex` counts shapes, and `p:cNvPr` ids have no gaps.** Four
  `SlideObjectType` members live in a slide's object list without drawing anything — notes,
  table cells, hyperlink definitions and `online` — and both the id allocator and the
  animation resolver counted them. So `addNotes` before the first shape gave every shape an
  id one higher than it emitted, and `shapeIndex: 0` produced a `<p:spTgt spid>` naming no
  shape on the slide, which is the dangling-target corruption the range check exists to
  prevent. Both now run over the objects that render.

- **A glow radius is clamped like the shadow measure beside it.** `a:glow/@rad` is
  `ST_PositiveCoordinate`, the same type as a shadow's `blurRad` and `dist`, and those went
  through the clamping converter while the glow was a bare multiply — so `glow: { size: -5 }`
  emitted `rad="-63500"` and `glow: { size: NaN }` emitted `rad="NaN"`. Out-of-range now
  warns under the new `glow/size-out-of-range`. A shadow's `dir` goes through the guarded
  angle converter for the same reason.

- **A chart area asked for a bare colour is painted, instead of coming out transparent.**
  `plotArea: { fill: 'FF0000' }` and `chartArea: { fill: '00FF00' }` emitted `<a:noFill/>`
  and the requested colour appeared nowhere in the chart part. `copyChartOptions` clones
  every nested options object defensively, and its clone of `fill` was an unconditional
  spread — which turns a string into `{0:'F',1:'F',…}`, an object naming neither a `color`
  nor a `type`, so `isStatedFill` read it as "this fill states nothing" and took the no-fill
  arm. The clone now leaves a string alone (a string needs no defensive copy) and
  `isStatedFill` counts one as stated.

  This was the one fill site of six where the bare-colour spelling did not work, which is
  what made the leniency invisible: five sites painted it, so nothing looked broken until
  the shorthand was written down. See the `FillOption` entry under **Changed**.

- **A converted text box whose `txBox` is spelled `true` stays a text box.**
  `p:cNvSpPr/@txBox` is the only thing separating a text box from an auto shape — the
  substitute test, "it has no preset geometry", is wrong, because PowerPoint gives every text
  box an explicit `<a:prstGeom prst="rect"/>` — and the script converter read it by comparing
  the raw attribute to `'1'`. The attribute is `xsd:boolean`, which admits `true` and `false`
  as well as `1` and `0`, so a deck from a producer that spells it `true` converted with every
  text box turned into an auto shape: different autofit, wrap and resize rules, and no
  fidelity note, because nothing observed the loss. Neither this library's write path nor
  PowerPoint emits anything but `1`, so no fixture could catch it.

  The parser that already handled the other four-form attributes on the read side —
  `Run.bold`, `Run.italic`, `Shape.hidden`, `flipH`/`flipV`, and a dozen more — has moved
  from `read/oxml/dom.ts` to `ooxml/xsd-boolean.ts`, beside the namespace registry and the
  child-sequence tables, for the same reason those live there: the lexical space of a schema
  type is a fact about the schema, and code outside `read/` should not have to take a DOM
  dependency to reach it. The DOM-typed `boolAttr` wrapper stays behind, and `boolValue` is
  re-exported from its old home, so every existing import is unchanged.

- **`dataBorder: { color: '' }` takes the data-point border default rather than black.**
  Three of the four `<a:ln>` builders defaulted their colour on `??`, which only catches
  nullish, so an empty string passed straight through to the colour validator: a
  `color/invalid-value` diagnostic and a `000000` outline, one line after `363636` was named
  as the default for exactly this. `dataBorder` is the reachable one — the chart-area,
  plot-area and table-cell borders all resolve their colour in the definition step before
  the emitter sees it — and the chart-area *builder* already read an empty colour as
  unstated. All four builders share one implementation now, so they can no longer disagree
  about it.

- **Text inside a shape wraps.** `addShape` builds its slide object without a `_bodyProp`
  bag, where `addTextDefinition` always writes one — and the body-property builder read that
  absent bag's missing `wrap` as `false` through a truthiness test, so every autoshape got
  `<a:bodyPr wrap="none">` and its text ran off the shape on a single line. `square` is both
  the `ST_TextWrappingType` the schema defaults to and what PowerPoint's own "Wrap text in
  shape" writes, so the un-decided case is now the default rather than its opposite: only an
  explicit `wrap: false` still turns wrapping off, and `addText`, which always states `wrap`,
  is unaffected.

  Master and layout placeholders built from a `defineSlideMaster` descriptor reached the
  emitter the same way and are corrected with it.

- **No slide size was bounded, and `defineLayout` warned about everything and refused
  nothing.** `p:sldSz/@cx` and `@cy` are `ST_SlideSizeCoordinate`, 914400 to 51206400 EMU —
  one to fifty-six inches — and nothing checked. `defineLayout`'s guard chain was six arms
  and every one a `warn`, so:

  - `defineLayout({ name: 'Badge', width: 0.5, height: 0.5 })` produced **no diagnostic at
    all** (both values are truthy and finite) and emitted a `sldSz` PowerPoint offers to
    repair. `width: -5` is truthy too, so `cx="-4572000"` reached the file.
  - `defineLayout(undefined)` warned `layout/invalid-definition` and then threw a raw
    `TypeError` on the next line's `layout.name` — after a warning describing the very
    input that could not survive, and against `docs/errors.md`'s statement that every
    failure this library raises is a `TsPptxError`.

  Both dimensions now go through the same clamp-and-warn the rest of the library's
  out-of-range options use (`docs/diagnostics.md`), under the new diagnostic code
  `layout/size-out-of-range`, and a non-object argument throws
  `InvalidOptionError('layout/invalid-definition')`. A numeric string is still advice
  rather than an error: it warns and the layout is defined, as before.

- **An embedded font's `typeface` was escaped by the weakest of three escapers.**
  `pptx.embedFont({ typeface })` takes a caller string straight to
  `<p:font typeface="…">` in `presentation.xml`, and the local escaper it went through
  handled `&`, `<`, `>` and `"` and nothing else. So a control character XML 1.0 forbids
  outright — a vertical tab, say — was written verbatim into the package, which is what
  every other emission site in the library strips; and a newline was silently normalised to
  a space by any parser that read the file back, which is what the write side's
  character-reference escaping exists to prevent.

  The three escapers are now one, in a dependency-free `src/xml-escape.ts` that the write
  side, the read/edit side and the shared embedded-font module all reach. The read/edit
  path's `.rels` and `[Content_Types].xml` writers gain the control strip and `&apos;` with
  it.

- **A chart's `dataBorder.color` rejected the `#` spelling every other colour option
  accepts.** The chart definer had a private copy of the six-hex test that also required
  `length === 6`, so `'#4472C4'` was neither a hex colour nor a scheme colour and the
  border silently became the `F9F9F9` fallback. The hex test is now one function
  (`isHexColor`), which strips the hash first, as the rest of the library always has.

- **Four converter policies were honoured by the shape mapper and skipped elsewhere.** Each
  produced a silently different deck with *no fidelity note*, which is why the round trip
  could not see them: it excludes exactly the declared losses, and an undeclared one is
  invisible when both IRs come from the same reader.

  - **An unwritable `schemeClr` token reached the script raw** from the table fill, the cell
    fill and the cell border. Only ten of the seventeen `ST_SchemeColorVal` values survive
    the write path's `clrMap`; the other seven degrade to a hex literal there anyway, so a
    cell filled `<a:schemeClr val="dk1"/>` emitted `fill: { color: 'dk1' }` and the generated
    script warned `color/invalid-value` and painted the cell the default text colour. A
    gradient stop degraded in the other direction — baked correctly, but with no note. All
    seven sites share one ladder now, and `table.fill.schemeToken`,
    `table.cell.fill.schemeToken`, `table.cell.borders.schemeToken` and
    `fill.gradient.schemeToken` join the note vocabulary.
  - **A graphic frame with no resolvable absolute frame emitted no geometry at all.** The
    shape mapper falls back to `resolvedFrame` and says so, because omitting geometry
    produces `x=0 y=0 w=<slide width> h=0` — broken output rather than lossy output. Graphic
    frames had a second mapper that returned `{}` in exactly that case, so a table or chart
    inside a group with an unusable transform was an undeclared loss. Both go through one
    mapper now, and the note's prose no longer reads "takes its geometry from the own" when
    the shape's own transform was the part that survived.
  - **A fully opaque source emitted `transparency: 0`.** `alphaToTransparency` documents
    that fully opaque is `undefined` — the write path emits no `a:alphaModFix` for a zero
    transparency, so the key cannot come back — and one of its five callers implemented it.
    The rule is the function's now. It was invisible only because the round-trip
    canonicaliser drops `transparency: 0` as an implied default, masking it.
  - **One `masterObject` arm returned `null` with no note**, against the invariant its caller
    states. It is unreachable today, so the note is defensive rather than a behaviour change.

- **A diff report on an output with extra slides understated its own slide count**: it
  reported the expected deck's slide count while the loop above it walked the longer of the
  two.

### Fixed

- **`pptx.tableLayout()` and the file disagreed about `cy`.** `cy` is the already-resolved
  EMU table height the auto-pager and the measured-fit pass stamp onto a table's options, and
  only the fit pass read it. So `addTable(rows, { cy })` with no `h` produced a file whose
  rows are pinned and a `tableLayout()` that reported every row auto-height with
  `heightExact: false`. That is the drift the one reading of `rowH` closed, arriving through a
  second option, so the fix is the shared `resolveTableGridEmu` that both measure paths now
  read rather than a third correction. The emitter deliberately keeps its own height: it
  resolves against the *placed* frame, which a layout placeholder can override, and the other
  two only ever see the options.

- **A `colW` entry that is not a width now says so.** `resolveTableColWidthsEmu` fell back to
  the even-distribution width for an unusable slot with no diagnostic, while the analogous
  `rowH` entry has warned for a while under the reasoning that an entry present but unusable
  is something the caller wrote on purpose. `{ colW: [2, NaN] }` now warns
  `table/invalid-col-width`; a *missing* slot stays silent, which is the same line `rowH`
  draws.

### Changed

- **A shape's `width`/`height` setters reject `0`.** They went through a private EMU guard that
  accepted a zero extent and rejected a negative one under its own `coord/negative`; every other
  read-side extent setter already applied `checkPositiveEmu`, whose rule is
  `ST_PositiveCoordinate` — no room for zero or below, and a zero-size shape is a degenerate
  result rather than a small one. `shape.width = 0` now throws `coord/not-positive`.
  `coord/negative` is retired; nothing else raised it. **Migration:** an intentional zero-size
  shape has to be spelled some other way; the `left`/`top` setters are unaffected, since a shape
  may legitimately sit off the left or top edge.

### Fixed

- **`tableToSlides` wrote its working state into the caller's options object.** Every definer
  copies its options bag before touching it, each with a comment saying why; this path did
  not, and resolved its slide margin, column widths, head rows and per-page `y` into the
  object it was handed. So a reused bag changed the next call's output — most visibly `y`,
  left holding the continuation start, which put the second call's *first* table where the
  first call's *second* page began. The auto-pager likewise wrote its resolved `colW` back
  onto the bag and its `autoPageCharWeight` onto the caller's *cell* options; the grid rides
  out on each `TableRowSlide` now (a new optional `colW`), and the weight is projected onto
  the cell the pager measures rather than written onto the one the caller owns.

- **Two serializers normalized the authored model.** `RenderContext.itemOpts` states the
  opposite contract. The text serializer computed `_bodyProp`'s four insets during render,
  although `addTextDefinition` already owns `_bodyProp`; it does now, which also means an
  unusable `margin` throws from `addText` rather than from `toBytes`, naming the call that
  carries it. The slide-number placeholder wrote `align: 'left'` back onto its props with the
  note "other readers rely on it" — the only other reader already defaulted with `??`.

- **Slide-margin resolution had three spellings that already disagreed.** The auto-pager, the
  HTML path and `addTableDefinition` each turned a master or slide margin into `[T,R,B,L]`
  inches. Two gated the master on `typeof !== 'undefined'` and coerced with
  `Number.isFinite(Number(m))`; the HTML one gated on truthiness and tested `Number.isFinite(m)`
  without the coercion. So a master with `_margin: 0` took the master branch in two of them
  and the caller branch in the third, and a master with `_margin: "0.25"` resolved in two and
  was ignored in the third. All three read `resolveSlideMarginsInches` now.

### Fixed

- **Four read getters gave a different answer for the same OOXML depending on which
  accessor reached it.** All four change additively: values that were `null` or wrong become
  right.

  - **A master's or layout's placeholder resolved its runs differently from the same shape
    read as an `AutoShape`.** `Placeholder` built its text frame with no inheritance context,
    so `Run.resolvedSizePt`, `resolvedFontFace`, `resolvedColor` and `resolvedBold` came back
    `null` through `SlideMaster.placeholders` while resolving correctly through
    `SlideMaster.shapes` — two views of one `p:sp` disagreeing about the same run. It also
    re-derived identity from a hard-coded `p:nvSpPr` where `nonVisualCNvPr` accepts any
    `p:nv*Pr`, and geometry through a helper no other caller used. `Placeholder` is now a view
    over an `AutoShape` and forwards all of those. `type` and `idx` stay its own: `idx`
    reports the absent attribute as `null`, where `AutoShape.placeholder.idx` defaults to
    `'0'` as PowerPoint resolves it.
  - **Table style flags accepted only `"1"`.** The `a:tblPr` flags are `xsd:boolean`, so a
    deck writing `<a:tblPr firstRow="true">` (LibreOffice and other producers do) read
    `firstRowHeader` as `false`. The style context's flags then came back all false,
    `cellStyleParts` emitted no `a:firstRow`, and `TableCell.resolvedFill` reported the
    `wholeTbl` shading for a header row PowerPoint paints in the accent colour.
  - **Eight percentage readers dropped the decimal form.** `a:ST_Percentage` is a union, and
    the `-?[0-9]+(\.[0-9]+)?%` member is the only one the Strict profile has, so
    `<a:srcRect l="10%"/>` made `Picture.crop` and `PictureFill.srcRect` report zero — and,
    since picture-fill crop carries into the script IR, made the converter write a crop of
    nothing.
  - **A table cell's text frame could not resolve a hyperlink.** It was built with no
    relationships, although `TableCell` receives them and uses them for `pictureFill`, so a
    cell run carrying `<a:hlinkClick r:id="rId5"/>` reported `url: null` while the identical
    run in a text box on the same slide resolved it.

  The two attribute-decoding fixes above are the read-side half of the shared parsers; the
  substrate landed earlier and this is the behaviour it changes.

### Fixed

- **Four chart options the emitters dropped or inverted.** None is reachable from the
  showcase decks, so the byte-identity harness had nothing to say about any of them; each
  carries its own test instead.

  - **Category and series axis units needed a format code to be emitted.** Three sibling
    axes had three rules for the same pair of options: the value axis emitted `majorUnit`
    unconditionally, the category axis only behind `catLabelFormatCode` or an XY chart, and
    the series axis only behind `serLabelFormatCode`. So
    `{ type: 'bar3d', catAxisMajorUnit: 3, valAxisMajorUnit: 4 }` emitted exactly one element.

    Ungating them was the wrong reading and is superseded within this same unreleased cycle
    (see **Axis units are emitted per axis TYPE** above): the three rules were three *content
    models*, not three spellings of one decision, and emitting the units on every axis made
    the category and series axes schema-invalid. What survives of this entry is that the
    numeric units no longer sit behind a *format code* — they sit behind the axis type.
  - **A pie's plot-level label flags were constants.** The per-point `<c:dLbl>` honoured
    `showLabel`/`showPercent`/`showValue`/`showSerName` while the plot-level `<c:dLbls>` wrote
    `showCatName="1" showPercent="1"` whatever the caller said. The constants were masked
    while every point carried its own override, which holds only while the pie has labels — so
    a pie with no `labels` and `{ showPercent: false, showLabel: false }` came back with both
    of those `false`s inverted. Both blocks read one builder now.

    An unlabelled pie was broken in two more ways by the same assumption: it keyed the slice
    count on the label count, so it emitted no `<c:dPt>` at all and both sheet ranges ran
    backwards (`Sheet1!$A$2:$A$1`). The slice count now comes from the values, `<c:cat>` is
    omitted rather than pointed at an empty range, and a series with neither labels nor values
    warns and plots nothing.
  - **The bubble workbook's table range used the column count as a row count.** One embedded
    workbook stated two different extents for the same sheet — `ref="A1:C3"` in
    `xl/tables/table1.xml` against `<dimension ref="A1:C5"/>` in `xl/worksheets/sheet1.xml`.
    Those two and the sheet's own column count now come from one `sheetExtent`. Impact is
    bounded because `<tableParts>` is deliberately never emitted, so Excel does not read the
    part; it is still relationship-linked, and the formula was wrong on its face. A bubble
    series with no name also omitted the schema-required `tableColumn/@name` where the
    category branch wrote `''`.

- **Added `valAxisMinorUnit`.** `c:minorUnit` is legal on a `c:valAx` and both sibling axes
  already took a minor unit; only the value axis was missing one.

### Removed

- **`ChartPropsBase.axisPos`.** It was declared and read by nothing — the only `axisPos` in
  the emitters is a local in `makeValAxis` computed from `barDir` and the axis id, and the
  category and series axes hardcode their own placement — so `{ axisPos: 't' }` still emitted
  `<c:axPos val="b"/>`. It was the one fully dead option in a sweep of all 213 names in the
  chart types. **Migration:** delete the option; it never placed anything. Per-axis placement
  would want `catAxisLabelPos`-style naming rather than one key shared across three axes.

### Fixed

- **Six options reached an OOXML attribute unchecked.** `ST_Double`, `ST_Percentage`,
  `ST_Skip` and `ST_TextAnchoringType` all reject a `NaN` or an out-of-vocabulary string, so
  each of these was a package PowerPoint reports as needing repair. All six now follow the
  project's stated rule — warn or throw rather than emit a degenerate result — and none of
  them changes anything for input that was already valid.

  - A **radial gradient's `center`** was a bare `Math.max(0, Math.min(100, …))`, which
    propagates `NaN` into `<a:fillToRect l="NaN"/>` while the stops on the same object already
    threw. It goes through `clampRangedInput` now: a non-number throws `percent/non-finite`,
    an out-of-range one clamps and warns under the new `gradient/center-out-of-range`.
  - **Chart axis crossings** were guarded by `typeof x === 'number'`, the one numeric guard
    `NaN` passes, so `valAxisCrossesAt: NaN` emitted `<c:crossesAt val="NaN"/>`. Both axes now
    share one `axisCrossing` helper that falls back to the axis' default rule and warns.
  - **A chart's `x`/`y` were cast from `Coord` to `number`** on the way into the title layout,
    so `x: '10%'` made `+` concatenate instead of add and `<c:x val="NaN"/>` came out. The
    cast is gone. A percentage still cannot be resolved there — the chart part is built
    without a `PresLayout` — so the chart's own offset is left out of the fold with a warning,
    and the `titlePos` the caller stated is still honoured.
  - **`catAxisLabelFrequency` / `serAxisLabelFrequency`** were typed as free-form strings and
    emitted verbatim into `<c:tickLblSkip val>`, which is an `xsd:unsignedInt` of at least 1,
    so `'every other'` reached the part. Both are now `number | string`, so the natural `2`
    typechecks (it did not before), and anything that is not a whole number of at least 1
    warns and emits nothing. Pass a number; the `string` half is a compatibility hangover.
  - **`valign` had three acceptors that disagreed**, and one of them let any unrecognised
    string through verbatim into `anchor=`. One `resolveTextAnchor` now serves all three,
    covering every spelling they accepted between them (`t`/`top`, `b`/`btm`/`bottom`,
    `c`/`ctr`/`center`/`m`/`middle`) and warning under the new `text/invalid-valign`
    otherwise. Typed callers are unaffected: `VAlign` is `'top' | 'middle' | 'bottom'`.
  - **`tableToSlides` read a cell's CSS `font-size` in px and wrote it as points**, so a
    default `16px` cell emitted `sz="1600"` — 16pt, a third larger than the 12pt the browser
    rendered. The sibling padding read a few lines below had already been corrected to
    `DEFAULT_PX_PER_INCH` and left a note saying the same stale "px to pt 1:1" assumption had
    been there; the font size was not corrected with it. It now converts at the same density,
    and a size with no absolute magnitude (`em`, `%`, a keyword) leaves the key off the cell
    rather than writing `NaN`.

    **Migration:** this changes every HTML-converted table. Text comes out at the size the
    browser rendered rather than a third larger, and because the auto-pager prices rows off
    `fontSize`, a paged table's break positions move with it. To keep the old sizes, set the
    cell font sizes explicitly on the table options rather than relying on the CSS read.

- **Three inheritance paths tested truthiness where they meant "the caller said nothing".**
  An explicit `false`, `0` or `''` counted as unset and was replaced by the value it was
  written to override. All three now test `=== undefined`.

  - **Run options.** `genXmlTextRun` is handed a run's own options and never the shape's, so
    what the caller stated on the shape is copied down onto each run. That copy walked every
    key of the shape's bag and took it whenever the run's value was falsy, so
    `addText([{ text: 'a', options: { bold: false } }], { bold: true })` emitted `b="1"`, and
    a run's `transparency: 0` took the shape's. The copy is now an explicit list of the keys
    `<a:rPr>` actually reads, each applied with `??`.
  - **Table cell options.** The guard was `table[name] && !cell[name] && cell[name] !== 0` —
    the `!== 0` arm rescued zero and nothing else, so a cell's `bold: false` was still
    overwritten by the table's `bold: true`.
  - **Chart data-label text.** `<c:dLbls>`' `<a:defRPr>` had two builders, identical but for
    the operator: `||` at the chart level against `??` per series. One chart could therefore
    carry both readings of one option — `dataLabelFontSize: 0` emitted `sz="0"` beside
    `sz="1200"`. There is now one builder, on `??`, and the size goes through the same
    `ST_TextFontSize` clamp every other font size does, so an explicit `0` is corrected to
    the 1pt minimum with a `font/size-out-of-range` warning instead of reaching the part
    outside its own type.

- **One run's font size no longer leaks onto the runs after it.** The `endParaRPr` bookkeeping
  wrote the first sized run's size back onto the *shape's* options bag, and the inheritance
  above then handed it to every later run and every later paragraph:
  `addText([{ text: 'big', options: { fontSize: 40 } }, { text: 'normal' }], {})` emitted
  `sz="4000"` on both runs. The size is now a paragraph-local. Two showcase decks change:
  five paragraphs whose `<a:endParaRPr>` carried an earlier paragraph's size now carry their
  own.

  A related question is deliberately left open. For string text (`addText('hi', opts)`) a
  run's `options` *is* the shape's object rather than a copy, which is why a shape-level
  `shadow` also becomes a glyph shadow. Naming the run-inheritable keys makes the array-text
  path agree with that path rather than diverge from it; whether `shadow` on an `addText` bag
  should mean the shape, the glyphs, or both is an API decision and is not made here.

- **An SVG image with a hyperlink emitted two relationships with the same id.** An SVG
  picture consumes two rels — the PNG fallback and the SVG itself — and `addImage`'s
  hyperlink then took the second of those ids a third time by incrementing the image's own.
  Both land in one `slideN.xml.rels`, so the part carried a duplicate `Relationship Id` and
  the picture's `r:embed` resolved to the hyperlink. PowerPoint reports the package as
  needing repair, and the OOXML oracle rejects it outright with `PackageOpenError`. Every
  hyperlink rel now allocates through `getNewRelId`, which skips every id the slide already
  holds.

- **A table's black text default was decided by a substring search.** A hyperlink anywhere
  in the grid stands the default down, because the default is direct formatting and would
  paint the words *after* a link black. The test was
  `JSON.stringify({ arrRows }).includes('hyperlink')`, so a cell reading "see the hyperlink
  docs" suppressed the default for the whole table and its text fell through to the theme's
  `tx1` — and every `addTable` serialized its entire grid to ask. It is now a walk over the
  cells, testing the cell's own `hyperlink` and the runs of a cell whose `text` is a run
  array.

- **A scalar `colW` is no longer floored to whole inches.** `colW: 2.4` on a three-column
  table emitted 7 inches of grid rather than 7.2, discarding up to a full inch; `colW` is
  documented as inches with no rounding rule, and the only rounding a length needs already
  happens in `inch2Emu`. The one-element `colW: [3]` form read its width by coercing the
  whole array rather than reading `colW[0]`, which is correct only by accident of
  array-to-primitive coercion. A `colW` that is not a positive number now warns under the new
  `table/invalid-col-width` and falls back to the default table width; it used to become
  `w: NaN` and surface far downstream as `coord/non-finite`, whose message describes a
  missing layout dimension and names nothing the caller wrote.

  **Migration:** a table sized by a fractional scalar `colW` gets slightly wider — the width
  you asked for. Pass the floored value explicitly to keep the old geometry.

- **The auto-pager's width arithmetic.** Three defects in `getSlidesForTableRows`, all in
  the same twenty lines, all producing wrong output rather than an error. No showcase deck
  auto-pages, so the byte-identity harness says nothing about any of them; each is covered
  by a case in `test/regression/table/table-autopage-width.test.js` that fails against the
  old code.

  - **Usable width was the sum of the slide margins, not the slide minus them.** The
    fallback the pager reaches when a table states neither `w` nor `colW` *added* the left
    and right margins, so a 10in slide gave about one inch of usable width and the pager
    then divided that across the columns. It writes the result back onto `colW`, so the
    number reached the emitted `<a:gridCol>`: a three-column table added with
    `{ autoPage: true, colW: [1, 2] }` — the count mismatch drops `colW` and sets no `w` —
    emitted three columns a third of an inch wide, against 3 inches each for the same table
    without `autoPage`. The width is now the slide less the table's own left edge and the
    right margin, which is what the height path and `addTableDefinition` already computed.
  - **A spanning cell was measured against the wrong columns.** The colspan filter read
    `idx >= iCell && idx < idx + cellColspan`, and the second half is true for every
    positive span, so a colspan-2 cell in a five-column table was priced at the full width
    of every column from its own position onward. Cells were also indexed by their position
    in the row rather than by grid column, so every cell after a colspan measured against a
    neighbour's width, and a rowspan opened in an earlier row shifted them again. Both are
    now resolved through a column cursor with the same placement rule the rowspan
    bookkeeping and `walkTableGrid` apply. A cell that wraps to fewer lines than it should
    prices its row short, so the page over-fills.
  - **A row longer than the first row threw.** The same sum used `Array.prototype.reduce`
    with no seed, so a table whose later row has more cells than row 0 defines columns
    failed with `Reduce of empty array with no initial value` instead of paging.

  Column widths now come from `resolveTableColWidthsEmu`, the resolver the table emitter and
  the measured-fit pass already share, so the widths the pager wraps text against are the
  widths the package carries. That also normalizes a short `colW` array to the column count;
  it used to leave the trailing columns measuring against a zero width.

  **Migration:** an auto-paged table that was relying on the old fallback width — that is,
  one with no `w` and no usable `colW` — will now be laid out across the full usable slide
  width and may paginate differently. State `w` (or a `colW` whose length matches the column
  count) to pin the old geometry.

- **`shadow.transparency` and `shadow.angle` follow the project's own out-of-range rule.**
  Both were the third option `docs/diagnostics.md` names as the tempting wrong one: they
  reported a warning *and* discarded the request, so the caller read a diagnostic and got a
  shadow they had not asked for. Four behaviours change, and none of them for an input that
  was already in range — every showcase deck emits byte-identically:

  - `transparency: 120` now clamps to `100` and paints fully transparent, warning
    `shadow/transparency-out-of-range`. It used to warn and leave the derived alpha unset,
    which fell through to the emitter's 0.75 default: a shadow at 25% opacity nobody
    requested.
  - `transparency: NaN` now throws `percent/non-finite`, as every other percentage option
    does. It used to warn and paint the same 0.75 default.
  - `angle: 400` now clamps to `359` (`dir="21540000"`), warning `shadow/angle-out-of-range`.
    It used to become `270` — not the nearest bound, and not the `@default 0` the option
    documents, but `DEF_TEXT_SHADOW.angle`, so a shape's shadow silently borrowed a text
    shadow's direction. A negative angle clamps to `0` for the same reason. The angle is
    clamped rather than wrapped modulo 360: `ST_PositiveFixedAngle` has a nearest legal
    neighbour like any other range, and the rule as written is "clamp to the nearest bound".
  - `angle: NaN` now throws the new `shadow/angle-non-finite`. It used to reach the package
    as `dir="NaN"`: the guard was written `if (corrected.angle)`, and `NaN` is falsy, so the
    range check it was standing in front of never ran.

  **Migration:** an untyped caller passing a numeric string (`angle: '45'`,
  `transparency: '40'`) now gets a throw where the value used to be coerced. Both options
  are typed `number`, and this is the same policy every clamped percentage already applies;
  convert before the call.

  `shadow.blur` and `shadow.offset` stay lenient about a value that is not a number at all
  (it collapses the feature to `0`); their range is the entry below.

- **`shadow.blur` and `shadow.offset` are bounded by the schema, and their JSDoc now says
  what is actually enforced.** Both are `ST_PositiveCoordinate` on `blurRad`/`dist`, which
  is *unsigned* — so `blur: -6` reached the package as `blurRad="-76200"` and PowerPoint
  reported the file as needing repair. A negative value now clamps to `0` and warns under
  the new `shadow/blur-out-of-range` / `shadow/offset-out-of-range`.

  The doc'd `range: 0-100` and `range: 0-200` are **not** what moved, and are gone from the
  JSDoc rather than enforced. Those two numbers are the limits of PowerPoint's own spinners,
  not of the format: a 150pt blur loads and paints, so clamping to 100 would have discarded
  a legitimate request to enforce a bound nothing was checking. This is the same call
  `clampLineSpacingMultiplePct` already makes — clamp to the schema's range, not to the one
  the option's prose describes, because only the values PowerPoint reports as needing repair
  are worth moving. The two bounds are now documented as what they are.

- **A bubble series' outline reads `chartColors` and `lineCap` the way every other plot
  does.** `plot-bubble.ts` built its `<a:ln>` differently from `plot-scatter.ts` and
  `plot-cat-axis.ts` in two ways that nothing explained, and both were omissions:

  - The stroke colour went through `genXmlColorSelection` directly rather than
    `chartColorLineFill`, so `chartColors: ['transparent']` — which means an invisible
    series — passed the literal `'transparent'` into colour validation, warned "not a valid
    scheme color or hex RGB", and painted the bubble outline **black**. That is precisely
    the hole `chartColorLineFill` exists to close; it had been added to the other two plot
    families and missed here.
  - The cap was hardcoded `flat`, so `addChart('bubble', …, { lineCap: 'round' })` accepted
    the option and silently dropped it — on both arms of the branch, the palette line and
    the `dataBorder` one. It is not cosmetic: a bubble outline carries `lineDash`, and the
    cap shapes the end of every dash in it.

  Bytes move only for a chart that sets one of those two options; `lineCap` left unset still
  resolves to the `flat` that was hardcoded, and every showcase deck emits byte-identically.

- **A `p:spPr` created on a shape that has a `p:extLst` is inserted before it, not after.**
  `SHAPE_AFTER_SPPR` was the one successor list in `src/ooxml/sequence.ts` written out by
  hand instead of sliced from a declared sequence, and it was missing `p:extLst` — the last
  child of `CT_Shape`. A `getOrAddChild(sp, 'p:spPr', …)` on such a shape therefore matched
  no successor and *appended*, producing a schema-invalid `p:sp`. Deriving the list from a
  declared `SP_SEQUENCE` closed the gap and removed the module's one exception to "every
  successor list is sliced out of a sequence" in the same edit. Reachable only on input that
  is already malformed — the schema makes `p:spPr` required — so this is the repair path
  no longer replacing one invalidity with another.

- **A chart's embedded workbook stamps one time, not two.** Its `docProps/core.xml` read
  the clock separately for `created` and `modified`, so a build that crossed a millisecond
  gave the two different values — nondeterministic output for a difference no reader acts
  on. Both now come from one reading, and it drops the milliseconds the deck's own
  `docProps/core.xml` has always dropped, so the two parts agree on the precision Office
  writes.

- **A non-finite value no longer reaches the embedded workbook either, for any chart
  family.** The fix above cleaned the chart's own cache; the workbook the chart is backed by
  is written by a separate builder, and that one still wrote `<v>Infinity</v>` (or `NaN`) into
  the cell. Excel refuses such a workbook outright, with 0x3EC. PowerPoint hides it, because it
  does not parse the embedding when it opens the deck — the deck opens and paints from the
  cache, and the failure surfaces only when a user picks **Edit Data**, at which point the
  chart looks fine and the workbook behind it will not load. A non-finite number now leaves the
  same empty cell a `null` value always has, which Excel reads back as an empty cell. Nothing
  warns on this side: every numeric cell in the sheet is mirrored by a cache point, and that is
  where `chart/non-finite-value` is already reported for the same value.

- **A non-finite value in a pie or doughnut series no longer produces a deck PowerPoint
  refuses to open.** Every other chart family cached its values through one builder, which
  warns on a non-finite number and leaves the point out; pie and doughnut built their own
  `<c:val>` and wrote whatever they were handed, so `values: [10, Infinity, 38]` reached the
  package as `<c:v>Infinity</c:v>` and PowerPoint rejected the file with 0x80070570, the
  corrupt-file error. `NaN` and `INF` in that position do the same. A pie now goes through the
  shared builder, so a non-finite value warns (`chart/non-finite-value`) and is dropped, as it
  always has been on a bar or a line.

  A *gap* — a `null` or `undefined` slice value — also changes spelling, from a `<c:pt>` with
  an empty `<c:v>` to no `<c:pt>` at all. That half is not a defect being fixed: measured
  against desktop PowerPoint the two are equivalent, opening without a prompt, resolving to the
  same object model and exporting to a byte-identical image. It changes because one spelling of
  a gap is better than two. Nothing here is visible to the schema — `<c:v>` is `s:ST_Xstring`,
  so the OpenXmlValidator passes all four spellings alike, which is why PowerPoint is the
  oracle this was settled against.

  **Migration:** an untyped caller passing a *numeric string* (`values: ['42', '7']`) to a pie
  or doughnut now gets `chart/non-finite-value` and an empty chart, where the value used to
  reach the cache and render. That is what every other chart type has always done with one —
  the option is `number[]` — so convert before the call rather than relying on the pie path.

- **A chart `x`/`y` spelled as a percentage or a unit string is honoured instead of being
  replaced by 1 inch.** Both are ordinary `Coord` values, and the guard in front of them
  was `!isNaN(Number(x))` — false for `'50%'` and `'2in'` alike, so the value was thrown
  away and the default substituted. A chart positioned at half the slide width landed an
  inch from the left edge, silently. The guard now defaults only what the caller omitted
  and lets the coordinate converter vet the rest, which also means a `NaN` position is
  reported (`coord/non-finite`) rather than quietly becoming 1 inch.

- **The coercing numeric globals are gone from `src/`, and banned.** `isNaN` and
  `isFinite` coerce their argument before testing it, so they answer a different question
  from the one at the call site: `Number('') === 0` makes `isNaN('')` false, and
  `isNaN(Infinity)` is false, so a value with no finite representation passes a guard whose
  job was to stop one. 39 uses of the former and 18 of the latter are now `Number.isFinite`
  (or `Number.isNaN`, where an infinity is deliberately clamped rather than rejected), and
  `.oxlintrc.jsonc` bans both under `no-restricted-globals` for `src/**` so the sweep
  cannot decay. What changed beyond consistency:

  - **`indentLevel` is range-checked.** It is written straight into `a:p/@lvl` with no
    converter in between, so `Infinity` reached the package as `lvl="Infinity"`. It is
    `ST_TextIndentLevelType` (0-8, whole numbers), and anything else now reports the new
    `text/invalid-indent-level` and is ignored.
  - **`autoPageHeaderRows` is checked as a count**: a whole number from 1 to the table's own
    row count. `''` used to read as `0` header rows, a fraction survived, and `Infinity` was
    accepted; all three now report the new `table/invalid-header-row-count` and fall back to
    `1`. Its sibling `autoPageLineWeight` had been range-clamped since it was written.
  - **Chart border widths go through `lineWidthToEmu`**, the ST_LineWidth clamp shape
    strokes have always used. A negative `plotArea.border.width` or `dataBorder.width` used
    to reach `a:ln/@w` as a negative attribute; a negative or non-finite one is now not a
    width at all and takes the documented default.
  - A non-finite `lineDataSymbolLineSize` used to collapse to a zero-width marker outline.
    It goes through the same clamp.

- **`pptx.tableLayout()` no longer disagrees with the file the export writes about row
  heights, and its numbers change.** `rowH` was read independently in four places: the
  `<a:tr h>` emitter, the export-time measured-fit pass, the auto-pager, and
  `tableLayout()` itself. `rowH: [0, 2]` on a two-row table with `h: 4` baked a 2.0in
  first row (`0` is falsy, so the writer fell through to the even split of `h`) while
  `tableLayout()` returned **0.2004in** for it and flagged `heightExact: true` — a tenfold
  error on the number a caller places the next shape against, reported as pinned. A
  negative entry reached the file as `<a:tr h="-914400">`, and a stringified one
  (`rowH: ['1']`, reachable from untyped JS) was honoured by the writer and rejected by
  `tableLayout()`. `docs/measured-text-fit.md` states the invariant this broke: a
  layout-time prediction must never disagree with what the export then bakes.

  All four now call `resolveTableRowHeightEmu`, beside the `resolveTableColWidthsEmu` the
  column side has always shared. **An entry pins its row only when it is a number greater
  than zero.** `0`, a negative, and anything that does not read as a finite number are not
  heights: the row is sized from the table's `h` instead, or grows to fit if there is none,
  and the new `table/invalid-row-height` diagnostic says which entry did that. A *missing*
  array slot stays silent, because that is how the auto-pager spells an auto-height row.
  Separately, `heightExact` is now false whenever the default-line fallback supplied a
  row's height, which is a number `tableLayout()` invented rather than one the file pins.

  **Migration:** read `heightExact` before trusting a height, as the API already asked. If
  you wrote `rowH: [0, …]` expecting a zero-height row, nothing ever gave you one — the
  writer split `h`, and that is now what every path reports. Emitted bytes are unchanged
  for every `rowH` that was already a positive number.

- **`fill: { gradient }` and `fill: { pattern }` paint what they name, instead of a black
  shape.** Which fill kind a props object asks for was answered in seven places and they
  disagreed. The stroke emitter inferred `gradient` from the sub-object; the shared fill
  dispatcher inferred nothing at all, so a `fill` carrying a `gradient` or a `pattern` and
  no `type` fell to the `'solid'` default and emitted a black `<a:solidFill>` — reporting
  it, worse, as `"" is not a valid scheme color or hex RGB!`, which blames a colour string
  the caller never wrote. Three `define/` modules each carried their own copy of the image
  half, the two `ShapeLineProps` rebuilds stamped a `type` on before the emitter could
  infer one (so `line: { pattern }` came out a default-black solid), and the slide
  background gate accepted a `color` or the literal `type: 'gradient'` and silently dropped
  `{ gradient }`, `{ pattern }`, `{ type: 'pattern', … }` and `{ type: 'none' }`.

  `resolveFillKind` is the one answer now, and every one of those sites asks it.
  **A sub-object selects its kind on its own; an explicit `type` beats a sub-object that
  disagrees.** That second half is the rule under which `{ type: 'none', gradient }` can
  still mean transparent, so it is stated rather than left as fallout from ordering, and it
  is now in the `ShapeFillProps.type` doc comment. With no `type`, the first sub-object
  present wins in declaration order — `gradient`, then `pattern`, then `image`.

  **Migration:** nothing to do if you always spelled `type` out. If you relied on a
  sub-object being ignored — `{ type: 'solid', color, gradient }` painting solid — that
  still works, and it is the case explicit-wins protects. What changes under you is a
  `fill`, `line` or `background` carrying only a sub-object: it used to come out black,
  default-stroked, or absent, and now paints what it names. A `background` spelled
  `{ type: 'none' }` now emits `<a:noFill/>` where it previously emitted no `<p:bg>` at
  all; omit the background entirely, or spell `{ type: 'inherit' }`, to keep inheriting the
  master.

- **Percent-valued options are clamped into their schema range instead of reaching the
  attribute raw.** `chartColorsOpacity: 150` wrote `<a:alpha val="150000"/>` against an
  `ST_PositiveFixedPercentage` maximum of 100000 (and `Infinity` wrote `val="Infinity"`);
  image `transparency` inverts, so 150 wrote a *negative* `<a:alphaModFix>` and -30 wrote
  one over the maximum; `biLevel.threshold` is a 0-1 fraction, so 5 wrote
  `thresh="500000"`; and `lineSpacingMultiple: 200` wrote `<a:spcPct val="20000000"/>`
  against a maximum of 13200000. Every one of those makes PowerPoint offer to repair the
  package. All four now clamp and warn, through the same `units-internal.ts` helpers the
  already-correct `transparency` and `opacity` paths were using — plus two that were
  missing, for the non-inverting 0-100 and the 0-1 forms. A `NaN` now throws under the new
  `percent/non-finite` code rather than reaching the attribute as `val="NaN"`; `Infinity`
  clamps to the bound like any other out-of-range number. In-range input is byte-identical.
  New diagnostic: `image/bilevel-threshold-out-of-range`.

## [3.7.0] - 2026-08-30

`groupObjects()` has always addressed objects by `objectName`, and there has never been a
way to learn those names. That is fine while one caller authored everything and useless
the moment a slide is assembled by independent renderers, which is exactly the case
`groupObjects()` was added for: nobody kept the descriptors, so the only handle on what is
on the slide is a name nothing reports. The two ways out were both bad — make every
renderer surrender its internals, or keep a parallel ledger that is wrong the first time a
renderer adds an object it did not announce. `slide.objects` is the accessor that was
missing, and it reports the groupable-kinds rule alongside each object so a consumer does
not pin a second copy of it.

This is also the first release to go out under two names. `pptx-ts` is the same build
published a second time, for people who cannot or will not type a scope; `@shbernal/ts-pptx`
stays canonical, and installing both at once is the one thing not to do.

### Added

- **The package is published under a second, unscoped name: `pptx-ts`.** npm has one
  package per name and no notion of a redirect, so an alias on the registry is not a
  pointer at another package: it is a second publish of the same content under a second
  name. That is what this is. `scripts/alias-package.mjs` stages a copy of the build the
  release gates just proved, and `.github/workflows/publish.yml` publishes that copy after
  the canonical publish has succeeded. Three things differ from `@shbernal/ts-pptx` at
  the same version, and nothing else does: `package.json#name`, a `README.md` banner
  naming the scoped package as the canonical one, and the absence of the `scripts` block,
  which is dev-only metadata in a published manifest and whose `prepack` would delete the
  staged `dist/` if a publish ever ran without `--ignore-scripts`. `version`, `exports`,
  `files`, `dependencies` and the `dist/` and `skills/` payloads are the same bytes, so
  the alias is the package the gates measured rather than a near-copy of it.

  **Install one or the other, never both.** Two copies of this library in one dependency
  tree are two module registries, and per-copy state such as the diagnostic handler stops
  behaving as one library. The scoped name stays canonical: the issue tracker, this
  changelog and every example in the docs use it, and nothing about cutting a release
  changes because the second name exists.

  The name is reversed rather than scope-stripped because `ts-pptx` on npm belongs to an
  unrelated project.

- **`slide.objects`: the authored objects on a slide, bottom-to-top in z-order.** The
  read-back half of `groupObjects()`. Each entry is a `SlideObjectInfo` carrying `type`,
  `objectName`, `isPlaceholder`, `canGroup`, and `children` (a group's members, nested to
  any depth). It is a snapshot rather than a live handle: a fresh array on every access,
  inert, and the way to act on it is to call the authoring API with the names it reports.

  **`objectName` comes back in the caller's spelling, not the stored one.** Every
  `add*Definition` stores the name attribute-escaped and `groupObjects()` escapes the
  caller's spelling before comparing, so reporting the stored form would hand back a name
  that escapes a *second* time on the way in and resolves to nothing — the same mismatch
  1.0.0 fixed on the input side, arriving from the other direction. A shape named `Q&A`
  reads back as `Q&A`. The promise is the round trip rather than invertibility, and it
  holds even where decoding is not a true inverse: a name authored as the literal string
  `&amp;` decodes to `&amp;` and re-encodes to the stored `&amp;amp;`.

  **`objectName` is always a string.** An object authored without one still carries the
  generated `Shape 3` / `Text 1` / `Group 2` identity PowerPoint shows in the Selection
  Pane, and that name addresses it as well as an authored one does — reporting `null` there
  would have withheld a usable handle to signal a distinction nothing downstream of
  authoring records anyway.

  **`canGroup` is the predicate `groupObjects()` throws on**, not a restatement of it.
  `GROUPABLE_TYPES` and the placeholder exclusion now live behind one `isGroupableObject`
  that both the accessor and the throw path call. A consumer that had to re-state the rule
  would be pinning a copy of a list it does not own: it would keep refusing the day
  grouping learns a new kind and go on offering the day one is withdrawn, with nothing
  failing in either direction. It answers about the object and not the call — an
  unresolved, duplicated or ambiguous name still throws, and no single object can speak
  for a selection.

### Changed

- **The publish workflow is re-runnable when only one of the two names went out.**
  Publishing two packages is not atomic, so the canonical publish can succeed and the
  alias fail after it. The unpublished-version guard therefore fails only when **both**
  names already carry the version, and each publish step is skipped when its own name
  already does. Re-dispatching the workflow on the same tag finishes the missing half and
  touches nothing else. Previously the guard failed whenever the version existed under the
  one name it knew about, which would have made `workflow_dispatch` useless in precisely
  the state it exists for. The alias is published last for the same reason: a failure in
  it costs a re-dispatch, never the release that already went out.

## [3.6.0] - 2026-08-28

This release finishes a colour job the read model had been doing half of. `a:prstClr` and
`a:hslClr` resolve, so five of the six DrawingML colour models report a colour wherever one
is read, and the three accessors that were hunting for `a:srgbClr` and `a:schemeClr` by tag
name — a gradient stop, a shadow, a cell border — take the colour element by position
instead and get all of them. A gradient stop and a cell border also carry the whole
`ResolvedColor` now, transform list included, so a consumer can tell "this source stated no
transforms" from "this reader could not see them", and `Run.resolvedItalic` closes the last
of the inherited character properties. Three definers stop writing their own state onto the
options object you hand them: 3.5.0 named `addText`, but `addShape` and `addTable` did it
too, which is what made one style literal spread across three shapes come out sharing a
name and silently losing its paging. Two allocations sized straight from numbers nothing
checked — a chart's cached point count, a table cell's spans — aborted the host process
with no exception to catch, and are bounded now. And `test:com` gains a sibling:
LibreOffice has no SmartArt layout engine, so it paints the drawing cache PowerPoint
recomputes, which is the only way to show 3.5.0's mirror reaching a pixel rather than a
part.

### Added

- **A gradient stop reads every colour model the reader resolves, and reports a preset
  name** (`ts-pptx/read`). `readGradientStops` hunted for `a:srgbClr` and `a:schemeClr` by
  tag name, so a stop written as `a:prstClr`, `a:sysClr` or `a:hslClr` came back blank in
  every field — `color`, `schemeColor`, `effectiveHex` all `null` — even though
  `resolveColorElement` resolves five of the six models everywhere else, and even though
  the `a:prstClr`/`a:hslClr` work above claims gradients among the places it reaches. It
  did not reach this one.

  `a:CT_GradientStop` is a sequence of exactly one `a:EG_ColorChoice`, so the stop's colour
  is whichever element sits in that slot; it is read by position now rather than by name,
  and every model the resolver handles arrives.

  **`GradientStop.presetColor` now exists.** Three places already documented it — the
  `presetColorHex` export note above, and the `RecolorColor` cross-references that describe
  themselves as mirroring "the {@link GradientStop} split
  (`color`/`schemeColor`/`presetColor`)" — against a type that had no such field. It is the
  raw `a:prstClr/@val`, matching the raw/resolved split the rest of the read model uses. A
  stop written as `a:sysClr` or `a:hslClr` has no raw field of its own and is reported
  through `resolvedColor` alone, which the interface now states rather than leaving to be
  inferred from three `null`s.

  Downstream, the script converter drops a stop it cannot resolve and falls back to no
  gradient below two stops, so a deck whose gradient used any of these models silently lost
  it; those stops now convert.

- **`GradientStop.resolvedColor` and `CellBorder.resolvedColor`** (`ts-pptx/read`). A
  `ResolvedColor` keeps three things: the base `hex`, the raw `transforms` list, and the
  `effectiveHex` after applying them. Two other places read a colour through the same
  resolver and kept only the last two, so the transform list was computed and then dropped
  on the floor. A gradient stop reported `position`/`color`/`schemeColor`/`effectiveHex`
  and a table cell border reported a resolved hex plus a raw `schemeColor`, with nothing
  between them: on a deck whose accent carries `<a:lumMod val="75000"/>`, all three of a
  solid fill, a gradient stop and a cell border resolved to the same `2F5597`, and only the
  solid fill could say what the transform had been.

  Why it matters to a consumer: a reader could not tell "this stop stated no transforms"
  from "this reader could not see them". Anything re-authoring a deck against a *different*
  theme needs the first fact, because `effectiveHex` alone is the theme baked in, and a
  `lumMod`-darkened accent carried forward as a literal hex stops tracking the theme it
  came from with nothing saying so.

  Both now carry the full `ResolvedColor` the resolver already built, under `resolvedColor`.
  The existing flat fields are unchanged and stay the convenience for painting:
  `GradientStop.effectiveHex` and `CellBorder.color` are both exactly
  `resolvedColor?.effectiveHex ?? null`. An empty `transforms` means the source stated
  none; a `null` `resolvedColor` means there was no resolvable colour to read. Nothing has
  to migrate.

  `ResolvedColor.transforms` is now typed as the already-exported `ColorTransform[]`
  rather than a structurally identical inline literal, so the three places that hand one
  around name the same type.

- **`Run.resolvedItalic`** (`ts-pptx/read`). `Run` resolved four character properties
  through the placeholder / list-style / master chain — colour, size, face and `@b` — and
  not `@i`, so a run inside a placeholder that inherits `i="1"` from the master's
  `p:txStyles` reported `italic: null` with no way to answer what it actually renders as.
  `@b` and `@i` are siblings on `CT_TextCharacterProperties` and PowerPoint writes them as
  a pair (every level `a:defRPr` in a stock master carries both); the write API already
  treats them as one, since `MasterTextStyleLevel` has `bold` and `italic` — so a deck
  could be *authored* with an inherited italic that could not then be read back. Anything
  painting from the read model as `props.X ?? resolved.X` had to render such a subtitle
  upright. It resolves the same chain `resolvedBold` documents (paragraph `a:defRPr` →
  slide `a:lstStyle` → layout → master placeholder `a:lstStyle` → master `p:txStyles` →
  `p:defaultTextStyle`) and reports `null` when the run states no `@i` and inherits none.

  Both flags now share one walker, which corrects a smaller thing on the way past: an
  unparseable `@b` (`b="yes"`) used to resolve to `false`, because the old code tested
  `=== '1' || === 'true'` and let everything else fall through as not-bold. It is parsed by
  `boolValue` now — the same `xsd:boolean` parser `Run.bold`/`Run.italic` use — so a value
  the schema does not allow reports `null` (unknown) rather than a confident `false`.

- **`a:prstClr` and `a:hslClr` resolve, so five of the six DrawingML colour models now
  report a colour** (`ts-pptx/read`). `resolveColor` handled `a:srgbClr`, `a:sysClr` and
  `a:schemeClr`; a colour written in any of the other three read as `null` everywhere in
  the read model — fills, lines, gradients, effects, table styles, slide backgrounds and
  the `theme: 'preserve'` flatten path alike. `a:prstClr` is the one that mattered, because
  this library emits it itself (a notes-page frame, a zoom tile's border) and because
  `Picture`'s duotone stops already surfaced the raw preset name with nothing to turn it
  into. `src/read/oxml/preset-color.ts` now holds the ECMA-376 §20.1.10.47 table: 190
  enumerated names, 140 distinct colours, since the abbreviated `dk`/`lt`/`med` prefixes,
  both spellings of grey, and case are spelling rules rather than separate entries.
  `a:hslClr` went in beside it through the sRGB-HSL conversion the colour-transform module
  already carries and already validates against PowerPoint for `lumMod`/`satMod`. The
  preset resolves to a hex and nothing else: `ResolvedColor` gains no `presetColor` field,
  because the read model reports the raw reference separately from the resolved one
  everywhere else, and the write API has no preset-name option for a round trip to reach.

  **`presetColorHex` is exported from `ts-pptx/read`**, so a caller holding a raw preset
  name — `GradientStop.presetColor`, `RecolorColor.presetColor` — can make it literal the
  same way the reader does.

  **`a:scrgbClr` is deliberately still unresolved.** Its channels are percentages of a
  colour space the schema does not pin down, so whether `50%` is linear-light or
  sRGB-encoded decides the answer and the two differ by a gamma curve. Reporting no colour
  is honest; reporting a guessed one is not. Settling it needs a render oracle, and until
  there is one that stays a decision rather than an oversight.

  One related repair on the way past: `a:ST_Percentage` is a *union* in the Transitional
  profile — the fixed-point integer Office writes and a `%`-suffixed decimal string — and
  the transform decoder read only the first, dropping a schema-legal `alpha="50%"` without
  a word. Both spellings are read now.

  It shows up downstream: three run colours in the `table` fixture move from
  `text.color.default` — the one fidelity note where the converter's output colour is not
  merely frozen but possibly *wrong*, because nothing resolved what the run inherits and
  the write path paints it black — to `text.color.inherited`, resolved and baked. The
  corpus note totals move with them (733 / 433).

- **`objectLock` on zooms, and honoured on 3D models** (`addSlideZoom` /
  `addSectionZoom` / `addSummaryZoom` / `addModel3d`). `SlideZoomProps` had no such field
  at all, while `Model3dProps` inherited one that the definer never propagated — so a
  zoom's locks could not be expressed and a model's were silently dropped. Both now land
  on the `a:graphicFrameLocks` of the `mc:Choice` graphic frame, which is the object
  PowerPoint 2016+/2019+ actually selects, moves and locks, and take the graphic-frame flag
  set (`noGrp`, `noDrilldown`, `noSelect`, `noChangeAspect`, `noMove`, `noResize`). A zoom
  keeps `noChangeAspect` on by default and `objectLock: { noChangeAspect: false }` lifts
  it, as on an image. The `mc:Fallback` picture in both keeps its own fixed `a:picLocks`
  set: it is only what a pre-2016 consumer draws, and folding a graphic-frame flag set onto
  a picture-locks element would warn about every flag the two element types do not share.

- **A second render oracle: `pnpm run test:lo`.** PowerPoint cannot be an oracle for
  markup PowerPoint recomputes, and SmartArt is the case that proves it. A deck stores
  every drawn string twice, in the `dgm:dataModel` PowerPoint reads and in the
  `dsp:drawing` cache every renderer without a layout engine paints, and PowerPoint
  rebuilds the cache from the data model on open. So a deck whose cache was never
  written and one whose cache was written correctly render identically in PowerPoint,
  `test:com` cannot separate them, and the drawing-cache mirror shipped in 3.5.0 was
  proven only by the bytes of `ppt/diagrams/drawing1.xml` — real evidence, but evidence
  about a part rather than about a pixel. The new gate renders through LibreOffice,
  which has no SmartArt layout engine and therefore paints the cache and nothing else,
  via `soffice --convert-to pdf` plus `pdftotext`. Three cases, each with its own
  sensitivity check: `mirrored` proves `DiagramPoint.text` reaches the renderer, `stale`
  edits the data model alone and proves the renderer keeps painting the *old* string,
  and every case asserts the ten sibling nodes it did not touch are still painted, so a
  mirror that writes the right string into the wrong point fails on its own sentinel.
  Text is read through PDF rather than PNG because LibreOffice's PNG export writes the
  first slide only and ignores a `PageRange` filter option. SKIPs cleanly when either
  tool is missing (`TSPPTX_SOFFICE` / `TSPPTX_PDFTOTEXT` override the search, and a
  set-but-wrong path errors rather than silently falling back).

  It covers more than SmartArt, and it runs in CI. Three of the constructs the
  byte-identity corpus never emits — `a:buBlip`, `a:prstTxWarp` and `numCol`/`spcCol` —
  are now asserted against a second implementation rather than against their own bytes,
  each as a construct deck diffed with an otherwise identical control deck and each made
  to fail on purpose before it was kept. `rtl="1"` and `altLang` were probed and dropped
  as unobservable to any renderer; `a:buClr` changes the raster but not the extracted
  text. Unlike `test:com` there is a CI leg (`ubuntu-latest`, ~1m6s off the critical
  path), running under `TSPPTX_RENDER_ORACLE=required` so a runner missing a tool fails
  instead of skipping green.

- **`CommonObjectDescriptor`**, the six key-tagged object descriptors a slide master and a
  group child both accept — `image`, `line`, `rect`, `roundRect`, `shape` and `text`.
  `SlideMasterObject` and `GroupChildProps` are now that type plus what is genuinely their
  own (a master adds `chart` and `placeholder`, a group child adds `group`). Both unions
  accept exactly what they accepted before, so nothing to migrate; the new name is what
  you will see in an editor hover, and it is exported so it can be named directly.

### Fixed

- **`addText` no longer writes its own state onto the options object you hand it.** The
  definer normalized options in place — a `_bodyProp` record, the assigned `objectName`,
  defaulted `color` and `line` — and the emitters then wrote more of it (`_lineIdx`,
  paragraph properties inherited from the shape). Reusing a style literal, the ordinary
  way to give several shapes one look, was enough to corrupt them: the first `addText`
  hung a `_bodyProp` on the literal, every `{ ...STYLE }` after it aliased that same
  record, and one box asking for `columns: 2` silently columnized the boxes before and
  after it. `objectName` leaked by the same route, spreading the first box's assigned
  name onto its siblings and colliding in the Selection Pane. Options are now copied on
  the way in, so the object a caller passes comes back exactly as it was written.
  Emitted bytes are unchanged (183/183 baseline parts): sharing *within* one `addText`
  call is preserved, because the string shorthand's shape and run genuinely are one
  object and the second normalization pass over it emits `<a:ln>` defaults the first
  cannot. Nested option objects (`bullet`, `shadow`, `fill`) are still shared by
  reference, which is how their relationship ids reach the emitters.

- **`addShape` and `addTable` no longer write their own state onto the options object
  you hand them either.** 3.5.0 named `addText` as the entry point that leaked; it was
  not the only one. `addShape` stamped the assigned `objectName` and the normalized
  `line` onto the caller's literal, so three shapes built from one spread came out named
  `Shape 0`, `Shape 0`, `Shape 0` and warned about the collision. `addTable` wrote nine
  keys back — `objectName`, `fontSize`, `margin`, `color` and the five `autoPage*` ones —
  and handed the caller's object to every plain string cell as that cell's options, so
  the cell emitters wrote onto it too. `autoPage` is the one with teeth: the auto-pager
  clears it once it has shredded the rows, so a literal reused for a second table
  silently lost its paging. Both now copy on the way in, along with `addTable`'s
  `border` array, whose slots were normalized in place. Emitted bytes are unchanged
  (183/183 baseline parts) — the options each definer owns still share identity within a
  single call, which is what the table's string cells and the text shorthand rely on.
  Nested caller objects follow `addText`: `fill` stays shared by reference for its rel
  id, and so does `shadow`, whose in-place normalization is idempotent and is what lets
  one shadow object give two shapes the same `<a:effectLst>`.

- **`Shape.shadow` and `Shape.innerShadow` read the shadow colour whatever model it
  uses.** The decode named `a:srgbClr` and `a:schemeClr` explicitly and dropped the other
  four DrawingML colour models on the floor, so a shadow coloured with `a:sysClr`,
  `a:prstClr`, `a:hslClr` or `a:scrgbClr` read as `color: null`. The colour element is the
  shadow's only child and it is required (`a:EG_ColorChoice`, a single-member group), so
  it now takes the first element child — exactly what `Shape.glow` alongside it has always
  done. `a:sysClr` was the live case: it resolves everywhere else in the read model,
  through its `lastClr` snapshot. `a:prstClr` reaches the resolver, which now carries the
  preset table — see the colour-model entry above — so a preset-coloured shadow reports a
  colour rather than being silently swallowed by the accessor.

- **A chart point cache can no longer size an allocation straight from the file.** Both
  chart readers turned a cached series into a dense array by trusting two numbers the deck
  supplies: `c:ptCount/@val` (`cx:lvl/@ptCount` on the 2016 side) and each point's `@idx`.
  `@ptCount` is `xsd:unsignedInt`, so `4294967295` is schema-valid, and
  `new Array(4294967295).fill(null)` is not a slow path — V8 answers `FATAL ERROR: invalid
  table size` and the host process dies with no exception to catch. It was reachable
  through the public API: `Presentation.load()` on such a deck returned normally and the
  first `chart.series` access killed the process. A single `<c:pt idx="900000000"/>` did
  the same on its own. Both readers now share one bounded decode: the array is sized by the
  points that are really present, and a point indexed past a worksheet's 1,048,576 rows —
  more than any workbook-backed cache could reference — is dropped. Two new diagnostics,
  `chart/point-count-mismatch` and `chart/point-index-out-of-range`, report either case
  rather than trimming silently. One consequence on legitimately sparse data: a cache that
  declares four points but carries three (the last cell blank) now reads three long instead
  of four. No value is lost — the dropped slots are the ones that read `null` — and the
  disagreement is warned about.

- **A table cell's `colspan`/`rowspan` can no longer size an allocation straight from the
  caller.** The merge-grid builder trusted them: `new Array(colspan - 1).fill(undefined)`
  at `colspan: 4294967295` is not a slow path but a process abort with no exception to
  catch, and a negative or fractional span shifted every column after it while emitting a
  `gridSpan` PowerPoint cannot make sense of. This is the write-side twin of the chart
  point-cache fix above, and less serious for the same reason — the caller is the program,
  not a hostile file — but the project's line is to warn rather than emit a degenerate
  result. A span that is not a whole number in 1..1000 now falls back to `1` and reports
  the new `table/span-out-of-range` diagnostic. The ceiling is a sanity bound, not a schema
  one (`a:tc@gridSpan` is a bare `xsd:int`); it sits an order of magnitude past
  PowerPoint's own maximum table, whose Insert Table dialog stops at 75 × 75. There were two
  allocations to guard, not one: the auto-pager sizes a per-column depth array from a column
  count that is a sum of colspans, so `autoPage: true` hit the same abort without reaching
  the merge grid at all. Both are fed by `addTableDefinition`, which is where the check runs
  — once, before anything reads a span, so the paged and unpaged paths agree on the grid and
  one bad cell warns once rather than several times. Emitted bytes are unchanged (183/183
  baseline parts): a cell whose spans are fine passes through untouched.

### Changed

- **The bundle-size budget watches every published entry point, not just the browser
  one.** `bundle-size:check` always measured an entry's whole transitive closure, shared
  chunks included — that part was never the gap. The gap was that only `browser.js` had a
  number, so a dependency landing in a chunk `ts-pptx/read` pulls in cost every consumer of
  that subpath and tripped nothing until it also reached the browser entry. All ten entries
  `package.json` publishes are budgeted now, which between them reach every `.js` file the
  build emits. Keying on entries rather than on chunks is what keeps it stable: an entry
  file name carries no content hash. Switching the other nine on immediately found two
  defects in the ratchet itself — a `{@link import('./x.js')}` in a doc comment counted as
  an import and made the closure demand a file no build emits, and the "you are far enough
  under to re-freeze" nag could not be satisfied on an entry small enough that `--freeze`'s
  rounding to a whole kB exceeded the slack threshold.

- **`Reflection.distPt` is now `Reflection.offsetPt`** (`ts-pptx/read`). Sibling accessors
  on one class spelled the same `@dist` attribute two ways: `Shape.shadow` and
  `Shape.innerShadow` reported `offsetPt`, `Shape.reflection` reported `distPt`. The three
  now agree. Migration is a rename at the call site — `reflection.distPt` →
  `reflection.offsetPt`; the value, the unit (points) and the omit-when-absent behaviour
  are unchanged. `offsetPt` won over `distPt` because the read API is meant to read well,
  not to mirror OOXML attribute spelling.

- **`NotesPlaceholder` now extends `Placeholder`** (`ts-pptx/read`). It was a copy of it:
  the same twelve identity, geometry and escape-hatch members, read off the same `p:sp`,
  down to `p:ph` being located the same way — plus a flattened `text` convenience and a
  `textFrame` that threads the notesMaster inheritance context onto the body frame. Those
  two are now the only things it declares. Nothing about any member's behaviour changed;
  what changed is that `notesPlaceholder instanceof Placeholder` is `true` where it was
  `false`. That is a widening, and the answer a caller writing one type guard over "any
  placeholder in the deck" would want, but it is worth knowing if you were relying on the
  `instanceof` to tell a notes placeholder apart from a layout one — use
  `instanceof NotesPlaceholder` for that, which still reports only notes placeholders.

## [3.5.0] - 2026-08-26

This release makes a text edit land where the text is actually read. `Paragraph.text` and
`DiagramPoint.text` are settable, and the SmartArt setter writes the fallback drawing cache
that every renderer without a layout engine paints from, not just the data model PowerPoint
reads. The diagram data model gains its tree (`Diagram.nodes`, `Diagram.point()`) and a link
from a node to the shape drawn for it. `importSlide` stops copying chrome the destination
already holds, which takes a duplicate layout and master off every import between decks
templated from one file, and with it off the carried slides `ts-pptx/script` emits. And two
read surfaces stop losing graphic frames: `inspectPptx` reports them instead of skipping
them, and a SmartArt slide is copied rather than transcribed into a script with a hole in it.

### Added

- **`Paragraph.text` is settable.** `Run.text`, `TextFrame.text`, `Shape.text` and
  `TableCell.text` could all be assigned; the paragraph in between could not, so replacing
  one bullet of a list meant either collapsing the whole frame or rewriting runs by hand.
  Setting it replaces that paragraph with a single run, keeping the first run's `a:rPr` and
  the paragraph's own `a:pPr` (level, alignment, bullet), and leaves sibling paragraphs
  untouched — the same rule the SmartArt drawing-cache mirror already ran on internally.

- **`DiagramPoint.text` has a setter, and it keeps the drawing cache honest.** A SmartArt
  diagram stores every string twice: as authored nodes in `ppt/diagrams/data{N}.xml`, which
  is what PowerPoint reads, and as a copy of every drawn string in
  `ppt/diagrams/drawing{N}.xml`, which is what every renderer without a SmartArt layout
  engine paints. Editing a node through `textFrame` wrote only the first, so the change was
  invisible in LibreOffice, Google Slides, thumbnailers and web previews until someone opened
  the deck in PowerPoint and saved it. The new setter writes both, replacing the one paragraph
  of the drawn shape that belongs to that node and leaving its siblings byte-identical (one
  shape commonly draws three nodes at once). Geometry is **not** recomputed, which is inherent
  rather than a defect: the drawn shape keeps its cached size, so a much longer string
  overflows its box in those renderers until PowerPoint re-lays the diagram out.

  `textFrame` still edits the data model alone and now says so; it is the escape hatch for
  run-level formatting. See `docs/reference/pptx-read.md`.

- **`Diagram.nodes` gives the data model its tree, and `Diagram.point()` resolves an id.**
  `connections` hands back the raw `parOf` edges, so every consumer that wanted "the top-level
  nodes, in order" wrote the same index-and-walk, and a hierarchy read as a flat list with no
  indication of depth. A `DiagramNode` carries its `point`, its `children` (in `@srcOrd`
  order), its `parent` and its `level`. Assistant (`asst`) points are in the tree and keep
  their own type; transition points are not, because they label an edge rather than a node and
  are reached through `DiagramConnection.parentTransitionId` / `siblingTransitionId`. A
  `parOf` cycle raises `diagram/parent-edge-cycle` rather than hanging or silently dropping
  the points it swallows.

- **`DiagramPoint.drawnShape` and `DiagramPoint.presentationId` link a node to what is drawn
  for it.** `drawnShape` resolves the `dsp:sp` of the fallback drawing part that carries the
  point's text, plus the paragraph index inside it, as a `TextFrame` bound to the *drawing*
  part. The mapping is not the obvious one and is many-to-one in both directions: a
  `dsp:sp/@modelId` is always a generated `pres` point's id rather than the authored point's,
  one shape draws several points at paragraph indices that can contradict document order, and
  one point has several presentations of which at most one holds text. `null` is a defined
  outcome, covering a package with no drawing part, a point with no `presOf` edge (every
  unlabelled transition point), a `pres` point that draws nothing, and a drawn shape with no
  text body. Measured over five layout families and 90 authored points.

### Changed

- **`importSlide` binds to chrome the destination already holds instead of copying it in
  again.** `copyPart` walked the source slide's `slideLayout -> slideMaster -> theme`
  subgraph and copied every part of it under a fresh partname, with no notion of what the
  destination arrived with. For the deck templated from its own source (`fromTemplate` keeps
  a package's chrome byte-identical and strips only its slides, which is exactly what
  `ts-pptx/script`'s template-anchored tier emits) every one of those parts was already there
  under its own partname, so each imported slide grew the deck a duplicate layout and master:
  one extra entry in PowerPoint's layout picker per slide, and a later
  `appendSlides({ layout: <name> })` that threw `layout/ambiguous-name` because two layouts
  answered to the name.

  A part is now reused when the destination holds it at the same partname with the same
  content type, the same bytes as they would be written, the same relationships, and the same
  test passing recursively; a master must also be registered in `p:sldMasterIdLst` and a
  layout must sit in a registered master's gallery. The decision is made at the page boundary
  and never below it, so a part reached from something this import copied is copied too and
  copied chrome stays self-contained. Anything short of identical still copies.

  **Migration.** Importing between two decks that share byte-identical chrome (the same file
  opened twice, or two decks made from one template) now adds fewer parts: no new layout,
  master or theme, and the imported page resolves to the destination's own. Rendering is
  unchanged, since reuse only happens where the bytes were the same. A consumer asserting on
  part counts after such an import sees the smaller numbers.

- **A converted script's carried slide no longer costs a layout, and batches bind by name
  again.** `ts-pptx/script`'s template-anchored tier copies a slide the write API cannot
  author (`slide.carried`) with `importSlide`, out of the very file it templated the deck
  from. That import used to duplicate the slide's layout and master, and because the duplicate
  repeated a layout *name*, it demoted every `appendSlides` binding in the emitted script from
  the name to a gallery position. With `importSlide` reusing chrome the destination already
  holds, layout names stay unambiguous and bindings print as `{ layout: 'Titelfolie' }`. A
  multi-master deck that genuinely repeats a layout name still falls back to a position, which
  is the case the fallback was always for.

- **`inspectPptx` reports graphic frames.** A `p:graphicFrame` (table, chart, SmartArt, or a
  payload this library does not model) was skipped outright and consumed no `zIndex`, so a
  deck whose slides are SmartArt or tables inspected as `elements: []` with `wordCount: 0`,
  indistinguishable from a deck of blank slides. It also put the two read surfaces in flat
  disagreement: `Slide.text` on the deep model has always flattened table cells and SmartArt
  node text.

  A frame now comes back as one element with `kind: 'graphicFrame'`, its own box, a `zIndex`
  in the shape-tree walk, and a new `graphicKind` field (`'table' | 'chart' | 'chartEx' |
  'diagram' | 'other'`). Its `text` is what a reader sees on the slide, so it counts toward
  the slide's `text` and `wordCount`; a chart contributes none, matching `Slide.text`, which
  treats data labels as chart data rather than slide body text. The *structure* is still not
  flattened: `textRuns` and `paragraphs` stay empty, and cells, series, and nodes are reached
  through `ts-pptx/read`.

  **Migration.** `PptxSlideElementKind` gains `'graphicFrame'`, so an exhaustive `switch` over
  it must handle the new member; `PptxSlideElement` gains `graphicKind`, `null` on every other
  kind. A consumer that counted elements, summed `wordCount`, or read `zIndex` values sees
  different numbers on decks containing graphic frames. Filter on `kind !== 'graphicFrame'` to
  get the old element set back (the old `zIndex` numbering is not recoverable, and was itself
  a walk with holes in it).

### Fixed

- **A slide holding SmartArt is no longer emitted as a script with a hole in it.**
  `ts-pptx/script` decides per slide whether the emitted script can describe it or whether the
  printer must copy the source slide verbatim, and that test named extended charts and nothing
  else. But three graphic-frame payloads produce no call, not one: an extended chart, a
  SmartArt diagram, and a frame the reader does not decode at all (an OLE object, ink, a 3-D
  model). So a chartEx slide round-tripped correctly by being copied while a SmartArt slide
  was transcribed and silently lost its diagram, with only a per-shape fidelity note recording
  the loss. Both sites now share one predicate, so they cannot drift apart again.

  The copy costs nothing in the layout gallery, because `importSlide` now binds to chrome the
  destination already holds (see above) and the template-anchored tier copies out of the very
  file it templated the deck from. Layout handles are still resolved up front rather than
  where first used, since an `importSlide` between two `appendSlides` calls moves the gallery
  underneath them.

  Migration: a deck whose SmartArt, OLE or ink slides previously came out `authored` now comes
  out `carried`. The emitted script imports those slides from the template instead of
  rebuilding them, which needs the source deck present (it already did for the template
  itself). The standalone tier is unchanged: it has no source package to copy from, so it
  still transcribes such a slide and still reports the per-shape note.

## [3.4.0] - 2026-08-26

This release makes SmartArt readable through the read model, and names what a frame the
reader still cannot decode actually holds. It refuses two option values that used to reach
the emitted XML as nonsense — a non-finite size or angle, a stringly-typed
`legendFontSize` — and gives three others the meaning their absence already carried: an
empty `chartColors`, a `null` side in a border tuple, and a table that named a built-in
style only to have the writer's own grid painted over it. It also takes `Math.random()` out of
the chart palette, so a chart with more series than colours emits the same bytes on every
build.

### Added

- **SmartArt is readable: `GraphicFrame.hasDiagram` and `GraphicFrame.diagram`.** A
  `p:graphicFrame` whose `a:graphicData/@uri` is the DrawingML diagram namespace answered
  `false` to all three host predicates and `null` to all three accessors, so its content was
  unreachable through the read model, and a consumer could not distinguish "this frame holds
  text I cannot get at" from "this frame holds a chart with no labels". `Diagram` decodes the
  `dgm:dataModel` part: `points` (every `dgm:pt` in document order, typed
  `node`/`asst`/`doc`/`pres`/`parTrans`/`sibTrans`), `connections` (the `dgm:cxn` edges that
  give the points their tree), and `text` (the authored node text, generated `pres` points,
  the `doc` root and unfilled placeholders excluded). A point's `textFrame` is an ordinary
  `TextFrame` over its `dgm:t`, so runs, formatting and the `resolved*` getters all apply.
  The layout, quick-style, colours and fallback drawing parts resolve as `Part`s, and the
  `doc` point's `layoutTypeId` names the SmartArt kind. Authoring SmartArt remains out of
  scope. See `docs/reference/pptx-read.md`.

- **`GraphicFrame.graphicDataUri` reports what an undecoded frame holds.** Every `has*`
  predicate is a comparison against this URI, so a frame that matches none of them is one the
  read model does not decode (an OLE object, an ink annotation). Reading it lets a consumer
  say which construct it could not reach instead of inferring loss from four `false`s. The
  class comment had described this accessor since the frame reader was written; it had never
  actually been public.

- **`resetDiagnosticState()` clears the once-per-process warning record.** Conditions that
  would flood a log are emitted once per distinct code and message, and that record was
  process-global with no way to clear it. So a service building a second deck got no warning
  for anything the first deck already tripped, which reads as "no problems found" rather than
  as "already mentioned"; and since most of those messages interpolate the offending value,
  the record grew by one entry per distinct bad value with nothing to release them. Call it
  between builds. It does not touch the handler — `setDiagnosticHandler` owns that, and the
  two reset independently because a host usually installs its handler once. See
  `docs/diagnostics.md`.

### Changed

- **`Slide.text` now includes SmartArt node text.** A slide whose whole message is a diagram
  flattened to the empty string, or to its title alone. Diagram nodes are body text that
  PowerPoint itself searches and spell-checks, so they contribute a block each, in shape-tree
  order, exactly as table cells already did. Chart data labels and speaker notes are still
  excluded, unchanged.

  Migration: if you were summing `Slide.text` lengths to detect empty slides, or diffing that
  string across versions, a deck containing SmartArt will now report more text. Read
  `GraphicFrame.diagram` directly if you need the diagram separated back out.

- **A converted script now distinguishes `diagram.all` from `graphicFrame.unknown`.** SmartArt
  used to be reported under the latter, whose stated cause is `unread`. With a reader in front
  of it, the loss is the write leg's: `pptxToScript` emits `diagram.all` (`unwritable`,
  `dropped`), the same shape as `chartEx.all`. `graphicFrame.unknown` keeps its meaning and now
  names only genuinely undecoded frames.

- **A non-finite size or angle is now refused rather than serialized.** `fontSize`,
  `charSpacing`, `lineSpacing`, `paraSpaceBefore`/`After`, `rotate`, a chart's
  `catAxisLabelRotate`/`valAxisLabelRotate`/`titleRotate`, and a `custGeom` connection-site or
  adjust-handle angle all threw `Infinity` straight through their converter and into the
  attribute — `sz="Infinity"`, `rot="Infinity"` — which is not a legal value and made
  PowerPoint offer to repair the package. Each now throws an `InvalidOptionError` with code
  `coord/non-finite` at the point of conversion. A *finite* out-of-range value is unaffected:
  it still clamps into the schema range and warns, as before.

  Migration: if you were computing one of these from data that can produce `Infinity` or
  `NaN` (a division by a zero row count is the usual one), guard it at the call site. A `NaN`
  is unchanged on the size options — every caller reads them for truthiness, so a `NaN` has
  always meant "absent" there.

- **A rotation past a full turn now reduces correctly.** `convertRotationDegrees` subtracted a
  single turn, so `rotate: 800` became 440 degrees rather than 80, and `rotate: 370` became 10
  while `rotate: -400` was not reduced at all. It is now `% 360`, sign-preserving: a rotation
  already within a turn — including a negative one — is byte-for-byte unchanged, since both
  `-45` and `315` are valid `ST_Angle` and the read side reports back what was authored.

  The same function was also being used for angles that are *not* modular: a `custGeom`
  connection site, a polar adjust handle's `minAng`/`maxAng`, an `angleRange` guide, and the
  chart label rotations. Those now go through a non-wrapping converter, so an adjust handle
  declared with `maxAng: 540` keeps its full travel instead of silently collapsing to 180.

- **An empty `chartColors` now means the same as omitting it.** `chartColors: []` used to give
  every chart the *bar* palette, including a pie or doughnut, whose default when the option is
  omitted is `PIECHART_COLORS`. Nothing chose that: normalization defaulted a *missing*
  `chartColors` per chart type, `Array.isArray([])` is true, so an explicit empty array walked
  through that pass untouched and met a later fallback that did not know the chart type. An
  empty array names no colours, which is what omitting the option means, so both now resolve to
  the built-in palette for the chart's own type. A combo subchart's palette resolves the same
  way for the first time — its options never went through that normalization pass at all.

  Migration: if you passed `chartColors: []` to a pie or doughnut and wanted the bar palette,
  pass it explicitly. Every other chart type is unaffected: its default already was that
  palette.

- **A chart's `legendFontSize` is no longer coerced from a string.** It was the one font-size
  option whose emitter wrapped the value in `Number()`, so `legendFontSize: '12'` from untyped
  JS worked there while the same string threw at `catAxisLabelFontSize`, `dataLabelFontSize`
  and the eight other spellings of the same option. The coercion is gone; a non-number now
  throws `InvalidOptionError` with code `coord/non-finite` wherever it is passed.

  Migration: pass a number, which is what the type has always said.

- **A `null` side in a cell's border tuple is now omitted rather than erased.** A four-side
  TRBL tuple replaced every falsy side with `{ type: 'none' }`, so `border: [solid, null,
  solid, null]` emitted four `<a:lnX>` elements — two drawn, two carrying an explicit
  `<a:noFill/>`. On `<a:tcPr>` those are not the same state: an absent edge inherits from the
  table style, theme banding and the master chain, while a present `<a:noFill/>` overrides
  that inheritance. So a hole in the tuple meant the opposite of the omission it reads as, and
  there was no spelling left for "draw the edges I name, leave the rest to the style".
  A `null` side now leaves its element absent, which is the distinction `normalizeOuterBorder`
  already documents for the perimeter ("a sparse side is *not* `{ type: 'none' }`"), and the
  tuple type accepts `null` so the sparse form is authorable from TypeScript. An explicit
  `{ type: 'none' }` still emits `<a:noFill/>`, so both states stay reachable.

  Migration: a tuple that relied on `null` to erase an edge should say `{ type: 'none' }`.

- **A table with a `tableStyle` keeps that style's grid.** Every cell with no border authored
  used to receive four explicit `<a:lnX><a:noFill/></a:lnX>` edges. That is direct formatting,
  so it beat the built-in style the caller had just selected, and a `MEDIUM_STYLE_2_ACCENT_1`
  table came out with the style's fills and banding but none of its rules — with no way to ask
  for them, since the only spelling for "leave this edge alone" was the one being overwritten.
  The default is now applied only to a table that named no style, which is where it earns its
  keep: PowerPoint's no-style look is a black hairline grid, and suppressing that is what the
  force-fill was for. A styled table emits no `<a:tcPr>` edges unless you author some, and
  `border: { type: 'none' }` still erases the grid explicitly. `outerBorder` composes as it
  always did: on a styled table it now draws the perimeter and leaves the interior rules to
  the style, rather than blanking them.

  Migration: a styled table that relied on the force-fill for a gridless look should author
  `border: { type: 'none' }`. Unstyled tables are byte-identical.

### Fixed

- **Auto-paged tables no longer lose a point of column width to floating-point rounding.**
  The chars-per-line figure `autoPage` wraps on comes from the column's width in points, and
  the conversion went through EMU: `(colWidth / EMU_PER_POINT) * EMU_PER_INCH` is the same 72
  as `colWidth * POINTS_PER_INCH`, but not exactly, and the `Math.floor` around it took the
  hit. Of the 200,000 widths from 0.001in to 200in, 51 came out one point narrow — 6.625in is
  exactly 477 points and was measured as 476.

  The effect is not cosmetic on those widths: a line that fits is wrapped to a second one, so
  every row of the table is measured as twice as tall as it is and the table pages onto more
  slides than it needs. At 6.625in and 15pt a sixteen-row table now fits on one slide where it
  used to split across two. Only those 51 widths change; every other one is byte-identical, and
  no showcase deck reaches one.

- **A slide-number field on a master or layout no longer caches a fake page number.**
  `<a:t>` inside an `a:fld` is the field's *cached* rendering, and a master has no slide number
  while a layout carries the library's internal 1000-and-up counter. So every master shipped
  `<a:t>null</a:t>` and every layout something like `<a:t>1004</a:t>`. PowerPoint recomputes
  the field on open, so nothing ever looked wrong there — but anything that reads the cache
  without recomputing (a text extractor, a search indexer, this library's own read path) saw
  it. Both now cache the placeholder glyph `‹#›`, which is verbatim what PowerPoint itself
  writes: every en-US master and layout in the read fixture corpus has exactly that. A real
  slide still caches its own number.

  This changes emitted bytes for `slideMaster*.xml` and any layout carrying a slide-number
  placeholder, and for nothing else.

- **Chart colours past the end of the palette now cycle instead of being drawn at random.**
  A chart with more series or data points than `chartColors` has entries used to pick the
  overflow colours with `Math.random()`, so the same deck built twice emitted different
  bytes and a 17-slice pie was recoloured on every build. All palette lookups now wrap back
  to the start of the palette, which is what the majority of them already did.

  Visible if you author more series or points than your palette covers: those now take the
  palette's own colours in order rather than an arbitrary one. The two built-in palettes
  each held their entries twice over, which only postponed the moment the broken wraparound
  was reached; they now hold each colour once, and a chart past their end repeats them in
  order. Output is unchanged for any chart that stayed within the palette.

- **A zero-valued geometry shortcut now emits its adjustment guide.** `genXmlPresetGeom` gated
  `rectRadius`, `angleRange` and `arcThicknessRatio` on truthiness, so a deliberate zero was
  read as an unset option. `addShape('roundRect', { rectRadius: 0 })` emitted no `<a:gd>` at
  all and PowerPoint fell back to the preset's own ~16.7% rounding, which is the one radius a
  caller asking for zero does not want; `arcThicknessRatio: 0` lost `adj3` while
  `angleRange: [0, 0]` kept `adj1`/`adj2`, an array being truthy whatever it holds. Each gate
  now asks whether the option was supplied. Radius 0 is a sharp corner, `[0, 0]` a closed arc
  and thickness 0 a zero-thickness band, and each builds the guide at `val 0`; every non-zero
  input is byte-identical.

- **`TextPropsOptions` declares the geometry keys the text-frame emitter already reads.**
  A styled text frame carries preset geometry through the same `genXmlPresetGeom` a shape
  does, and it honoured `angleRange`, `arcThicknessRatio`, `points` and `shapeAdjust` at
  runtime while the declarations named only `shape` and `rectRadius`. So `addText` with an
  arc's angles built exactly the right `<a:gd>` guides and TypeScript rejected the identical
  object literal with TS2353 — and any consumer whose option map is derived exhaustively from
  the published types could not compile against the writer's real surface. The four keys now
  reuse the vocabulary `ShapeProps` already carries. Declaration-only: no emitted bytes move.

## [3.3.0] - 2026-08-25

This release includes a breaking cleanup to the in-memory byte export API, speaker notes on
both read-model paths that lacked them (authoring onto a loaded deck, and the batch import),
a chart type that is now refused rather than emitted as a chart with nothing in it, the
Office-2016 chart family carrying across that same append bridge, plus repository and
project-site changes.

### Added

- **`isChartType` narrows a string to the chart catalog.** The companion to `isChartExType`,
  which reported only which *half* of the catalog a type belongs to and answered `false` for
  a string that is in neither. Useful for validating a chart type that arrives as data — from
  a config file, a spreadsheet column, a CLI flag — before it reaches `addChart`, which now
  throws on one it does not recognise (see *Changed*).

- **`importSlides` carries speaker notes, per request.** `ImportSlidesRequest` gains
  `importNotes?: boolean`, the batch spelling of the option `importSlide` has had all
  along; the default is unchanged, so a request that does not ask still arrives without
  notes. Closes the half of #17 the previous change left open.

  Per request rather than per batch because a stitch mixes sources: the notes of a
  library's cover page are worth carrying where a scratch deck's are not. What the flag
  does *not* decide is the styling — a presentation holds at most one `notesMaster`, so the
  destination's own is reused when it has one and the first carried master is installed
  when it has none, exactly as `importSlide({ importNotes: true })`, `appendSlides` and
  `Slide.addNotes` do. The four paths cannot between them produce a second notes master.

  Two things had to move for the batch's own guarantees to survive the addition:

  - **The up-front dry run now walks the notes graph.** Notes are copied after the pages
    are, so a batch that could not finish carrying them was the one remaining way back into
    a half-stitched deck. `checkSelectionCopyable` takes the opted-in pages and walks each
    one's notes subgraph under exactly the rules `carryNotes` will follow — including
    skipping the source notes master when the destination has one of its own, since that
    master is then never read and a check that walked it anyway would reject a batch the
    copy would have completed. A rejected batch still leaves the deck byte-identical.
  - **A notes slide is now copied as a part its page owns.** `carryNotes` opens an
    ownership scope (`page-owned.ts`), so a page named twice — reachable in one call for
    the first time here, and already reachable through two `importSlide` calls — gets notes
    of its own each time, and so do the parts those notes own. Sharing one of those between
    two pages is a package PowerPoint refuses to open (`0x80070570`); media stays shared,
    as everywhere else.

- **`Slide.addNotes(text)` on the read model** — the write counterpart to the
  `notesText` / `notesTextFrame` / `notesSlide` getters, and the only way to annotate a
  slide with **no notes part at all**. That state is reachable in normal use — it is what
  `importSlide` without `{ importNotes: true }` leaves behind — and it was a dead end: the
  notes accessors are getters, and the settable `notesTextFrame` is `null` precisely when
  there is no part, so a consumer that imported a page could not give it notes of its own
  through any API. Reported as #17.

  A `
` starts a new paragraph, matching the write-side `addNotes`; runs carry no
  formatting of their own, so per-run colour/size/hyperlinks are set afterwards through
  `notesTextFrame`. Called on a slide that already has notes it replaces the body text and
  leaves the rest of the part — geometry, the `sldImg`/`sldNum` placeholders — alone.

  Creating the part means creating what a notes slide must bind to, and a presentation may
  hold at most one notes master. The deck's own is reused when it has one; otherwise one is
  installed, bound to a **clone of the deck's own theme** rather than to the slide master's
  theme part, so no two masters claim one part and the notes resolve against the
  destination palette. That is the same single-master rule `importSlide({ importNotes:
  true })` and `appendSlides` already follow, so mixing the three cannot produce a second
  notes master.

  The part it builds is the generator's, not a second design: `makeXmlNotesSlide` was split
  so its fixed three-placeholder frame (`makeXmlNotesSlideSkeleton`) is shared with the read
  side instead of copied into it, and `test/read/add-notes.test.js` asserts the two paths
  emit the same notes body — modulo empty-element spelling, which the read model's
  serializer normalizes. This is the first `src/read/` → `src/gen/` import; it buys the
  anti-drift guarantee that `src/ooxml/` exists to give the constants both halves share.

### Fixed

- **A chartEx chart appended onto a loaded deck arrived empty.** `Presentation.extractSlides()`
  — the bridge `ts-pptx/read`'s `appendSlides` serializes generated slides through — built every
  chart with the classic `<c:chartSpace>` builder, which has no arm for the Office-2016 family
  (`waterfall`, `funnel`, `treemap`, `sunburst`, `histogram`, `pareto`, `boxWhisker`,
  `regionMap`). A waterfall came out as a 2.6 kB chart part with axes and no plot element at
  all, registered under the classic content type, behind a slide still pointing at it through
  `<mc:AlternateContent><cx:chart>` — a chart-shaped hole PowerPoint opens and shows empty.

  Both halves of the bridge now know the second shape. `extractSlides` builds a chartEx chart
  with `makeXmlChartEx` and hands its two mandatory style sidecars along in the descriptor's new
  `chartEx` slot; `appendSlides` injects it as its own `chartEx{N}.xml` part under the Microsoft
  chartex content type, wires the slide to it through the MS `chartEx` relationship rather than
  the ECMA `chart` one, and writes `style{N}.xml` / `colors{N}.xml` at the chart part's rId3 /
  rId2. The sidecars are not decoration: PowerPoint reports a chartEx part without them as
  corrupt (`0x80070570`), which the schema validator does not see. The generator builds them, so
  the read subsystem still imports nothing from the emitters at runtime.

  `write()` was never affected: it routes on `isChartExType` and always has. Classic charts
  through `appendSlides` are unchanged, byte for byte. The interim
  `UnsupportedFeatureError` (`chart/chartex-not-extractable`) that refused these charts is gone
  from `UnsupportedFeatureErrorCode` — it was added and removed inside this same unreleased
  cycle, so no released version can throw it.

- **The read reference documented three of `ImportSlideOptions`' seven fields.**
  `importNotes`, `rescale`, `remapLiterals`, and `embedFonts` were absent from
  `docs/reference/pptx-read.md` — `importNotes` appeared nowhere in `docs/` at all — so the
  documented way to carry a slide's speaker notes across an import was invisible to anyone
  reading the docs rather than the `.d.ts`, and #17 was filed as a missing feature on that
  basis. The interface block now matches the interface, `importNotes` has a prose section
  covering the notes-master policy, and the neighbouring claim that "deliberate re-branding
  (a `restyle` mode) is not yet implemented" — untrue since `restyle` shipped, and
  contradicted by the section above it — is corrected. No behaviour change: the option has
  worked in all three `theme` modes since it shipped. `ImportSlidesRequest` was likewise
  absent from the interface block entirely, and the batch section is now explicit about
  what notes do there rather than saying only that they are dropped.

### Changed

- **A chart type outside the catalog is refused at `addChart` instead of emitting a chart with
  no plot.** `addChart(data, { type })` took the type on trust: the `CHART_NAME` union
  constrains TypeScript callers, but nothing enforced it for JavaScript ones, and the two chart
  emitters partition `ChartType` between them and each treats the other's members as not its
  own — so an off-catalog string matched no arm anywhere and fell through to an empty-string
  default. `addChart(data, { type: 'nonsence' })` was *accepted*, and the typo arrived in the
  deck as a chart frame with axes and nothing inside it. Both the single-type form and every
  `ChartMulti` entry of a combo are now checked against the catalog, throwing
  `InvalidOptionError` (`chart/unknown-type`) naming the valid types. The two emitters no longer
  have a silent default at all: reaching one with a type the other owns throws `InternalError`
  (`chart/type-not-routed`), which no public input can now produce. Every catalog type builds
  exactly the bytes it did before.

- **`toBytes()` replaces `stream()`.** The old method did not stream: it assembled the
  complete archive, then copied fflate's `Uint8Array` into a Node `Buffer`. It therefore
  failed in browsers without a `Buffer` shim, returned an unnecessarily broad type, and
  ignored `onMediaError`. `toBytes({ compression, onMediaError })` now returns exactly
  `Promise<Uint8Array>` in Node, browsers, and workers without the extra archive-sized
  copy. Migrate `await pptx.stream()` to `await pptx.toBytes()`; callers that need another
  representation should continue using `write({ outputType })`.

- **`pnpm run verify` no longer sizes itself off the host's CPU, and no longer builds the
  site.** The gate could take a developer machine down: Vitest sizes its fork pool from
  `availableParallelism() - 1`, and every concurrent fixture spawned its own 55 MB
  `OOXMLValidatorCLI` process, so the real ceiling was `workers × maxConcurrency` — a
  property of the developer's core count rather than of this repository, and the wrong
  one. A faster CPU made the spike bigger, never the run safer. Three changes, all in
  test/build tooling; nothing the published package does is affected.

  - `test/validator.js` **batches** validation. `OOXMLValidatorCLI` accepts a directory
    and validates every package in it in one process, and measured against the pinned
    release a run costs ~0.40s of startup plus ~0.048s per additional deck at ~55 MB
    regardless — so one deck per process was paying that startup, and that 55 MB, some
    500 times per suite. Each worker now holds at most one validator child. Full suite:
    55.1s / 4.03 GB / 17 concurrent validators → 48.4s / 3.40 GB / 7.
  - `vitest.config.ts` derives `maxWorkers` from **memory free at startup** rather than
    core count, since peak RSS is linear in the pool size. Idle machines still land on
    the CPU bound; loaded ones scale down instead of swapping. `VITEST_MAX_WORKERS` pins
    it explicitly.
  - `docs:build` moved from `verify` to `verify:full`. 19.7s of its 26.6s is
    `vitepress build`, which proves something about the site, not the library; `docs:check`
    keeps the docs validated in the loop. CI is unaffected — `docs.yml` already built the
    site on every pull request, so `check:static` was building it a second time.

  `verify` is ~97s → ~44s and `verify:full` ~77s, both on a 12-core box. Composites are
  now assembled by `scripts/run-steps.mjs`, which expands script names into leaf commands
  and runs them in one process tree, removing the ~0.7–1.3s package-manager relaunch that
  `verify` was paying 13 times over. `package.json` remains the single definition of each
  step. See `docs/testing.md`.

  One latent test bug surfaced and is fixed: seven schema fixtures assert on warnings by
  swapping `console.warn` across an `await`, which is unsound under `describe.concurrent`
  — a neighbour's `finally` restores the original mid-capture. It was a race, so it was
  always latent; changing how fixtures interleave is what exposed it. Those fixtures now
  run in a sequential sibling suite, and new ones are routed there automatically.

- **Schema validation moved to the shared `ooxml-validate` oracle.** This repo and
  `ts-xlsx` each carried their own validator pinned to a different Open XML SDK version —
  3.2.0 here, 3.5.1 there — so the two enforced different rule sets while looking like
  they enforced the same one. Both now validate through one published package, and the
  SDK pin lives in that package rather than in either consumer. Test tooling only;
  nothing the published package does is affected.

  - `tools/ooxml-validator/` is gone — installer, update checker, version pin and the
    `check-validator` script with it. `ooxml-validate` fetches its oracle from GitHub
    Releases on first use (never from a postinstall hook), verifies the download's
    checksum *and* its build provenance, and caches it under `~/.cache/ooxml-validate`.
    CI caches that directory instead of running an install step.
  - `test/validator.js` shrank to an adapter. Batching, the one-child-at-a-time queue and
    the CI gate belong to the package now; the suite's `validateBuf` /
    `validatorAvailable` surface is unchanged, so no fixture moved.
  - **Every input file now comes back with an explicit `valid` flag.** The old batcher
    read "absent from the output" as "clean", which was safe only for as long as
    unreadable packages kept being reported. Nothing infers cleanliness from absence any
    more.
  - Diagnostics are `{ id, type, description, partUri, xpath }` — the same five fields,
    renamed with the report contract. A non-package file is now `PackageOpenError` /
    `Package` rather than a null id and an `OpenXmlPackageException` type; the two
    self-check fixtures in `test/schema-cases.js` pin the new shape.
  - Validating at SDK 3.5.1 changed nothing about what this project emits. All 153
    generated decks plus the 44 static fixtures were run through both oracles before the
    switch: identical diagnostics, file for file, nothing new and nothing gone. The
    conformance target is still `Microsoft365`, and `pnpm run schema:versions` still
    prints the coverage table it printed at 3.2.0.
  - Set `OOXML_VALIDATE_NO_BATCH=1` (was `TSPPTX_VALIDATOR_NO_BATCH`) to pin a batch
    failure to one fixture.

- **The browser demo is a page of the site, not a workspace.** `demos/vite-demo` — a React
  + Vite + Bootstrap app whose only output was a download — is deleted. Its replacement is
  `/demos` on the project site, which builds the same deck in the tab and then **shows you
  the slides**, by reading the package back with
  [`pptx-html`](https://www.npmjs.com/package/pptx-html) and painting them as SVG. A
  download button is still there, and still goes through the browser runtime's own
  `writeFile`.

  This removes React, Bootstrap, sass-embedded and a second ESLint configuration from the
  repository, and leaves one toolchain (VitePress + Vue) building one site. The site's
  application code lives in a new top-level `www/`, so `docs/` stays what the docs kit
  governs: markdown content. See `www/README.md`.

  `demos/` is now only what its name says — clone the repo, run a script, get a `.pptx`.

  The browser test lane did not shrink: that page is the Playwright `demo` fixture, so
  `writeFile`'s object-URL path and the browser-vs-Node byte-identity comparison are still
  covered, now against the site's build rather than a second app's.

- **The quarterly-review showcase is Kestrel Analytics, in forest and brass.** Same eleven
  slides, same five masters, same charts, table and timeline — a different fictional company
  and a different palette, because `pptx-html` already uses "Meridian" as its own example
  and two neighbouring projects demonstrating the same fictional company is confusing.

  This is only visible if you build the showcase. The output file name changes from
  `Meridian_Q3_Business_Review.pptx` to `Kestrel_Q3_Business_Review.pptx`, and every key in
  the deck's exported `BRAND` is renamed rather than merely re-valued — a constant called
  `navy` holding a green is the exact thing that file's own doc comment argues against.

- **The test suite shares one module registry per worker (`isolate: false`).** `dist/` is
  over 1 MB of JavaScript, `text-*.js` alone is 610 KB, 78 test files import
  `dist/node.js` and 64 import `dist/read.js`, and Vitest's default isolation had all 236
  files re-evaluating that graph from scratch. It was the single largest line in the run.
  Measured across three paired runs at `VITEST_MAX_WORKERS=4`, alternated so machine load
  could not favour one setting, import time goes from 136.0s / 114.5s / 84.8s to
  34.4s / 30.6s / 22.0s, and takes 22 to 37% of wall clock with it. The 3.5x to 4x on the
  import column is the honest number; the wall figures overlap because the box was getting
  quieter. On the single-worker path CI runs, a full `verify` now reports 14.7s of import.
  Two module-level caches written to span files finally do: `test/validator.js` joins
  validation requests across files instead of spawning one .NET child per file, and
  `test/read/corpus.js`'s `irFor` memo stops being rebuilt per file.

  Isolation made cross-file state leakage impossible rather than merely absent, so two
  things stand in for it rather than a hope that nobody leaks. `test/setup-globals.js`
  resets `setDiagnosticHandler` after every test, the one process-global the library owns,
  and the deliberate leak/detector pair in
  `test/regression/api/global-state-reset.test.js` fails when that setup file is removed.
  `sequence.shuffle.files` randomizes file order, so an order dependence fails instead of
  hiding behind a stable one; four full shuffled runs pass 3337 tests each. Tests *within*
  a file stay in source order deliberately, because `captureDiagnostics()` and the
  warn-capturing schema fixtures rest on it. The other module-level state under `src/` is
  idempotent caching, which is better shared than rebuilt.

- **Test helpers live in one module each, and four contracts widened on the way through.**
  Every helper involved had been written out again in seven to twenty-six files, and the
  copies had drifted the way copies do. `test/read/corpus.js` owns the read fixture corpus,
  `test/read/opc.js` the relationship assertions, `test/regression/chart/chart-parts.js`
  the chart part readers, and `test/helpers.js` `bytesEqual()`, `throws()`, `partBodies`
  and the 67 bytes of 1x1 transparent PNG that had been pasted under six different names,
  which made a grep for "the tests' image" miss most of them. corpus.js importers go from
  30 to 52; twenty-eight files had defined `openFixture` inline, which corpus.js had
  exported all along with zero importers. That is not only tidying, because a file that
  builds its own fixture path never loads corpus.js and so sits outside the `MIN_CORPUS`
  floor there, the check that stops a corpus which has silently resolved somewhere else
  from turning every invariant into a loop over nothing. Twenty-six files were outside it.
  None are now.

  `defineRegressionSuite` takes exactly two arguments. The three-argument form's provenance
  tag was destructured out and then dropped, so thirty-six suites recorded where their
  regression came from in a string no reporter ever printed; those tags are in the suite
  name now, and a second positional argument is an error rather than something ignored.
  Its `fn` goes to vitest as-is, which takes `helpers.js` off the top of every regression
  failure's stack, and per-case `skipIf` / `runIf` / `skip` / `only` / `todo` / `fails` /
  `concurrent` / `timeout` fields let a case reach the modifiers it could not get at from
  inside a plain array.

  The shared untouched-parts assertion is stronger than the dozen longhand loops it
  replaces, in two ways that each closed a live hole. It fails when it compared nothing, a
  loop whose filter stops matching being indistinguishable from success in a reporter. And
  a part missing from the output is a failure rather than a skip: `remove-slide.test.js`
  skipped absent parts outright, so any part vanishing, including one the removal had no
  business touching, read as a pass. It names the four parts a removal is expected to take
  and asserts that set exactly. Four checks in `animations-transitions.test.js` widen while
  passing through, having filtered to `ppt/slides/*` and left the rest of each package
  unasserted; load then save is byte-identical across every part of all 44 committed
  fixtures, so the narrowing bought nothing and hid the remainder.

  `roundtrip.test.js` named its decks in a five-name literal, and `docs/testing.md` asked
  for a manual edit to that literal when promoting a fixture. Nobody made it. The corpus
  reached 44 while the OPC contract kept being proved against the same five, so the decks
  that actually stress it (chartEx, model3d, math-omml, embedded fonts, av-media, modern
  comments) were never round-tripped there at all. It reads `fixtureNames` now, promotion
  is the only step there is, and 30 cases become 274 at a cost of about 22s. The widening
  found something on its first run: `bar-chart-data-labels.pptx` fails "saved output passes
  the validator" **as committed**, before this library touches it, with three Microsoft365
  errors inside PowerPoint's own chart `c:extLst` that the SDK does not model. So the
  assertion is now "a round trip introduces no *new* validator errors", which stays
  strictly stronger everywhere else (an empty verdict before still demands an empty verdict
  after) and avoids the choice between excluding that fixture and implying the library
  caused what PowerPoint authored.

  `test/schema-cases.js` gains the only two cases in the file that expect a non-zero error
  count. Every other case asserted zero, so the tier could not tell "this deck is valid"
  from "the validator reported nothing", and a keying bug in the batcher, a spawn failure
  the JSON parse swallows, or an output-shape change on a version bump would have turned
  the whole tier green while proving nothing, with no other step in `verify` noticing. One
  case perturbs a freshly built deck with an undeclared attribute on `p:sp` and demands
  exactly one `Sch_UndeclaredAttribute` at the expected part and XPath; the other feeds
  bytes that are not an OPC package and demands the package-level row, which is the
  property the batcher's old "absent means clean" shortcut rested on. Both were confirmed
  to fail with `validateBuf` stubbed to return `[]`, and both assert on the id and the
  XPath, never on the description, so an upstream rewording cannot break them. The failure
  formatter was also dropping two of the CLI's five fields, the XPath that names the
  offending element and the stable machine id; it prints all five now.

  One consequence to know about, recorded in `AGENTS.md` and `docs/testing.md`: `verify`
  alone no longer covers the script round trip over the corpus. `script-roundtrip.test.js`
  and `script-standalone.test.js` each carried a byte-equivalent copy of what
  `script:roundtrip:all` already runs, same corpus and same `diffDeckIr`, so `verify:full`
  paid for the identical 44-deck two-tier subprocess sweep twice. The script keeps that
  job, being the more capable copy, and the suites keep what the round trip rests on and
  cannot itself establish. Run `pnpm run script:roundtrip:all` before pushing changes to
  `src/script/`.

- **The commit-message gates come from `shbernal/lefthook-rules`, not a local script.**
  `scripts/check-commit-msg.mjs` encoded a rule that is not about this project: a heredoc
  delimiter leaking into a commit subject when the shell dialect is wrong. ts-xlsx carries
  the same damage, two commits of it. Both gates are lefthook remotes now, pinned to `v1`
  with a 24h refetch. `no-shell-quoting-leak` is not a transcription of the script it
  replaces. It drops the leading-backtick test, which would reject three legitimate
  subjects here that open with a code span and never fired on a damaged one, and it adds
  the closer test the script lacked, since two of the five damaged commits end with a
  second delimiter line and the script caught them only because the opener leaked as well.
  It reads every line rather than the subject alone, truncating at the `git commit -v`
  scissors line first. Over this history it rejects those five commits and none of the
  other 4356. `no-ai-attribution` arrives with it, and a remote job costs about 3 ms
  against about 19 ms for the same logic in a Node process, which a commit-msg hook pays
  every time.

  Two properties of the mechanism are worth knowing. A remote that cannot sync prints a
  warning and still exits 0, so `prepare` survives an offline install and the rules simply
  do not run. A `configs:` entry the tag does not hold is silent. After a fresh clone,
  confirm both job names appear in a commit.

- **The OOXML schema MCP runs locally over stdio.** Both MCP configs point at
  `mcp-server-ooxml` rather than the hosted `https://api.ooxml.dev/mcp` endpoint. The
  schema graph ships inside the package, so a lookup needs no account and no network round
  trip, and the same question returns the same answer every run. Codex gets a 60s startup
  allowance because the first run downloads the tarball.

  It is not the same tool surface, so the docs that route agents to it had to change with
  it. It gained `ooxml_values`, `ooxml_diff_profiles` and `ooxml_explain`, and it lost the
  prose half entirely: no spec PDFs, no `ooxml_section`, and no OPC part, content-type or
  relationship catalogue. `AGENTS.md` and `docs/ooxml-agent-context.md` now say outright
  that those questions fall through to `microsoft_learn` and then to web search, rather
  than leaving an agent to discover the dead end. `ooxml_search` also narrowed to a
  substring match on symbol names, which the Annex D note records: a semantic query would
  not have worked there even if the geometry addenda were indexed.

- **`scripts/` is held to `noImplicitAny`, annotated in JSDoc.** `tsconfig.scripts.json`
  relaxed the flag while the tree was unannotated. Paying what it wanted, 291 annotations
  across 24 files, is the cheaper half of the "migrate `scripts/` to TypeScript" question:
  the annotations buy the checking, and the file extension buys nothing on top of it while
  costing 49 path citations in `docs/` and every glob in the lint, format and tsconfig
  lists. Types come from the library itself, `import('../dist/read.js').Slide` rather than
  a restatement of its shape.

  The flag was not bookkeeping. It surfaced six silent defects.
  `scripts/append-ceiling.mjs` reported a false loss, its bullet-glyph probe reading
  `paragraph.bullet`, an accessor the read model replaced with `bulletDetail`, so it
  compared `undefined` and called the construct lost; it reports 23 of 24 surviving now,
  the one failure being the documented `placeholder idx` loss.
  `scripts/read-blindness-census.mjs --all` never censused a notes slide, because it
  reached for a `NotesSlide.element_` that does not exist and its guard was therefore
  always false, while `docs/testing.md` already advertised `--all` as covering notes.
  Three dead `??` fallbacks sat on properties the types do not carry
  (`hyperlink.slidePartName`, `GradientStop.hex`, `ResolvedColor.value`).
  `typeof slide.notesSlide === 'function'` evaluates a getter and tests its result, so
  that arm was unreachable and the getter ran twice on the way to the arm taken.
  `scripts/docs-init.mjs` could render every template with `undefined` as the project
  name, from a `.split('/').pop()` TypeScript knows can miss. And
  `scripts/powerpoint-com-smoke.mjs` dropped a signal kill: `close` passes `null` for a
  signalled child, which flowed into a `code !== -1` retry test and made a killed cscript
  look like a clean run.

  `@types/istanbul-lib-coverage`, `-lib-report` and `-reports` are new devDependencies,
  since `scripts/coverage-merge.mjs` imports three untyped Istanbul modules and the
  alternative was a hand-written shim that could mistype the API silently; with real types
  that file went from 23 errors to 2. `noPropertyAccessFromIndexSignature` stays off, with
  its reason recorded: it is house style about `obj.foo` against `obj['foo']`, not
  correctness, and these scripts read a lot of JSON that is `Record<string, ...>` by
  nature. `.oxlintrc.jsonc` and `scripts/README.md` both asserted this tree was
  deliberately unannotated, so both are rewritten against a fresh measurement rather than a
  guess: with the type-aware rules restored, `scripts/**` reports 289 findings against the
  recorded 1276, and 285 of the 289 are still `no-unsafe-*`. The residue is `JSON.parse`
  and untyped imports, not unannotated parameters, so the rules stay off for a different
  reason than before. `test/**` is unannotated and unmeasured since, and is a separate job.

### Added

- **Font collections (`.ttc`/`.otc`) are now readable by the measured-fit metrics.**
  `parseFontMetrics` and `pptx.registerFontMetrics` previously threw `Unsupported
  OpenType signature ttcf` on a collection, which is how most of the fonts a consumer
  would reach for to measure East Asian text ship on Windows: MS Gothic, Yu Gothic,
  SimSun, Microsoft YaHei, Microsoft JhengHei, MingLiU and Nirmala UI are all `.ttc`, and
  so is Cambria, so this was never only a CJK gap. Measured fit on those faces fell back
  to the cmap-less heuristic, or did not engage at all.

  A collection holds several fonts, so one has to be chosen. Both APIs take a new `font`
  option, either a 0-based index or a name matched case-insensitively against the family,
  full, and PostScript names; `registerFontMetrics` additionally uses its own `face`
  argument as the selector when `font` is omitted, so the common call needs nothing extra:

  ```ts
  await pptx.registerFontMetrics('MS PGothic', 'C:/Windows/Fonts/msgothic.ttc')
  await pptx.registerFontMetrics('Cambria Math', 'C:/Windows/Fonts/cambria.ttc', { font: 1 })
  ```

  A selector matching nothing **throws** (`font/collection-face-not-found`,
  `font/collection-index-out-of-range`) rather than falling back to the first font, and
  the message lists what the file holds. Falling back is the failure worth refusing:
  measuring the wrong member produces perfectly plausible numbers that nothing downstream
  can question, and MS Gothic against MS PGothic is a 26% difference on Latin advances.
  The selector means the same thing for a plain `.ttf`, which is a one-entry list, so a
  wrong index or an unmatched name is an error there too rather than being ignored.
  Registering a plain `.ttf` under any `face` name is unchanged: only a collection is
  name-selected.

  `ts-pptx/measure` gains `listFontFaces(bytes)` (every font in a file, with its `name`
  table identity; a plain `.ttf` is a one-entry list) and `isFontCollection(bytes)`.

  Grounding: the unwrap rests on the claim that a member's table records carry offsets
  absolute to the file, which is how members share one `glyf`. A self-generated fixture
  cannot test that, since a builder and a reader can agree on a format real files do not
  follow, so the genuine Windows collections are checked against advances read by WPF's
  `GlyphTypeface` (38 faces across 15 collections, `windows-collections.oracle.json`).
  A collection synthesized from the repo's own Silkscreen files carries the same claims
  on every platform, asserting that a member's metrics equal exactly the same font parsed
  as a plain `.ttf`. See `docs/measured-text-fit.md` ("Font collections").

- **A font the parser rejects now raises a classified `MediaError`** with the new
  `font/parse-failed` code, instead of opentype.js's bare `Error` escaping the taxonomy.
  A consumer catching library failures by class and `code` could not previously catch a
  corrupt or unsupported font file that way.

- **Measured fit now reports the code points a registered font cannot render**, rather
  than measuring them silently. `measureText()` returns them in a new
  `uncoveredCodepoints` array (sorted, deduplicated), and `applyMeasuredFit` warns once
  per export with the new `measure/uncovered-codepoints` diagnostic, naming the face and
  the code points.

  This is the one approximation in the model that does not err safe. Everywhere else the
  numbers are deliberately conservative — raw advances with no kerning, the taller of two
  line heights, a width safety factor that wraps a hair early — so a wrong answer is a
  too-tall box. Font fallback is different: PowerPoint substitutes another face per code
  point and lays the run out in *that* face's advances, while this model has no fallback
  and charges the registered font's `.notdef`, a single flat number unrelated to the glyph
  that paints. Malgun Gothic's `.notdef` is 0.663 em — wider than the 0.5 em halfwidth
  Katakana it lacks, so those gain a harmless phantom line, and narrower than the 1.0 em
  Plane 2 ideographs it lacks, so a run of *those* measures short: 24 of them in a 150 pt
  box lay out on 2 lines here and 3 in PowerPoint. A short measurement is a `fit:'resize'`
  box baked smaller than its text, which is the overflow the resize path has no safety net
  for.

  Only faces that **are** registered are audited: an unregistered face measures through
  the cmap-less average-advance heuristic, which has no coverage to report and is already
  surfaced in `approximatedFaces`. The audit (`collectUncoveredCodepoints` in
  `src/measure/font-metrics.ts`) is shared by the layout-time query and the export pass,
  so the two cannot diverge. Nothing about the measured numbers changed — this is the
  signal that says when not to trust them.

- **`Presentation.importSlides(requests)`, batch slide import across one or more loaded
  sources** (#19). Stitching a deck from several sources needs each imported page at a
  specific position in the *final* slide list, and needs the whole stitch to succeed or
  fail as one. `importSlide` in a loop can leave a half-stitched deck when request three
  of five fails on a size mismatch; `importSlides` validates everything up front — pages
  exist, output positions are unique and within the final list, sizes match, and a
  read-only dry run of the copy proves every part it would reach is present — so a
  rejected batch leaves the target byte-identical whichever rule rejected it. The returned
  array is parallel to `requests`. It also gives the
  batch's `slide → slide` links a rule: a jump link on a selected page must target another
  selected page (or one an earlier import from that source already brought across) and is
  rewritten to the fresh partnames, so importing page 3 of 10 neither drags pages 1–2
  across as dependencies nor strands the link — `appendSlides`' `import/unresolved-slide-link`,
  enforced natively. One request is one output page, so naming the same source page in
  several requests asks for that many independent copies of it — the page part is the one
  thing an import never shares, and a duplicated page's jump links stay inside the batch
  (pages duplicated together are copied in lockstep, so a duplicated linked pair becomes
  two self-contained pairs). Pages come across under `'copy'` theme semantics; notes,
  embedded fonts and `rescale` have no batch spelling yet, so reach for `importSlide` when
  you need one of those.

- **`compose()` on the showcase modules**, beside the existing `build(outFile)`: it
  assembles the deck and returns the presentation, having written nothing. The preview needs
  the bytes; `pnpm demos:build` needs a file. `build` is now `compose` plus a destination.

- **`pnpm run typecheck:site`**, and `docs/tsconfig.json` grew a runner. It covered
  `docs/.vitepress/**` on paper and was executed by nothing; it now also covers `www/**/*.ts`
  and runs in `verify`, `check:static` and pre-push. `.vue` files stay outside it — `tsc`
  does not read single-file components — which is why the demos page's logic is a plain
  `.ts` module and the component is markup around it.

- **`pnpm run lint:chars`, which keeps em dashes out of what a reader sees.** The README,
  the site and the docs are the only text a reader of this project ever meets, and nothing
  checked what characters went into them. `charcheck` is that gate, wired into both hooks
  and `check:static`, and `charcheck.config.js` carries one rule per surface, because the
  surfaces differ in what reader-visible text even means. Markdown prose is scoped so a
  dash inside a fenced or inline code sample is not a finding: it is part of the sample,
  not something to reword. `www/**/*.vue` is scanned as markup, which reads template text
  and allowlisted attributes while leaving the component's comments and its stylesheet
  alone. The VitePress config and theme are scanned as strings only, since the nav labels,
  the tagline and the meta description are the sole reader-visible text in those files.
  Generated trees are ignored by path, so a scan of a built checkout agrees with a scan of
  a clean one.

  The gate arrived on a tree that already held 684 em dashes in `docs/`, which warned under
  a frozen `--max-warnings` while they were worked off. They are worked off, so the second
  severity and its `DOCS_CLEAN` allowlist are gone, every surface errors, and
  `--max-warnings 0` is on both `lint:chars` and the pre-commit hook so that a rule added
  at `warn` later fails rather than scrolling past in a hook's output. No hook passes
  `--fix`, because a fix is a guess about prose and rewriting a sentence on its way into a
  commit is not a guess a hook should make unsupervised. Reading that diff was not a
  formality either: `--fix` broke a bracketed label in three places, nested parentheses two
  deep in four more, converted only one half of five dash *pairs*, and turned a table cell
  whose entire content was a dash into `|: |`, twice over seven such cells. None of that is
  caught by a check.

  Clearing the backlog also caught the gate lying. Eleven genuine dashes were sitting
  behind a green run, every one of them the same shape: a dash ending a hard-wrapped line
  whose continuation began with an inline code span. A trailing `\s*` in the pattern
  matched the newline, carried the match off the end of the line and into the next node,
  and charcheck dropped the finding without a word. Filed as `shbernal/charcheck#16`,
  worked around with a horizontal-only `[ \t]` until the fix shipped, and back to `\s*` on
  0.2.3, which reports the longest match that fits inside the region. The pin is exact for
  that reason: below 0.2.3 this pattern under-reports, and it does so as a pass.

- **`checkDocsJson` gates both directions of the sidebar.** It validated nav to page and
  stopped there, which catches a page that was deleted or renamed and says nothing about
  one that was added and never listed. That is the failure that actually happened: three
  pages sat outside the sidebar with `docs:build` green, and finding them meant diffing
  the tree against `docs.json` by hand. Page to nav is checked now, with two exemption
  mechanisms, because the two exempt things are not the same shape. `NAV_EXEMPT` is by
  exact key, currently `doc-index`, which `scripts/docs-index.mjs` generates for the
  llms.txt build rather than for anyone to navigate to; listing an exempt page *in* the
  nav is itself an error, so an exemption that stops being true reports rather than
  silently granting slack, the same reasoning as the allowlist in `scripts/path-refs.mjs`.
  `GENERATED_TREES` is by prefix, currently `reference/api/`, where TypeDoc emits a few
  hundred pages behind one sidebar link, which is why it cannot be an exact-key exemption:
  `reference/api/index` stays eligible for the nav direction. The prefix form is also what
  keeps the verdict a property of the repo rather than of the machine, since that tree is
  gitignored and so is absent in a fresh checkout and present in any working copy that has
  run `docs:api`. Skipped by prefix, both states agree.

### Fixed

- **Two copies of one page shared the parts that page owned, and PowerPoint refused the
  deck.** A page copy shares the deck-wide graph underneath it — layout, master, theme,
  images — and that sharing was applied to everything below the page, including the parts
  PowerPoint treats as belonging to exactly one slide. Duplicating a page with a chart or a
  SmartArt diagram therefore wrote a package with two slides resolving to one chart part or
  one diagram set, and PowerPoint would not open the file at all: `0x80070570`, the whole
  deck rejected rather than the duplicate page. The schema validator accepts such a package
  — nothing in ECMA-376 says a chart part has one referrer — so CI was green the whole
  time. Every way of duplicating a page was affected: `cloneSlide`, `importSlide` of one
  source page twice, and (once it allowed the same page twice) an `importSlides` batch — as
  was `importShape`, one level down, where carrying the same chart or SmartArt frame onto
  two slides pointed both frames at one part.

  A page copy (or a carried shape) now takes its own copy of what it owns: charts and
  chartEx, the five SmartArt parts, OLE embeddings, tags, comments, the notes slide. Ownership is transitive
  — a chart's embedded workbook and `chartUserShapes` drawing come with it, measured
  against desktop PowerPoint, which refuses a shared workbook exactly as it refuses a
  shared chart — while media blobs, fonts and deck furniture stay shared, which is what
  PowerPoint itself does. The classification lives in `src/read/api/ops/page-owned.ts` as a
  list of what may be *shared*, so a relationship type nobody has classified is copied: a
  wrongly shared part is a deck nobody can open, a wrongly copied one is some duplicated
  bytes.

  Two behaviours change beyond the corruption. A cloned page now gets its own notes slide,
  wired back to the clone, where before the clone and the source shared one notes part (the
  caveat the `cloneSlide` docs used to carry). And `importSlide`/`importSlides` no longer
  dedupe a chart or diagram across repeated imports of the same page, so such a deck grows
  by one chart subtree per copy.

- **`cloneSlide` lost every relationship of a page imported in the same session.** The
  clone copied the source's `.rels` *part bytes*, and a page brought in by `importSlide`
  holds its relationships in memory until the deck is saved: there was no such part yet, so
  the clone came out with none at all — a slide whose `r:id`s resolved to nothing, its
  chart with no chart part behind it and its layout unbound. The clone now copies the live
  relationship set, which is the same thing for a page loaded from disk.

- **`importSlide` twice from one loaded source wrote a package PowerPoint refuses to
  open** (issue #18). Copying a part out of a source deck is idempotent per source
  package, which is right for a theme, master, layout or image — a second copy is
  waste. It was also being applied to the slide part, which is the one thing an
  import duplicates on purpose. So the second `importSlide(source, i)` with the same
  loaded `source` returned the first copy's partname, and the deck ended up with two
  `p:sldId` entries and two relationships naming a single slide part: `0x80070570`,
  "the file or directory is corrupted and unreadable", with the whole package
  rejected rather than just the duplicate. Nothing in the API said so first —
  `slides.length` counted the entry that had no part behind it — so the failure
  surfaced only when someone opened the deliverable. Only `theme: 'copy'`, the
  default, was affected; `preserve` and `restyle` always allocated a fresh part. A
  page an `importSlides` batch had already brought across hit the same registry, so
  batch-then-import failed the same way.

  The page now goes through the selection plan `copyPart` already honoured for a
  batch: it is materialized fresh on every call while everything under it (layout,
  master, theme, media) stays shared. Importing one source page twice yields two
  independent copies over one master, which is what makes a before/after pair of the
  same page possible without loading the source deck twice.

  Wiring a slide part into `p:sldIdLst` a second time now throws
  `InternalError('slide/part-already-in-deck')` rather than writing the deck. No
  caller can reach it with valid input — the guard exists because this class of
  defect is invisible until the file is opened, and a thrown error at the call site
  is worth more than a clean run and a corrupt deliverable.

- **Measured fit put Chinese and Japanese text on one line too many.** The wrap
  simulator broke lines only at whitespace, so a CJK run tokenized as a single
  unbreakable word: it fitted beside nothing, moved to the next line whole, and left
  the rest of the current one empty. Harmless while such a run sat alone on its line
  (the over-long-token fallback packed it identically), wrong the moment anything
  preceded it. The inflated line count then propagated: `fit:'shrink'` baked a
  `fontScale` PowerPoint would never have chosen, and `measureText`/`overflowsBox`
  reported vertical overflows that were not there.

  `src/measure/text-fit.ts` now treats Han, Kana, Bopomofo, the fullwidth and halfwidth
  forms, the compatibility blocks and the Plane 2/3 ideograph extensions as per-character
  wrap opportunities, matching how PowerPoint lays those scripts out.

  **Hangul is deliberately not in that set.** UAX #14 permits breaking Korean between
  syllables; PowerPoint does not do it, because Korean is written with spaces between
  words. A Hangul run still breaks when it is longer than a line, but through the
  over-long-token character-wrap fallback rather than a break class. Including it
  would have *under*-reported the line count for Korean, which is the direction that
  overflows, and which the resize path has no safety net for.

  Both halves are pinned to a new PowerPoint-authored fixture,
  `test/read/fixtures/autofit-cjk-wrap.pptx` (11 boxes, Malgun Gothic, one claim
  each), read by `test/read/cjk-line-breaking-oracle.test.js`. Reverting the fix fails
  three of its cases; adding Hangul back fails a fourth. Two limitations are recorded
  there rather than fixed — PowerPoint's kinsoku rules (`。`/`、` hang past the right
  inset instead of starting a line) and its font fallback for code points the named
  face lacks. Kinsoku costs nothing: same line count, narrower widest line, so the
  height stays conservative. Font fallback does not have that property, and is now
  reported through `uncoveredCodepoints` and a warning (see Added) rather than only
  documented. See `docs/measured-text-fit.md`.

  Thanks to **@flyisland** for the report and the original fix (#20, #21).

- **Four source files were being tracked as binary, and it had already cost us.**
  `src/diagnostics.ts`, `src/measure/font-metrics.ts`, `src/script/fidelity.ts` and
  `www/demos/deck-preview.ts` each held a **literal `U+0000` byte** inside a template
  literal used as a map-key separator — `` `${code}<NUL>${message}` `` and friends. Git
  classifies such a file as binary, which is why the last two showed up as `Bin 0 -> 4739
  bytes` and `Bin 4938 -> 6949 bytes` in their own commits rather than as readable diffs.

  The bytes are now written as the `\0` escape. Same value at runtime, and the files are
  text again: they diff, they blame, `git log -S` reaches them, and — the part that
  matters — `grep` stops skipping them silently.

  Two stale references had been hiding behind exactly that, both invisible to the passes
  that should have caught them. `src/measure/font-metrics.ts` still opened with
  `PptxGenJS: Write-side font metrics`, the last upstream branding anywhere in `src/`,
  which the rebrand content scrub could not see; and it still cited
  `PLAN-measured-text-fit.md`, consolidated into `docs/measured-text-fit.md` back in
  `604ff7cf`, which `path-refs:check` does not catch because a citation needs a `/` to be
  recognised as a path. Both now say what is true.

- **The desktop-smoke skill told agents to run a generator that no longer exists.**
  `.agents/skills/powerpoint-desktop-smoke/` walked through `cd demos/node && node demo.js
  All` and a per-feature-group `node demo.js <Group>`, against
  `demos/node/output/TsPptx_Demo_*.pptx`. That generator went with the upstream-era demos;
  `demos/node/` is one HTTP-streaming example now. The workflow is rewritten around
  `pnpm demos:build` and `demos/showcases/output/`, and the bisect step no longer promises
  a feature-group split there is no generator for. `path-refs:check` does not reach
  `.agents/`, so nothing had flagged it.

- **Two config comments described files that had been renamed.** `.oxfmtrc.jsonc`
  justified ignoring `docs/docs.json` by citing `scripts/docs-new.py` and
  `scripts/docs-init.py` and "both python writers"; they are `.mjs`, and write with
  `JSON.stringify(config, null, 2)`. The reasoning was still correct, the evidence for it
  was not. It also still ignored `demos/browser/js/*`, a tree not in this repository.

- **`src/types/` had never been formatted.** `.oxfmtrc.jsonc` carried `types/` among a
  group of build-output directories, next to `dist/` and `build/`, mirroring the `/types/`
  in `.gitignore`. It lost the anchor on the way across, and oxfmt's patterns are
  gitignore-flavoured, so an unanchored directory name matches at *any* depth. There is no
  `types/` at the root and never has been, so the pattern had only ever matched one thing:
  the library's entire public type surface. Seventeen files the formatter had never read,
  while `format:run`'s `"src/**/*.ts"` glob says plainly that it covers them. The ignore
  list is documented as a safety net rather than the definition of what gets formatted;
  here it was quietly the definition, and it won.

  Four of the seventeen had drifted, and `src/types/style.ts` carried a 129-column line
  against a `printWidth` of 120, visible in any diff of that file for as long as it has
  existed and invisible to every gate. Reformatted here: three union types split one
  member per line, one import collapsed onto one. No semantic change, `typecheck` is
  clean, and the emitted `.d.ts` is unaffected. The header already recorded this hazard in
  its other direction, `tsconfig*.json` matching at any depth and reaching workspaces the
  repo does not format, caught by `format:check` failing on three untouched files. That
  one was loud. This one could not be, because an ignored file looks exactly like a
  well-formatted one.

- **Two claims in the demos READMEs had gone stale in a way `path-refs:check` cannot
  see**, since the file each one cites exists and what the text says *about* it is what
  rotted. The showcases README told the next author to register a third deck in
  `build.mjs`; `SHOWCASES` moved to `lib/showcases.mjs` when the byte-identity harness
  started enumerating decks from it, so following that instruction would have built the
  deck by one list while leaving it out of the gate, the exact drift the move was made to
  prevent. It now says where the registry lives, and why that is not an arbitrary file
  choice.

  And `demos/common/` is not read only by the Field Notes deck. Three of its twelve
  surviving assets are read by the test suite: `cc_logo.jpg` by four regression suites and
  by a fixture-authoring script, `logo_square.png` and `lock-green.svg` by both sides of
  the browser lane. The table called the directory "shared images and media the Field
  Notes deck draws on" one commit after that same directory was pruned on the rule "no
  showcase references it", a rule which, applied literally to what was left, takes the
  browser lane's fixtures with it. There is a table of who actually reads what now, and an
  instruction to grep the repository rather than `showcases/`. A `<deck>/data.mjs` is also
  the quarterly review's alone, not the shape both decks have; Field Notes carries its
  captions inline.

- **TypeDoc emitted nineteen warnings on every docs run.** Nine came from
  `{@link PresentationCore.measureText}` and
  `{@link PresentationCore.registerFontMetrics}`. `PresentationCore` is the
  default-exported base class and is not itself an entry point, so the target resolved to
  a symbol absent from the docs, and the miss multiplied across `index`, `node` and
  `browser`. They point at `TsPptx` now, the subclass each entry point exports, which
  resolves to a documented member and renders a real link. TypeDoc's own suggestion,
  mapping the names to `"#"` in `externalSymbolLinkMappings`, silences the warning and
  leaves a dead anchor.

  Eight more were symbols a documented type references but that are not documented
  themselves, plus one stale `intentionallyNotExported` entry for `ImageBaseProps`, which
  is exported now. For the `st-enums.ts` tuples this is deliberate: `types/table.ts`
  re-exports the four `a:cell3D` *type aliases* and keeps the backing tuples internal, so
  the fix is to declare that intent rather than widen the public surface. The last warning
  was an `@example` on a member of the `SlideMasterObject` union, where TypeDoc allows no
  block tags; it is prose now, and says why.

  `docs:api` reports zero warnings. A full `docs:dev` still prints esbuild's
  `Unrecognized target environment "es2024"`, which is not ours to fix: VitePress 1.6.4
  pins Vite 5.4, whose esbuild 0.21.5 predates ES2024 target support, and it reads the
  root tsconfig when bundling the VitePress config. TypeDoc still documents three of the
  ten published entry points, and covering the other seven surfaces about forty latent
  link warnings, so that wants its own pass.

### Removed

- **41 of the 53 files in `demos/common/` — about 34 MB of 37 MB.** They were the upstream
  demos' feature-checklist props (`starlabs_*`, `title_bkgd*`, `cc_*`, `fediverse_*`,
  `krita_*`, `sample.{aif,avi,m4v,mov,mp3,mpg,wav}`, `earth-big.mp4`, `base64Images.js`,
  and the rest), and nothing referenced them: not a showcase, not a test, not the browser
  harness. What is left is the 12 files the Field Notes deck draws, the regression suite
  loads, and `scripts/browser-harness-server.mjs` serves.

  `images/image2.jpg` was the last one out, and it survived the first sweep because the
  grep that found the other forty matched it: `src/script/print/common.ts` documents the
  read half's asset naming (`image1.png`, `image2.jpg`, and so on), so the name appears in
  the tree without anything loading the file. That is a collision, not a use. After the
  removal the only occurrence of the string anywhere in tracked source is that doc comment.

- **`tslib` and `@arethetypeswrong/core` from `devDependencies`.** No tsconfig sets
  `importHelpers` and nothing in `dist/` imports tslib; it arrived in a bulk dependency
  bump and was never used. `@arethetypeswrong/core` is a direct dependency of
  `@arethetypeswrong/cli` at an *exact* pin, so declaring it again here was a second,
  looser constraint on the same package whose only possible effect was to drift out of
  step with the pin the cli actually wants.

- **`demos/.prettierrc.json`.** Prettier is gone; oxfmt does not read it, `demos/**` is
  oxlint-ignored, and `format:run`'s globs do not reach it. It configured nothing.

- **`yaml` and `fast-xml-parser` from `devDependencies`.** `yaml` arrived for
  `scripts/upstream-signals-ledger.mjs`; that script and its test are gone and upstream
  tracking is retired, so nothing in the repo imports it. The comment in
  `scripts/docs-frontmatter.mjs` explaining that the hand-rolled frontmatter parser
  deliberately does not use it now states the standing constraint instead, since "the
  `yaml` devDependency is deliberately not used here" stops parsing once there is no such
  devDependency: this repo carries no YAML library, and do not add one for that file. A
  real parser would reject frontmatter the hand-rolled one tolerates, which is what would
  turn `docs:check` from a lint into a gate on YAML pedantry.

  `fast-xml-parser` had exactly one consumer,
  `test/read/fixtures/authoring/extract-autofit-calibration.mjs`, and `@xmldom/xmldom` is
  already a runtime dependency, so the swap adds nothing and removes 11 packages from the
  lockfile. That script still parses XML itself rather than going through this library's
  read model, which covers every field it reads, and the reason is worth not
  re-litigating: `autofit-calibration.json` is an **oracle**, held against the shrink
  solver in `src/measure/text-fit.ts` by `test/read/autofit-calibration-oracle.test.js`,
  so deriving it through the reader would make the oracle a function of the code it exists
  to judge. It would also make `dist/` a prerequisite for regenerating a fixture, and this
  is the one recipe in `authoring/` that needs nothing but Node. The three replacement
  helpers walk *direct children* rather than searching descendants, which is what every
  field lookup here means (`p:spPr` then `a:xfrm` then `a:off`); the one place recursion
  is correct is the `p:sp` sweep, where a group nests shapes. The regenerated table is
  byte-identical to the committed one across all 149 cases and every field, so
  `autofit-calibration.json` is absent from that diff rather than buried in a 15,374-line
  whitespace churn the pre-commit formatter used to flatten back.

- **`tools/data2chart.html` and its stylesheet.** A gitbrent-era standalone page: paste
  tab-separated data, get a snippet building a line chart, optionally `eval()` it in the
  tab to produce a deck. Every load-bearing piece of it is gone. It loads
  `dist/pptxgen.bundle.js` from a CDN, an artifact `docs/runtime-and-package-support.md`
  lists as something this package deliberately does not ship and
  `scripts/package-smoke.mjs` asserts the absence of. The code it emitted would throw even
  if the bundle loaded, since it wrote the three-argument
  `addChart(pptx.charts.LINE, data, opts)` signature against a `pptx.charts` enum bag
  where the API is `addChart(data, { type })`. Its `images/favicon.png` has no counterpart
  in the repository, and nothing anywhere referenced the page. Rebuilding it would have
  meant a second standalone browser app one commit after the last one moved onto the site,
  for a payload of about 35 lines of TSV transposition. Deleting it also makes
  `runtime-and-package-support.md`'s "the legacy upstream browser demo for that workflow
  is not included in this repository" true, rather than contradicted by a file sitting in
  the tree.

- **Ignore patterns for seven paths nothing writes**, from `.oxfmtrc.jsonc`
  (`.docusaurus/`, `static/`, `libs/`, `build/`, `src/bld/`) and `.gitignore`
  (`bower_components/`, `.nyc_output/`, `/types/`, `src/bld/`). The first two are the
  layout of a site generator this repository replaced with VitePress, the next two a
  package manager and a coverage tool it does not use, and the remainder match no
  directory at any depth. Verified inert rather than assumed: bare oxfmt considers 600
  files before and after, and `git status --untracked-files=all` reports nothing newly
  visible. The reason to bother with patterns that cost nothing to run is the `src/types/`
  fix above. A list that mixes live entries with speculative ones cannot be read, because
  a pattern's presence stops being evidence that it matches something, and the one entry
  that did match, and matched the wrong tree, read as unremarkable for as long as it sat
  among six that matched nothing. Both files now say to add a path when a tool writes it,
  rather than in case one might.

## [3.2.0] - 2026-08-10

### Added

- **`bullet: 'inherit'`, the state omission cannot spell** (#15). `TextBaseProps.bullet` had
  two states and three meanings: a bullet, or `false`/omitted — and both of the latter emitted
  an explicit `<a:buNone/>` plus `indent="0" marL="0"`. There was no way to author a paragraph
  that states *nothing* about its bullet. That matters because bullet properties resolve down
  the `a:lstStyle` → placeholder → layout → master chain, and `a:buNone` **overrides** that
  chain rather than deferring to it: a paragraph inheriting a bullet from its layout's
  `a:lvl1pPr` lost it, and one inheriting no bullet still stopped tracking the master, so a
  later edit there no longer reached the slide. The `indent="0" marL="0"` written beside it
  flattened an inherited hanging indent to zero in the same stroke — a second, separate loss on
  the same element. `bullet: 'inherit'` emits neither a bullet child nor `a:buNone` nor the
  margins.

  Omission keeps meaning `a:buNone`, and that is deliberate: it has meant that since the writer
  existed and every deck authored against it depends on the default. Same resolution as
  `fill: { type: 'inherit' }` (#10) — name the state that had no name, leave alone the one that
  has. This is that fix one element down, and #14's mirror image: there the *explicit off*
  needed a spelling distinct from silence, here the explicit off is the only thing that could
  be said.

  Two sites downstream of the emitter tested `bullet` for **truthiness**, which a truthy string
  that draws nothing would have broken: the line grouping (a bullet starts a new paragraph)
  would have split every run into its own paragraph, and the leading-glyph strip (which removes
  a literal `•` so it is not drawn twice beside the `a:buChar`) would have eaten a real
  character from text that emits no bullet to duplicate. Both now ask whether bullet markup is
  actually emitted.

- **`paraMarginLeft` and `paraIndent`, the other half of that element.** #15's report named
  the `indent="0" marL="0"` beside the `a:buNone` as "a second, separate loss on the same
  element", and `bullet: 'inherit'` closed it only for the state that says nothing. The
  attributes themselves still belonged to the bullet: `a:pPr/@marL` and `@indent` had no option
  of their own, so `bullet: false` could not suppress a bullet without also flattening an
  inherited hanging indent to zero, a drawn bullet always re-hung the first line by the
  writer's 27pt default, and a first-line indent (a *positive* `@indent`, the prose form) was
  unauthorable in every state. `indentLevel` is a different fact — it writes `a:p/@lvl`, the
  discrete outline level — and `bullet.indent` reached `@marL` only by drawing a bullet with
  it.

  Both options take points and both admit `'inherit'`, which writes no attribute at all and so
  leaves the margin to the `a:lstStyle` → placeholder → layout → master chain. Omission keeps
  meaning what it has always meant — whatever the `bullet` state writes — so no existing deck
  moves, and the third state is spelled rather than assumed, exactly as one element up. An
  explicit value wins over the bullet default in every state. Out-of-range values are clamped
  to `ST_TextMargin` (0–4032pt, unsigned) and `ST_TextIndent` (−4032–4032pt) with a warning,
  under the new `text/paragraph-margin-out-of-range` and `text/paragraph-indent-out-of-range`
  diagnostic codes, since PowerPoint reports an out-of-range value as needing repair.

- **`ts-pptx/script` carries a paragraph's own margins, and `text.indent` is retired.** It was
  5/44 on the corpus and the largest note left on `a:pPr` once the bullet one closed; the
  standalone tier's note count falls from 713 to 705 and the template-anchored tier's from 419
  to 411. The mapper emits the read margin as a number, and `'inherit'` where the paragraph
  states none but its bullet state would otherwise write one — which is what makes an inherited
  margin survive rather than being replaced by a default. The reader did not move here either:
  `Paragraph.marginLeftPt` and `Paragraph.indentPt` already separated a stated margin from an
  absent one, so this was a missing write spelling, the same shape as the bullet note above.
  The retired note's exclusion entry was **empty** — neither IR carried the field, so the round
  trip compared two models both missing it — which means the margins are now verified rather
  than merely declared; a build that ignores the new options reports 36 undeclared differences
  across the corpus.

- **Fixed in the same pass: a paragraph's properties are the first run's, decided once.** The
  emitter retried `genXmlParagraphProperties` on each run of a paragraph until one produced
  non-empty XML. That was unreachable while every `bullet` state wrote something, and became
  wrong the moment an empty `a:pPr` was possible: a paragraph whose first run said `'inherit'`
  took its properties from a *continuation* run instead, which by convention states no bullet —
  so it got back the very `a:buNone` the first run asked to leave out, appended **after** that
  run's `<a:r>`, where a `pPr` is not allowed. No deck the byte-identity corpus authors moves.

- **`ts-pptx/script` carries inherited bullets instead of declaring them lost.** A paragraph
  read as having no bullet child of its own now maps to `bullet: 'inherit'` rather than to an
  absent option, and `text.bullet.inherited` is retired — it was the **largest** fidelity note
  on the corpus at 34/44 fixtures, along with `layout.text.bullet.inherited` at 5/44. Across
  the 44-fixture corpus the standalone tier's note count falls from 1018 to 713 and the
  template-anchored tier's from 712 to 419. The distinction always survived the read leg —
  `Paragraph.bulletDetail` is `null` for a paragraph with no bullet child and `{ kind: 'none' }`
  for one stating `a:buNone` — so this was a missing write option rather than an unreadable
  construct, which is worth noting because the note filed it as `unread`. Consumers diffing
  printed scripts will see `bullet: 'inherit'` appear on most paragraphs.

- **`ts-pptx/script`'s standalone tier rebuilds a layout's decoration instead of dropping
  it.** A source layout became a `defineSlideMaster` call carrying a title and a
  background, and nothing else: the bands, rules, wordmarks, triangles and quote marks
  that make a deck recognisable as somebody's template were declared lost as
  `master.decoration` and thrown away, so a standalone script produced a deck that
  rendered its slides faithfully and wore a different suit. Each of a layout's
  non-placeholder shapes is now transcribed into that layout's
  `defineSlideMaster({ objects })` array, by **the same mapper the slides go through** —
  so a rectangle on a layout is decided by the code that decides one on a slide, and the
  two cannot drift. Fills, gradients, scheme tokens, custom geometry, adjust handles,
  rotation, effects and text all come across on the shape mapper's existing terms.

  The note claiming this was unwritable was wrong about which half was holding it, and
  measuring it is what showed that: `{ shape: { type } }` accepts any preset, `custGeom`
  included, and its `ShapeProps` carries fill, line, shadow and adjust handles, so most
  of the supposed write-side ceiling was already reachable. Three kinds genuinely are
  not, and each is handled rather than assumed away. A **group** has no variant, so it is
  flattened into its children — visually lossless, because `absoluteFrame` composes the
  group's offset, rotation, flips *and* child-space scaling into each child, and noted as
  `layout.group` because one selectable object becomes several. A **connector** has no
  variant either and is re-authored as a `line` preset: it paints the identical stroke,
  keeps its rotation (which the slide-side `addConnector` mapping loses), and matters
  more than it sounds — PowerPoint's line tool authors a `p:cxnSp`, so connectors are 18
  of the 45 shapes the fixture corpus's layouts actually draw. A **table** is dropped,
  with a `layout.decoration` note naming the shape.

  Layout **placeholders** are still not emitted, and that is unchanged and deliberate:
  the write path seeds every slide with each layout placeholder it did not populate, so
  re-declaring one would put an empty ghost shape on every slide bound to the layout.
  `master.decoration` survives for a **master's** own shapes and moves from a reading
  problem to a structural one — `defineSlideMaster` creates a *layout* under the single
  shared master, so a master's shape tree has no write-side counterpart to receive them.

  Fidelity notes recorded against a layout shape are namespaced under a new
  `LAYOUT_NOTE_PREFIX` (`layout.`), exported from `ts-pptx/script`, because the shared
  mapper speaks the slide vocabulary: `layout.line.width` is `line.width` seen from the
  chrome. The prefix is load-bearing twice over — the template-anchored tier suppresses
  every note under it (it rebuilds no layout, so none of them describes its output), and
  the round-trip check refuses to let one excuse a difference on a *slide*, which a
  shape name repeated between a layout and the slides bound to it would otherwise do.
  `isKnownNoteConstruct` is exported alongside it and resolves a prefixed construct to
  its slide entry in the coverage table.

- **`SlideMaster.shapes` / `SlideLayout.shapes` reach a template's own content, and
  `showMasterSp` says whether to draw it** (#12). Both classes exposed `placeholders`
  and nothing else, so the `p:sp`/`p:pic`/`p:graphicFrame` under a master's or layout's
  `p:cSld/p:spTree` — the bands, rules, logos and footer furniture a deck is recognized
  by — had no modeled path out of the read API at all; only `part.dom` reached them,
  which is the raw-XML hatch rather than the model. On a corpus of PowerPoint-authored
  decks this was the single largest read gap: a consumer walking `slide.shapes`
  reproduces the deck's content and none of its identity. Both getters return the same
  `AnyShape` union `Slide.shapes` does, from the same `buildShapes` dispatch, so
  shape-walking code applies unchanged — and the members carry the full paint surface
  (`resolvedFill`, `resolvedLine`, `presetGeometry`, `rotation`, `absoluteFrame`), which
  the smaller `Placeholder` class never had. Tokens resolve against the *owning* part's
  context: a master shape's `schemeClr accent2` goes through the master's own `p:clrMap`
  and theme, not a slide's. Groups recurse and compose to slide-absolute frames as at
  slide level, and both classes get `shapeByIdDeep`.

  `placeholders` is unchanged, and is now documented as the filtered view of the same
  tree — both hand out the same live `p:sp` elements. Read a placeholder there to
  *place* it, through `shapes` to *draw* it.

  Shipping the accessor alone would have traded one wrong answer for another, so
  `Slide.showMasterSp` and `SlideLayout.showMasterSp` ship with it: `@showMasterSp`
  (ECMA-376 attributeGroup `AG_ChildSlide`, `xsd:boolean` defaulting to `true`, so absent
  means shown) is how a slide or a layout suppresses the master's decorative shapes —
  PowerPoint writes it on section dividers and full-bleed layouts. Without it a consumer
  that gained access to master shapes would paint them onto slides that deliberately hid
  them. Both are read-only; the write API authors neither. The layout arm has a genuine
  oracle in `mixed.pptx` and `read-stress.pptx`, which each carry `showMasterSp="0"` on
  their title layout.

- **`fill: { type: 'inherit' }` authors a shape whose interior comes from the style
  reference or the placeholder** (#10). 3.1.0 gave `type: 'none'` its `<a:noFill/>` back
  (#9) and, in doing so, left the *inherit* state with no spelling on the shape and text-box
  path: that emitter defaults a **missing** `fill` to `<a:noFill/>`, so omitting the option
  is an explicit transparent interior, not silence. The two arms were therefore both
  no-fill, and 3.1.0's own migration advice — "omit `fill` instead" — pointed at the wrong
  one. `'inherit'` emits no `EG_FillProperties` member at all, which is the state that lets
  `p:style/a:fillRef` or the placeholder paint the shape. The default is untouched: every
  existing `addShape`/`addText` call without a `fill` keeps emitting `<a:noFill/>`, which is
  why this is additive rather than a fix to the ternary. `ShapeLineProps` inherits the new
  member, where it means the same thing the stroke side already got for free — omit the
  paint child, keep the `<a:ln>`. On a table cell and a slide background, where *omitting*
  the option already meant inherit, `'inherit'` is simply the explicit spelling of that.

- **Any commit is installable straight from GitHub: `npm i github:shbernal/ts-pptx#<sha>`.**
  It looked like this already worked, and it never did. `dist/` is gitignored and `prepare`
  only installed git hooks, so a git-URL install packed a tarball in which every `exports`
  entry named a file that had never been built — the install succeeded and the first import
  failed. `prepare` now also runs `scripts/ensure-dist.mjs --if-missing`, a new mode that
  builds an *absent* `dist/` and leaves a stale one alone. That distinction is the whole
  design: the existing freshness check would be wrong in `prepare`, where it would make
  `pnpm run build` and `pnpm run watch` build twice over, and every script that needs a
  current build already front-loads its own unconditional `ensure-dist`. Absent means
  build; stale is somebody else's question. Two things had to move with it, both only
  reachable once consumers could run this path at all: `ensure-dist` invokes the build
  through `npm_execpath` — the package manager actually running it — rather than a
  hardcoded `pnpm`, which is declared here as `packageManager` but never installed as a
  dependency and so resolved to a shim a plain-npm consumer does not have; and
  `install-hooks.mjs` now skips when `INIT_CWD` places the caller outside the checkout,
  which is what a consumer's install looks like. Without that it ran `lefthook install`
  inside the package manager's throwaway clone and propagated lefthook's exit status into
  someone else's `npm install`. This is for trying an unreleased fix, not for production
  dependencies: the install builds from source, so it pulls this package's
  `devDependencies` and takes minutes.

### Changed

- **BREAKING (write): `TextPropsOptions.strike` admits `'noStrike'`, and `underline.style`
  matches `ST_TextUnderlineType`** (#14). `strike` was `boolean | 'sngStrike' | 'dblStrike'`,
  which had no spelling for the explicit off even though the serializer passes any truthy
  string straight to the attribute — so `strike: 'noStrike' as 'sngStrike'` already produced
  the right XML and only the published type refused it. It is now
  `boolean | 'noStrike' | 'sngStrike' | 'dblStrike'`. The breaking half is `underline.style`:
  `'dotDashHeave'` was a typo for `'dotDashHeavy'` and is corrected, and the missing
  `'words'` is added, so the union is the enumeration's full 18 members. Migration: replace
  `underline: { style: 'dotDashHeave' }` with `'dotDashHeavy'` — the old spelling was not a
  legal `ST_TextUnderlineType` value, so any deck written with it carried an invalid `@u`.

  The doc comments now also say which state silence is an alias for, on all three
  decorations. `false` and an omitted `strike` both write no attribute and therefore state
  *nothing*, leaving the run with whatever it inherits — the same as `bold`/`italic`, whose
  falsy arm is likewise an omission. `'noStrike'`, `underline: { style: 'none' }` and
  `caps: 'none'` are the spellings that state "off" and override an inherited decoration.

- **BREAKING (read): `Shape.slide` is now `Shape.host`, typed `ShapeHost`.** A shape proxy
  no longer belongs to a slide specifically — the same `p:sp` can sit in a slide's,
  a layout's, or a master's `p:spTree` — so the back-reference names what it actually
  is. `ShapeHost` is the small contract all three classes satisfy (`part`, `partName`,
  `opc`, `relationships`, `themeContext()`, `shapeByIdDeep()`) and is exported from
  `ts-pptx/read`. Migration: `shape.slide` → `shape.host`; where you genuinely need the
  `Slide`, narrow with `shape.host instanceof Slide`. `Slide` gains an `opc` getter
  (`=== presentation.opc`) and `SlideMaster`/`SlideLayout` gain public `opc` and
  `relationships` getters, all to satisfy that contract. Nothing else about a shape
  changed, and `Slide.shapes` is untouched.

- **The `ts-pptx-upstream` skill covers the far end of the cycle, not just the filing.**
  It ended at "write the workaround", which is the half that happens on its own — the
  half that rots is the release landing and nobody finding the stopgaps it retired. A
  new step 7 is that sweep: `npm view` for what is out, closed issues, `rg 'ts-pptx#'`
  for every marked stopgap here, then per stopgap bump, run the check the comment names,
  delete, leave a test behind, and comment upstream with the check that now passes. It
  rests on the marker actually saying something, so step 6 no longer prescribes
  `remove once fixed upstream` — a line that cannot be checked at bump time and sends
  the reader back to re-derive the reproduction — but the wrong output as an observable,
  the exact check that proves the fix, and the code the stopgap becomes. `ts-pptx#` is
  named as the literal token to grep for, which is what makes step 7 a command rather
  than a memory. Two other gaps a downstream consumer found by working through it: step
  4 now checks `npm view @shbernal/ts-pptx version` *before* the tracker, since a report
  against a stale pin costs a maintainer a full triage and ends in a close; and step 1
  carries the `FidelityNote.cause` triage — `unread` and `unwritable` are gaps worth
  filing, `unsupported` is the format's own limit and is not — which is our own
  machine-readable verdict on whose bug a silent loss is, and was being re-derived in
  each consumer's notes instead of read off the note.

- **The README installs the skill the way an agent will actually run it.** The documented
  line was the interactive one, in a section addressed to agents: it prompts for the
  skill and the runtimes, which unattended is a hang rather than an install. The
  non-interactive form is now the default shown, with the interactive one as the aside,
  and `--all` is called out — it installs into every runtime the CLI knows, around
  seventy, and leaves an `agent/` directory at the consumer's repository root. Also
  written down: the installed copy is a copy, so a version bump does not update it, and
  `skills experimental_install` restores the file from `skills-lock.json` without
  creating any runtime link, so it is a record rather than a restore command.

- **`docs/RELEASING.md` closes the loop the skill now depends on.** Issues here close
  when the fix merges, which is right for this repo and the wrong signal for a consumer:
  merged-and-unreleased lasts weeks, and a workaround deleted on the strength of a closed
  issue breaks against the version actually installed. Post-publish now comments the
  carrying version on each issue the release closed. `AGENTS.md` gains the rule #10 was:
  when a fix gives an option value a meaning it did not have, work out what its *absence*
  now means, because an emitter that defaults through a ternary has already assigned one
  state to silence.

### Fixed

- **`readModelToIr` carries a text decoration's explicit off** (#14). A run stating
  `a:rPr/@u="none"`, `@strike="noStrike"` or `@cap="none"` was read correctly by
  `Run.underline` / `Run.strike` / `Run.caps` and then mapped to `undefined`, so the
  resulting `CallIr` carried no option and re-emitting wrote no attribute. All three now
  carry — `underline: { style: 'none' }`, `strike: 'noStrike'`, `caps: 'none'` — and only an
  *absent* attribute maps to an absent option.

  `none` / `noStrike` are not the same fact as stating nothing. Each is a member of its own
  enumeration (ECMA-376 §20.1.10.81, §20.1.10.78, `ST_TextCapsType`) and would be redundant
  with omission if omission were the only way to be off. It is not, because run properties
  resolve down the `a:lstStyle` → placeholder → layout → master chain: a run that would take
  `u="sng"` from its list style and states `u="none"` is not underlined, and the same run
  with the attribute dropped is. So the loss was invisible on a deck with no inherited
  decoration and a visibly wrong answer on one that has any.

  It was undeclared either way — the same shape as #13. `DeckIr.fidelity` named neither
  construct, and `canonicalDeckIr` did not carry the field, so `diffDeckIr` compared two
  models that were *both* missing it and reported the deck clean; a consumer's round-trip
  harness structurally could not see this. Carrying the tokens closes both at once, and
  neither state notes. Two PowerPoint-authored fixtures state these tokens: `mixed.pptx` and
  `table.pptx` hold 132 runs stating `u="none"` and `strike="noStrike"`, 100 of which also
  state `cap="none"`. No new note fires, so `script:census` is unmoved.

- **`readModelToIr` keeps a baked `normAutofit`'s `fontScale` and `lnSpcReduction`** (#13).
  The mapper flattened every `normAutofit` frame to `fit: 'shrink'`, so both numbers were
  gone by the time a consumer held the `DeckIr` and everything downstream re-emitted a bare
  `<a:normAutofit/>`. Neither end of the round trip was at fault — `TextFitShrinkProps`
  already carries both fields and `TextFrame.autofitFontScale` already reads them back — so
  this was a mapping omission with an exact fix available: a frame that bakes either
  attribute now emits the object form `fit: { type: 'shrink', fontScale, lnSpcReduction }`,
  and one with neither keeps `fit: 'shrink'`.

  The two spellings are two states, not one value at two precisions. ECMA-376 §21.1.2.1.3
  defaults each attribute to 100%/0% only when it is *omitted*; PowerPoint recomputes an
  unbaked scale on edit and draws a baked one exactly as written until then. A deck baked at
  `fontScale="40000"` therefore came back painting its text two and a half times too large,
  in a file valid either way, which is why nothing caught it.

  It was also silent, which is the worse half. `printScript` named no fidelity note, so a
  consumer following the documented rule — trust the tier's own notes — was told the tier
  lost nothing; and `canonicalDeckIr` did not carry the fields either, so `diffDeckIr`
  compared two models that were both missing them and reported the deck clean. A round-trip
  oracle built on `diffDeckIr`, which is what the docs recommend building, could not detect
  this class of loss at all. Carrying the numbers through the IR closes all three at once.

  One arm still loses something, and now declares it: the write path rejects a percentage
  outside 0–100 and drops the attribute with a warning, so a source outside that range falls
  back to bare `'shrink'` with a `text.autofit.fontScale` / `text.autofit.lnSpcReduction`
  note (`dropped`/`unwritable`) rather than passing through a number that would vanish. No
  corpus fixture is malformed, so both read 0/44 and the census is unmoved.

- **Chart area and plot area fills honour `type` instead of only `color`** (#11).
  `ChartPropsFillLine.fill` is typed `ShapeFillProps`, so both areas looked like they took
  every fill kind a shape does — and `c:spPr` really is `a:CT_ShapeProperties`, the same
  optional `EG_FillProperties` group, so nothing in OOXML said otherwise. But both emitters
  gated on `fill?.color`, so every spelling that carries no colour fell to the `<a:noFill/>`
  arm and did nothing: `type: 'gradient'`, `type: 'pattern'`, and `type: 'inherit'` the
  moment it was added above. They now go through the shared fill dispatch, so a chart area
  can take a gradient or a pattern, `'none'` states a transparent area explicitly, and
  `'inherit'` emits no fill child at all and leaves the area to the chart style.

  Two spellings deliberately still mean no-fill, because the gate is on a fill being
  *stated* rather than merely present. `normalizeChartOptions` defaults `plotArea.fill` to
  `{}`, so every chart ever authored arrives at the emitter carrying a fill object — a
  presence check would have painted all of them a default grey. And `{ transparency: 50 }`
  with no colour is not a fill: there is nothing for the alpha to apply to. That was a
  documented `@example` on the option and has never worked; it now reads
  `{ color: '696969', transparency: 50 }`.

  `type: 'image'` remains unavailable on a chart and now warns
  (`image-fill/unresolved-media`) instead of silently doing nothing: a blip fill needs a
  media relationship on the chart part, and only shape and slide-level fills register one.
  No existing deck changes — the only inputs whose output moved are ones that used to emit
  `<a:noFill/>`, an invisible area nobody asked for.

## [3.1.0] - 2026-08-09

### Added

- **The package ships a `ts-pptx-upstream` skill, and `InternalError` says where to send a
  report.** Most code that calls this library is written by an agent working in a repository
  that is not this one, and an agent that hits a library defect writes a workaround instead of
  filing — which is the rational move from where it stands, since the workaround unblocks its
  user today and the tracker belongs to a repo it is not in. So the report never happens and the
  next consumer rediscovers the same defect. Three layers answer three different reasons the
  report dies. `InternalError` now appends the tracker URL from its constructor rather than from
  its throw sites, so a site added later cannot forget it; it is the only class that does, because
  it is the only one that already declares whose bug it is, and a "report this" banner on every
  malformed package would train callers to skip the line that always means something. The
  `ErrorCode` TSDoc — and therefore the `.d.ts` an agent in `node_modules` actually reads — now
  states the test for the other four: the supported bar is *"the output opens cleanly in Microsoft
  PowerPoint"*, and it reads in both directions, so a `PackageReadError` on a file PowerPoint opens
  cleanly is our gap, not bad input. And `skills/ts-pptx-upstream/` is published in the tarball, so
  `npx skills add ./node_modules/@shbernal/ts-pptx` works offline and always matches the installed
  version. The skill's load-bearing instruction is the one about the deck: presentations carry
  client names, unreleased strategy and pricing, and the tracker is public, so it spends most of
  its length on reducing a failure to a script that builds its own deck — and passes
  `--repo shbernal/ts-pptx` on every `gh` call, since `gh` in a consumer repo would otherwise file
  our bug into theirs. A third issue form, `agent-report.yml`, is where the error message's URL
  lands; its attachment dropdown deliberately has no option for a file containing real data, and
  "what should have happened" asks for the reason — an ECMA-376 clause, PowerPoint's own behaviour,
  or the docs — since that is what separates an actionable report from a matter of opinion. No API
  changed. `InternalError.message` gained a trailing pointer, which is not a contract: the class and
  the `code` are API and the message never was.

- **`TableCell.fillNoFill` reads an explicit `<a:noFill/>` on a cell** (#7) — the cell-side
  counterpart of 3.0.0's `Shape.fillNoFill`, and what `TableCell.noFill()` has always been
  able to write. `hasOwnFill` is not this question: it is `true` for any
  `EG_FillProperties` child, so it cannot separate a suppressed fill from a gradient, a
  pattern, a picture or an `a:grpFill`, and every colour accessor (`resolvedFill`,
  `fillColor`, `fillSchemeColor`) reports `null` for a no-fill cell exactly as it does for
  one inheriting the style's shading. Deriving it as "has an own fill that no accessor
  recognises" also changes meaning silently the day a further fill kind gets an accessor.

### Changed

- **Development toolchain: ESLint + Prettier → oxlint + oxfmt, and TypeScript 6 → 7.**
  This is a contributor-facing change with **no runtime effect on the published package** —
  no API was added, removed or altered, and the generated OOXML is byte-identical across
  all 183 parts of the reference decks. Enforcement was held level rather than relaxed: of
  the 90 lint rules previously enabled on `src/`, 89 carry over (the missing one, `no-octal`,
  is already a syntax error in an ES module), type-aware linting stays on, and
  `no-floating-promises` / `no-misused-promises` now also cover `scripts/` and `test/`,
  where the old configuration could not afford them. See `docs/development.md`.

- **Published `.d.ts` files are textually different, though no type changed.** TypeScript 7
  prints declarations slightly differently: string literal types use single quotes
  (`readonly shapeType: 'autoShape'` rather than `"autoShape"`) and redundant parentheses
  are dropped (`fontRef?: (StyleFontRef | null) | undefined` becomes
  `fontRef?: StyleFontRef | null | undefined`). Chunk filename hashes shift as a
  consequence. No declaration, signature or export moved, and `publint`,
  `@arethetypeswrong/cli` and the packed-package smoke test all pass — but a consumer
  diffing shipped declarations between versions will see churn, which is why it is recorded
  here rather than left as an implementation detail.

  Note that documentation generation deliberately stays on TypeScript 6: TypeScript 7 is
  the native Go compiler and ships no JavaScript compiler API for TypeDoc to import. That
  copy is confined to the private `tools/api-docs` workspace package and reaches nothing
  that is published.

### Fixed

- **`fill: { type: 'none' }` emits `<a:noFill/>` instead of nothing at all** (#9). The
  option's own name states the shape is transparent, and it was the one call that did not
  produce that state: `genXmlColorSelection` had no `none` case, so the fill child was
  omitted entirely and the interior fell back to `p:style/a:fillRef` or the placeholder —
  a shape carrying a style reference rendered in the theme's accent colour rather than
  transparent. `line: { type: 'none' }` on the same options object has always emitted its
  `a:noFill`; the two now agree. This reaches every caller of the shared fill dispatch, so
  a table cell (`TableCellProps.fill`) can author a transparent cell the same way.
  Round-trip consequence: `addShape({ fill: { type: 'none' } })` → save → load now reports
  `Shape.fillNoFill === true`, where it reported `false` before. The one behaviour that
  changes for existing decks is that `type: 'none'` no longer produces the *inherit* state
  by accident — if you were relying on it to mean "leave the fill to the style", use
  `type: 'inherit'`, added in the next release. (This sentence originally said to omit the
  `fill` option, which is wrong on the shape and text-box path: a missing `fill` emits
  `<a:noFill/>` there. See #10.)

- **`readModelToIr` carries a line's `@cap`, and declares a dropped `@algn`** (#8). Both
  legs of the cap mapping already existed — `ShapeLineProps.cap` authors the attribute and
  3.0.0's `Shape.lineCap` reads it back — but the script tier's `lineOption` never consumed
  it, so a deck this library wrote could not survive its own converter. The loss was
  *undeclared*, which is the part that mattered: the round-trip gate excludes exactly what
  a fidelity note names, so it passed green while the deck changed. `@cap` extends every
  dash by the stroke width and decides whether each draws as a rectangle or a lozenge, so
  on a thick dashed rule the before and after are visibly different. `@algn` is readable and
  has no write option, so it now records a `line.align` note (`dropped`/`unwritable`) —
  for `algn="in"` only, since `ctr` is what an omitted `@algn` already renders as.

- **`latexToOmml` emits accents as `<m:acc>` rather than `<m:limUpp>`** (#6). `\hat`,
  `\bar`, `\vec`, `\dot`, `\ddot`, `\tilde`, `\acute`, `\grave`, `\check`, `\breve`,
  `\mathring`, `\H`, `\dddot` and their short-form aliases (`\^`, `` \` ``, `\'`, `\"`,
  `\.`, `\=`, `\u`, `\v`, `\r`) all landed as over-*limits*, with limit spacing and
  semantics, because temml emits a bare `<mover>` (correct for a browser — MathML renderers
  derive accent positioning from the operator dictionary) and mathml2omml has no dictionary
  and keys strictly off `accent="true"`. The pipeline now carries the small dictionary
  subset that closes the gap, and while it is there it swaps temml's *spacing* modifier for
  the combining mark ECMA-376 §22.1.2.20 says an `accPr` character should be — so `\vec{v}`
  gets an arrow accent instead of a full-size arrow hung over the base.

  Scoped to `latexToOmml`: `mathmlToOmml` passes hand-written MathML through unchanged,
  because there `accent` is the caller's to set. Constructs that were already mapping well
  are untouched (`\widehat`/`\overbrace` stay `m:groupChr`, `\overline`/`\underline` stay
  `m:borderBox`, `\stackrel` stays `m:limUpp`), and two stay limits by necessity: `\utilde`
  and other under-accents, since OMML has no under-accent object and the symmetric
  `accentunder="true"` makes mathml2omml emit an *over*-accent; and `\ddddot`, whose
  two-character operator has no single `m:chr`.

- **A table cell's explicit `a:noFill` survives read → script → write.** Previously a
  suppressed cell fell out of `cellFill` as "no fill option" and the copy took the table
  style's banding — the opposite of what the source showed. Enabled by the
  `TableCell.fillNoFill` reader above and the `fill: { type: 'none' }` writer fix above.

## [3.0.0] - 2026-08-09

### Added

- **`Shape.fillNoFill` reads an explicit `<a:noFill/>` on a shape's fill** — the fill-side
  counterpart of `lineNoFill`, and the only accessor that separates a deliberately
  transparent shape from one that inherits its fill through `p:style/a:fillRef`. Every other
  fill accessor (`fillColor`, `fillSchemeColor`, `resolvedFill`, `gradientFill`,
  `patternFill`, `pictureFill`) reports `null` for both, so a consumer honouring the read
  model painted an `a:noFill` rectangle in the theme's accent colour. The same class had
  `ChartFill.noFill`, `ChartLine.noFill` and `CellBorder.noFill` already; the shape fill was
  the omission. `Shape.noFill()` has always been able to *write* this state.

- **`Shape.lineCap` and `Shape.lineAlign` read `a:ln/@cap` and `a:ln/@algn`.** `@cap` was a
  write/read asymmetry inside this library: `ShapeLineProps.cap` authors it and nothing read
  it back, so a deck this writer produced could not round-trip through its own reader without
  losing an attribute the writer put there on purpose. Both report the raw OOXML token, the
  way `lineDash` reports `@val` — `'flat'` / `'rnd'` / `'sq'` and `'ctr'` / `'in'` — and
  `null` when unset rather than a defaulted value. On a thick dashed rule the cap decides
  whether each dash reads as a rectangle or a lozenge and changes the drawn length of every
  one; SVG's `stroke-linecap` is the exact equivalent.

- **`GroupShape.childFrame` reads a group's own child coordinate space**
  (`p:grpSpPr/a:xfrm/a:chOff` and `a:chExt`), as a `ChildFrame` of `offsetX` / `offsetY` /
  `extentX` / `extentY` in EMU, or `null` when the group has no transform. `absoluteFrame`
  reads these internally to compose slide-absolute geometry and remains the right answer for
  anything that *paints* — this is for a consumer that *rebuilds* a group as OOXML, which
  needs the source child space to reproduce its scaling and could otherwise only rebuild
  groups whose child space is the identity. Named after the OOXML attributes rather than the
  read model's `left`/`top`/`width`/`height`, because it is the source rectangle of a
  mapping, not a frame on the slide.

- **`clipPath()` names the clip silhouettes you would otherwise re-derive.** A `ClipShape`
  is data — a named silhouette plus its options — and `clipPath(shape, w, h)` resolves it to
  the freeform `points` path `addImage` emits as its `<a:custGeom>` clip mask. The first
  silhouette is the half-disc a cover-slide picture placeholder cuts. See
  [`docs/image-in-shape.md`](docs/image-in-shape.md).

  ```js
  import { clipPath } from '@shbernal/ts-pptx'

  const w = 5.22, h = 7.5
  slide.addImage({
    path: 'cover.jpg', x: 0, y: 0, w, h,
    points: clipPath({ kind: 'half-disc', flat: 'right' }, w, h),
    sizing: { type: 'cover', w, h },
  })
  ```

  `flat` names the edge the straight side sits on; `preset` picks the proportion, `'deep'`
  (the default) or `'shallow'`. Paired with `sizing: 'cover'` this is a standalone
  reproduction of what a picture placeholder does — a layout `custGeom` clipping an
  inherited blipFill — with no placeholder, and no layout, involved.

  **The box size is an argument for a reason.** A `custGeom` point written as `%` resolves
  against the *slide*, not the picture, so a box-relative silhouette has to be emitted in
  inches already scaled to its box. `clipPath` multiplies its fractions out at build time,
  which is what lets one silhouette scale to any region — and is why handing a path to a
  picture of a different size puts the clip somewhere else entirely. That trap is the whole
  reason this is worth shipping rather than leaving to each caller.

- **`slide.addModel3d()` embeds a 3D model** — PowerPoint's *Insert ▸ 3D Models*. A glTF
  binary (`.glb`) travels inside the package, and PowerPoint 2019+ renders it live and lets
  the viewer orbit it. See [`docs/3d-models.md`](docs/3d-models.md).

  ```js
  slide.addModel3d({
    path: 'assets/engine.glb',
    preview: { path: 'assets/engine-render.png' },
    meterPerModelUnit: 1 / 240, // the model's largest bounding-box dimension
    x: 1, y: 1, w: 6, h: 4,
  })
  ```

  Two things are worth knowing before using it:

  - **Supply `preview`.** Everything that is not PowerPoint 2019+ draws that picture instead
    of the model — including PowerPoint's own slide thumbnails, PDF export, and print. The
    library has no 3D renderer, so omitting it embeds a gray placeholder and emits a
    `model3d/preview-missing` warning. Same bargain as `addOleObject()`'s `cover`.
  - **Set `meterPerModelUnit`.** The `am3d` scene is measured in metres. PowerPoint reads the
    model's bounding box and normalizes its largest dimension to 1 metre; ts-pptx does not
    parse glTF, so it emits `0.5` (correct for a model 2 units across) and leaves the rest to
    you. Left at the default, a model 240 units across becomes a 120-metre object with the
    camera inside it. Set it to `1 / <largest bounding-box dimension>`.

  `camera` overrides the viewpoint (`pos`/`lookAt`/`up` in metres, `fov` in degrees); the
  defaults are the ones PowerPoint wrote for a 2×2×2 cube. Out-of-range and non-finite values
  throw rather than being coerced. Linked (non-embedded) models, animation scenes, and a typed
  read accessor are out of scope for now — a model read through `ts-pptx/read` surfaces as an
  inert `graphicFrame` and survives load → save and `importSlide` byte-intact.

- **Browser support is now proven in CI, not assumed.** A Playwright lane
  (`pnpm run test:browser`, CI job `browser`) runs the package in headless Chromium
  against `demos/vite-demo`. No library code changed — this converts an existing
  claim into evidence.

  The assertion worth naming is cross-runtime byte identity. The demo imports the same
  showcase module `pnpm demos:build quarterly-review` builds, and `src/zip.ts` pins
  `FIXED_MTIME`, so the two packages are directly comparable: all **113 parts** of the
  browser-built deck are byte-identical to the Node-built one. Every serializer, the zip
  writer, part ordering and relationship numbering are therefore runtime-invariant — by
  comparison, not by inspection. A second spec reads the object-URL download back with
  jszip (an implementation independent of the `fflate` the library writes with) to confirm
  it is a real OPC package.

  Two boundaries stay exactly where they were, and are now stated in
  [`docs/runtime-and-package-support.md`](docs/runtime-and-package-support.md) rather than
  left to inference:

  - **Runtime support is not layout fidelity.** Nothing in the lane depends on a rendered
    page. Real `offsetWidth`, the resolved cascade, and browser-chosen fonts remain out of
    active scope; a layout difference between two browsers is not a defect in this
    package's browser support, whereas a `.pptx` a browser builds differently from Node is.
  - **`tableToSlides` measurement is still unavailable without a layout engine**, exactly as
    documented — `offsetWidth` is `0`, widths fall back to computed CSS and then an equal
    split, and `data-pptx-width` / `data-pptx-min-width` pin them.

  The explode/normalize/diff machinery the byte-identity gate has always used moved to
  `scripts/pptx-parts.mjs` so both gates share one definition of "the same bytes"; the
  refactor was verified byte-identical against a baseline frozen with the pre-refactor
  script.

- **The whole `RuntimeAdapter` now runs in a real browser**, not just the download path.
  A second Playwright fixture serves the shipped `dist/browser.js` **unbundled** over a
  static server and drives decks written to reach the three loaders `demos/vite-demo`
  cannot, because its showcase draws every asset rather than loading one:

  - `loadMedia` — a fetched raster image lands in the package as the same bytes Node reads
    off disk, *and* as the same bytes as the source file. The two implementations return
    different strings for the same image (Node raw base64, the browser a `FileReader` data
    URI) and everything downstream reconciles them, including the image sizer.
  - `createSvgPngPreview` — the `<canvas>` rasterizer, whose branches nothing exercised
    before: a real PNG where Node can only stub a placeholder, plus the undecodable-SVG and
    zero-dimension arms, each of which must fail rather than ship a blank fallback.
  - `loadFontData` — a font fetched over HTTP bakes the same `fontScale` and embeds the
    same `/ppt/fonts/` bytes as one read off disk.

  The deck definitions are written once and built twice, once per runtime, so a divergence
  in the fixture cannot read as a divergence in the runtime.

  Two things fell out of loading the shipped file unbundled. `opentype.js` turns out to be
  a *dynamic* bare import inside the measure/fit chunk — invisible to every bundled
  consumer, and now documented for anyone loading the entry over a plain
  `<script type="module">`. And Node and the browser are *expected* to disagree on exactly
  one part: the SVG PNG fallback, where Node has no rasterizer. The lane asserts the shape
  of that disagreement so it cannot quietly become a different one.

- **`tableToSlides` is tested against a table a browser actually laid out** — a third
  Playwright project, `html-table`, on the same unbundled harness server.

  `pickColWidthBasis` chooses between three column-width bases, and its *first* arm — the
  rendered `offsetWidth` — had never executed anywhere. The Node suite drives happy-dom,
  where `offsetWidth` is `0` for every cell, so it always took a fallback arm; the unit
  suite reached the function by handing it numbers directly, which proves the `if` and not
  the pipeline behind it. So the primary path of the feature, including the `arrColSrc`
  arithmetic that fixed the spanning-`data-pptx-width` defect, was covered only at its own
  function boundary.

  The fixture is built so the measured basis and the computed-CSS basis **disagree**
  (`offsetWidth` is the border box, computed `width` the content box), because a test that
  only showed "the widths came out proportional" would be equally green if the measured arm
  never ran. The spec re-derives both bases from the live page and fails if they ever
  converge. Sensitivity-checked by disabling the arm: exactly the two arm-dependent
  assertions go red, reporting the CSS ratio.

  **This does not move the scope line, and the wording matters.** What is asserted is that a
  measurement is *taken and honoured* — proportionally, with `data-pptx-width` still winning
  outright. Nothing asserts that Chromium's numbers are the right numbers or that another
  engine would agree; that is live-DOM layout fidelity, it has no oracle, and it stays out
  of active scope. A layout difference between two browsers is still not a defect in this
  package; a `.pptx` a browser builds differently from Node still is — which the lane pins
  directly, by converting the same markup in both runtimes and asserting Node falls back to
  the CSS basis where the browser measures.

  The new project contributes its V8 coverage to the merge like every other browser spec, so
  the merged report moved up on all four axes — statements 93.91 → 94.03, branches
  83.71 → 84.16, functions 98.29 → 98.58, lines 96.11 → 96.20. Two notches in
  `scripts/coverage-gates.json` are ratcheted with it (statements 92 → 93, branches
  82 → 83), which is what keeps a gate from carrying two points of slack.

- **Coverage from both lanes is merged into one number** (`pnpm run coverage:gate`, CI job
  `coverage`). `scripts/coverage-merge.mjs` folds the browser lane's V8 coverage into the
  Node report on one rule — *the Node report defines the shape, the browser lane
  contributes hits* — so the merged denominator is identical to the Node report's and the
  two percentages are directly comparable. Merged: statements 93.91, branches 83.71,
  functions 98.29, lines 96.11.

  It also makes this repo's **point-of-slack rule fail a build** rather than live in prose:
  `scripts/coverage-gate.mjs` is red both when a number falls below its notch *and* when it
  clears it by less than a full point. Prose does not fail a build, which is how the
  exclusion drop left `functions` at 0.35 of slack while an acceptance criterion of
  "thresholds still pass" was satisfied.

  The four numbers in `vitest.config.ts` remain the Node suite's own floor, and the browser
  lane keeps its per-function gate — a percentage cannot say *which* adapter function
  stopped running.

- **The package is now bundled for Node and run, inside `pnpm run test:package`.** The
  browser lane put a real bundler in front of the `browser` condition; nothing asked the
  same of the `node` entry, and the two are different questions — Node's resolver finds a
  specifier on disk at call time, while a bundler must resolve every one of them, including
  dynamic ones, at build time. A package can be perfectly importable and still be
  unbundlable.

  `bundleForNode()` esbuild-bundles the installed tarball with `platform: 'node'` across
  every export subpath but `/browser`, then runs what it emitted. It fails if the build
  warns, if anything other than a Node builtin stayed external, or if the emitted bundle
  cannot write a `.pptx`. Both the npm and pnpm fixtures, since pnpm's symlinked store is a
  different shape for a bundler to walk.

- **A bundle-size budget for the browser entry** (`pnpm run bundle-size:check`, part of
  `check:package`). Nothing measured shipped size before; a size promise nobody measures is
  a promise that quietly stops being true. It fails only on a step change and asks for a
  re-freeze only when a win is worth banking, because bytes move on every commit. The
  figure is a growth detector, not a download size — `dist/` is unminified and every real
  browser consumer minifies it.

### Changed

- **BREAKING: `Paragraph.bullet` is replaced by `Paragraph.bulletDetail`.** The old accessor
  reported a *tagged string* — `'none'` / `'char:•'` / `'autoNum:arabicPeriod'` — which is
  ambiguous when the glyph is itself a colon and reads as a bare glyph if you do not know
  better. That is not hypothetical: this library's own script converter first consumed it as
  one, and `'none'.codePointAt(0)` put a literal `n` bullet on every converted deck, silently.

  `bulletDetail` returns a discriminated union with no parsing left to get wrong, and carries
  what the string could not — `a:buAutoNum/@startAt`, and the bullet's own `a:buFont` /
  `a:buSzPct` / `a:buSzPts` / `a:buClr`. It also reports a fourth kind the old accessor
  dropped to `null`: a picture bullet (`a:buBlip`), with its image part resolved.

  ```js
  // before
  para.bullet // 'autoNum:arabicPeriod' — startAt, font, size and colour unreachable

  // after
  para.bulletDetail
  // { kind: 'autoNum', scheme: 'arabicPeriod', startAt: 5,
  //   font: 'Wingdings', sizePct: 80, sizePt: null,
  //   color: 'C00000', schemeColor: null, resolvedColor: { … } }
  ```

  Migration: `bullet === 'none'` → `bulletDetail?.kind === 'none'`;
  `bullet?.startsWith('char:')` → `bulletDetail?.kind === 'char'`, with the glyph at
  `.char` rather than after the colon; `bullet.slice('autoNum:'.length)` → `.scheme`.
  A paragraph that inherits its bullet still reports `null`.

  Numbering is content rather than styling: a list continuing "5. Deploy" that came back as
  "1. Deploy" was a different slide, and `numberStartAt` was a pure write/read asymmetry —
  `addText` accepted it and nothing could produce it.

- **BREAKING (output): `<a:buSzPct/>` is emitted only when `bullet.size` is given.** It used
  to be written unconditionally, pinned to `val="100000"`, on every object-form bullet and on
  `bullet: true`. An explicit 100% is not the same as leaving it out — it *overrides*
  whatever bullet size the layout's or master's list style sets, so every bullet this path
  wrote silently forced its glyph back to full size. The same class of bug as the explicit
  `a:buNone` an omitted `bullet` emits, and invisible until `bulletDetail` gave the
  round-trip check something to see it with. An out-of-range `bullet.size` now warns and
  emits nothing rather than warning and pinning to 100%. Decks that want the old behaviour
  can pass `bullet: { size: 100 }` explicitly.

- **`sizing.w` / `sizing.h` are optional, and `sizing: { type: 'stretch' }` names what used to
  be nameless.** The emitter has always defaulted the fit box to the picture's own extent; the
  type demanded both anyway, so every `cover`/`contain` call restated `w`/`h` it had already
  supplied a line above. `sizing: { type: 'cover' }` is now the ordinary form, and passing them
  still means what it always did — a fit box deliberately different from the picture.

  `stretch` emits the plain `<a:stretch><a:fillRect/></a:stretch>` a raster already gets. It
  exists so the fill-the-box behaviour can be *asked for*, which is what makes the vector
  default below opt-out-able rather than a trap.

- **BREAKING: an SVG is placed at its own aspect ratio by default instead of being stretched
  to its box.** `addImage({ svg, w, h })` used to fill the box whatever the glyph's proportions
  were, so any icon with a non-square `viewBox` — a minority in every real icon set, and never
  the one you check — came out squashed. Every consumer's answer was the same wrapper that
  routes each call through `sizing: 'contain'`; that wrapper is now the library's default.

  ```js
  slide.addImage({ svg: icon, x: 1, y: 1, w: 3, h: 1 })                            // letterboxed, centered
  slide.addImage({ svg: band, x: 0, y: 0, w: 13.33, h: 0.4, sizing: { type: 'stretch' } }) // opt out
  ```

  **Scope, deliberately narrow.** Rasters are untouched: a photo's box is chosen for it, filling
  it is what PowerPoint does, and letterboxing every existing deck's pictures would be a change
  of a different order. A vector is different in kind — it *states* its ratio in a `viewBox`,
  and disagreeing with that statement is a defect rather than a layout choice. Nothing is
  emitted when the ratios already agree, so a square glyph in a square box produces the same
  bytes it always did, and an SVG carrying neither `viewBox` nor `width`/`height` stretches
  silently — no sizing was requested, so there is nothing to warn about.

  **Two related SVG fixes fall out of the same root cause** — the write path treated vector
  sources as unmeasurable long after `src/media/image-size.ts` learned to read a `viewBox`:

  - `{ svg, w: 4 }` derives its height from the intrinsic ratio (a 2:1 viewBox → 4in × 2in),
    where before an omitted dimension silently became 1 inch. Rasters have always done this.
  - `{ svg }` with **neither** dimension still falls back to 1 inch, and that is not an
    oversight. An SVG's user units are dependable relative to each other and merely
    conventional in absolute terms; treating them as 96-DPI pixels would insert a 24-unit icon
    as a quarter-inch object. The ratio is trusted, the magnitude is not.

  If you have a wrapper that adds `sizing: 'contain'` to every icon, delete it — the emitted
  XML is identical either way. If you were relying on a stretched vector, name it: `sizing: {
  type: 'stretch' }`.

- **`vitest.config.ts` no longer excludes anything of this repo's own from coverage.**
  The `dist/browser.js` / `dist/browser-*.js` entries are gone; the second never matched
  anything, because tsdown bundles the adapter *into* the entry. Dropping them took the
  measured functions figure 98.33 → 97.35 while actual tested-ness went up, which is the
  shape of an honest denominator: the Node suite cannot execute an adapter that needs
  `fetch`, `FileReader` and a canvas.

- **`@shbernal/ts-pptx/math` is Node-only by decision, not by accident.** It loads its
  optional peers through `createRequire`, which is what keeps `latexToOmml()` and
  `mathmlToOmml()` synchronous; the browser-compatible alternative is a dynamic `import()`
  that would make both async — a breaking change to every existing caller, for a use case
  nobody has raised. If a browser consumer turns up, the answer is an additional
  `/math/async` subpath, not a change to this one. Recorded in
  [`docs/runtime-and-package-support.md`](docs/runtime-and-package-support.md) so it is not
  re-litigated per release.

- **The docs called `tableToSlides`'s no-browser width path a *degradation*. It is a
  *fallback*, and the difference is not cosmetic.** Every user-facing statement of it — the
  README, `docs/project-target.md`, `docs/runtime-and-package-support.md`, `AGENTS.md`, the
  `/html` entry's own doc comment — said column widths "degrade to computed CSS", which
  reads as *the same answer, less precisely*. They do not measure the same box:
  `offsetWidth` is the border box and a computed `width` is the content box, so padding
  alone can put the two bases in different **proportions**. The `html-table` fixture is
  built to demonstrate exactly that — 1:1 measured against 2:1 from CSS on one table — so
  the repo had the fact and the docs contradicted it. Now stated wherever the fallback is
  described, with the remedy: pin the column with `data-pptx-width` where both runtimes
  have to agree.

  No behaviour changed, and deliberately so. Normalizing the CSS basis to the border box
  would need computed padding and border widths, which the DOMs that reach that arm need
  not resolve (a `%` padding computes to nothing usable without layout) — so it would
  converge the two only sometimes, and it would collapse the one discriminator the browser
  lane has, turning `table-widths.spec.mjs` back into a test that passes whether or not the
  measured arm ran. The reasoning is recorded on `pickColWidthBasis` itself.

- **The browser lane stays Chromium-only, deliberately.** The APIs in play (`fetch`,
  `FileReader`, canvas, object URLs, `<a download>`) are not where engines are known to
  disagree, and no divergence has been reported or observed. An engine gets added when
  there is something concrete to add it for — also written down rather than left as a
  default.

### Fixed

- **`readModelToIr` now emits its `table.rowAuto` note when *every* row is auto-height**, not
  only when some are. The guard excluded the all-auto case explicitly, and that case is both
  the more common one — a table authored with no explicit row heights has `a:tr/@h="0"` on
  every row — and a real loss: no `rowH` is emitted at all, so `addTable` divides the frame
  height evenly and three auto rows come back pinned to a third of the frame each. The table
  still looks the same; what is lost is the *implicitness*, which matters the moment someone
  edits a cell and expects the row to grow. A round-trip oracle gated on "nothing undeclared"
  is only as good as its note set, and this was a difference passing through undeclared.

- **`autoPage` let every continuation slide take one row more than fitted, so the last row
  hung off the bottom edge.** Affects `addTable(rows, { autoPage: true })` and
  `tableToSlides()` alike, and only tables whose cells carry top/bottom margins — which is
  every table converted from HTML, since cell padding becomes a cell margin.

  The pager charges each row its cells' top and bottom margins before deciding whether the
  row fits. On a page break it did that and then zeroed the accumulator, so the first row of
  each new page — and only that row — was placed for free. The page then filled to its budget
  as if it had that space, and it did not. The deeper the padding, the further the overflow:
  at 8px of cell padding a 60-row table paged `[10, 11, 11, 11, 11, 7]` where every full page
  had room for 10.

  The symptom is easy to miss because the error is *constant*: every generated page overflows
  by the same amount, so the pages agree with each other and disagree only with the first
  one. That is also why the existing continuation-slide regression stayed green — it compared
  continuation pages to each other.

  This closes upstream `gitbrent/PptxGenJS#1200`, filed against `tableToSlides` and long
  assumed to be a browser-layout question. It is not: the browser supplies column widths, and
  nothing the vertical arithmetic reads. `test/browser/table-autopage.spec.mjs` pages the
  same table in headless Chromium and on a DOM that renders nothing and asserts the two
  produce the same slides.

## [2.0.0] - 2026-08-05

### Removed

- **BREAKING: `Presentation.defineTableStyle()` and `TableProps.styleDrivenCells` are
  gone, along with the `TableStyleProps` / `TableStyleRegionProps` types.** A custom table
  style is unreachable markup in PowerPoint: it emitted well-formed, schema-valid XML that
  can never paint.

  PowerPoint resolves `<a:tableStyleId>` against its **own** table-style gallery and never
  reads a style definition out of the package. Measured by rendering in PowerPoint desktop
  16.0, not inferred from the schema:

  - a deck pointed at a **built-in** GUID that the package does not define renders
    correctly — so the gallery, not the part, is what is consulted;
  - a PowerPoint-authored deck with one style's GUID rewritten to a novel value in *both*
    `ppt/tableStyles.xml` and the slide, bytes otherwise identical, loses that table's
    styling entirely (black hairline grid on white) while its untouched neighbours keep
    theirs;
  - the same holds for a definition placed inline in `<a:tblPr>` as `<a:tableStyle>`, and
    for one nominated by the part's `def=`. Lifting a genuine PowerPoint-authored
    `<a:tblStyle>` block under a custom GUID does not help either, so the markup we emitted
    was never the problem.

  `styleDrivenCells` was actively harmful under that finding: it stood down the per-cell
  `border` and `color` defaults so a style region could take over, but no custom region ever
  paints, so it traded a correct grid for PowerPoint's default one.

  **Migration.** Style tables with direct formatting, which is what carried them all along:

  ```js
  // before
  const brand = pptx.defineTableStyle({
    name: 'Brand',
    wholeTbl: { border: { type: 'solid', color: 'D9D9D9', width: 0.5 } },
    firstRow: { fill: '1A2B3C', color: 'FFFFFF', bold: true },
  })
  slide.addTable(rows, { tableStyle: brand, hasHeader: true, styleDrivenCells: true })

  // after
  slide.addTable(rows, {
    hasHeader: true,
    border: { type: 'solid', color: 'D9D9D9', width: 0.5 },
    headerRow: { fill: { color: '1A2B3C' }, color: 'FFFFFF', bold: true },
  })
  ```

  Banded rows move into the row data (set each row's `fill` as you build it) or into
  `columns[i]` for vertical banding. `TableProps.tableStyle` still exists and still works —
  it is now typed as `TableStyle` alone, since only a built-in GUID renders.

  `ppt/tableStyles.xml` is still emitted, as a bare stub naming a default style id, because
  PowerPoint expects the relationship and content-type override. The diagnostics
  `table-style/region-overridden`, `table-style/missing-argument`, `table-style/missing-name`
  and `table/style-driven-cells-inert` are removed with the API.

  **The read side is unaffected.** `Table.resolvedStyle`, `TableCell.resolvedFill` and
  `importSlideMasters({ tableStyles })` still resolve style graphs out of imported decks —
  those definitions were written by PowerPoint, and PowerPoint honours its own.

### Fixed

- **`tableToSlides` cell padding is now converted from px to inches.** A cell's
  computed CSS `padding-*` was read in px and assigned straight to
  `TableCellProps.margin`, which is **inches**. The magnitude therefore passed
  through unscaled: a perfectly ordinary 4px padded cell emitted
  `marL="3657600"` — a **4 inch** text inset, wider than most columns — and any
  cell padded 1px or more also tripped the `margin/legacy-points` warning. The
  stale `px->pt 1:1` note it was written under had been true when cell margins
  were points; it was not true after margins became inches.

  Padding now resolves at 96px/in, so `padding: 4px` is `4/96in` (`marL="38100"`).
  96 rather than 72 because CSS defines the reference pixel as 1/96in and this
  conversion exists to mirror what the browser laid out; it is also the density
  the `"<n>px"` coordinate unit already uses, so the two px sites agree. The
  whole-px rounding is gone with it — a fractional computed padding keeps its
  precision, and the rounding happens once, in EMU.

  **This changes emitted bytes for any padded HTML table**, on the browser path
  as well as `@shbernal/ts-pptx/html`. Cells will look substantially tighter,
  because they are now inset by what was asked for. Nothing else about the
  conversion changed, and a table whose cells set no padding is unaffected.

- **`fitColumns: 'shrink'` measured against the wrong slide margin.** The space to the
  right of a table is bounded by the **right** margin; the calculation subtracted
  index 3 of the TRBL margin tuple, which is the **left** one. Invisible with the
  default symmetric `[0.5, 0.5, 0.5, 0.5]` — and wrong by the difference for any master
  whose `margin` is asymmetric, which is exactly the layout someone sets a left gutter
  on. The existing tests all used symmetric margins, so none of them could fail; the
  new one uses `[0.5, 0.25, 0.5, 2]`, where the two answers differ by 1.75in.

- **A styled table cell's own fill is no longer dropped by `pptx-to-script`.** A cell
  with an explicit `a:solidFill` inside a table that also has a `tableStyle` replicated
  as *unfilled*, because the mapper could not tell that colour apart from one the cell
  merely inherited from the style's header/banding rules — and baking an inherited
  colour in would freeze the banding, so it dropped both.

  `TableCell.hasOwnFill` (new, read side) tells them apart: an `a:tcPr` either carries
  an `EG_FillProperties` child or it does not. `TableCell.resolvedFill` already
  branched on exactly that internally; the flag simply was not exposed. A cell's own
  fill is now carried and an inherited one is still left to the style GUID, which
  reproduces the banding exactly. Neither case loses anything, so the `table.cell.fill`
  fidelity note is gone rather than merely narrowed.

  Measured on PowerPoint's own `table-cell-image-fill.pptx`: the red cell keeps its red.

### Added

- **Table editing on `ts-pptx/read`.** The read proxies were read-plus-text-edit only:
  `TableCell.text` was the sole setter, and every other change needed the `element_`
  escape hatch plus a manual `markDirty()`.

  Cell properties: `setAnchor`, `setVerticalText`, `setHorzOverflow`, `setAnchorCtr`,
  `setMarginsEmu`, `setBorder(edge, …)` (the four edges and both diagonals),
  `setFillColor`, `setFillSchemeColor`, and `noFill()`. Structure on `Table`:
  `addRow`, `removeRow`, `addColumn`, `removeColumn`, `mergeCells`, `unmergeCell`.
  Each mutates in place and marks the part dirty, matching the `text` setter.

  Every insertion respects the `CT_TableCellProperties` **sequence**. That is the
  whole hazard: an append-only setter produces an out-of-order `a:tcPr`, which
  PowerPoint reports as a corrupt file rather than as a bad edit, and which no getter
  would notice. A schema-validation case now authors a deck, edits it through these
  setters, saves it, and validates the result — the only shape of test that catches it.

  Structural edits keep the grid rectangular and every merge's continuations in step.
  Inserting a row or column *through* a merge extends it rather than splitting it;
  removing a merge origin promotes its first continuation, so the region survives one
  row shorter; removing a column inside a merge drops a covered cell rather than the
  origin, so the region keeps its content. `mergeCells` **rejects** a rectangle that
  cuts through an existing merge instead of silently widening it to fit.

  Unlike the write path, an invalid value here **throws** rather than warning and
  dropping — new codes `table/invalid-cell-anchor`, `table/invalid-cell-vert`,
  `table/invalid-cell-overflow`, `table/invalid-cell-margin`,
  `table/invalid-cell-border`, `table/row-index-out-of-range`,
  `table/column-index-out-of-range`, `table/merge-range-invalid`. On the write path a
  bad option comes from a deck being built, and dropping one value beats failing the
  build; here it comes from a caller editing one attribute, and doing nothing silently
  would leave them looking at an unchanged deck with nothing to explain it.

- **`TableProps.tableFill`** — the table's own background, written as a real `a:tblPr`
  fill that the cells sit on top of. The existing `TableProps.fill` is *stamped onto
  every cell* instead, so nothing ever reached `a:tblPr`. The two usually render alike,
  which is why the difference is worth stating: with `fill` there is no such thing as
  an unfilled cell, so a cell can never fall back to a table background — and a deck
  read back from PowerPoint carries the `a:tblPr` shape, not the flattened one.
  `fill` is unchanged (changing it would repaint every existing deck); both JSDocs now
  say which is which.

  Takes the same `ShapeFillProps` a cell does — solid, gradient, pattern or picture.
  Read back via `Table.resolvedFill`, `.pictureFill`, `.gradientFill`, `.patternFill`
  and `.fillSchemeColor`, the same five a cell has.

  No table-level **effect** surface: PowerPoint's UI exposes none, so a source deck
  will not contain one and there would be nothing to reproduce.

- **`TableCell.gradientFill` / `TableCell.patternFill`** (read) — a cell's `a:gradFill`
  and `a:pattFill`. `TableCell.resolvedFill` reports `null` for every non-solid choice
  by design, and it was the only fill accessor a cell had besides `pictureFill`, so a
  gradient- or pattern-filled cell was indistinguishable from an unfilled one.

  That was not only a reading gap. `pptx-to-script` had nothing to fall back on, so a
  gradient cell **replicated as an unfilled cell** — or, when the table had a style, as
  whatever banding colour the style graph resolved to. Both now round-trip. Writing
  them always worked; it is now documented and pinned by tests, and the output is
  confirmed to open in desktop PowerPoint.

- **`TableCellProps.diagonal`** — a cell's corner-to-corner rules
  (`a:tcPr/a:lnTlToBr` / `a:lnBlToTr`), PowerPoint's "Diagonal Down/Up Border". The
  read model has always decoded them; there was no way to write one, and
  `pptx-to-script` dropped them with a note. Kept off `border`'s tuple deliberately:
  widening that to six entries would break every existing caller for a rare feature,
  and the diagonals are not edges. A merged region draws its diagonal **once**, on the
  span origin — covered cells inherit the origin's edges but never its diagonals,
  because a diagonal is one corner-to-corner stroke and repeating it per covered cell
  would draw a sawtooth.

- **`TableCellProps.anchorCtr`** — centres a cell's whole text *block* horizontally
  (`a:tcPr/@anchorCtr`), independent of each paragraph's `align`. The two compose:
  `align` places each line inside the text block, `anchorCtr` places that block inside
  the cell. Read back via `TableCell.anchorCtr`. `false` is the schema default and
  emits nothing.

- **`TableCellProps.cell3D`** — a 3-D bevel on a cell (`a:tcPr/a:cell3D`): preset,
  width/height in points, `prstMaterial`, and an optional light rig. Niche —
  PowerPoint's table UI has no control for it, so it reaches a deck from a theme or
  another producer — but it round-trips through PowerPoint verbatim, so authoring and
  replicating one now works. Read back via `TableCell.cell3D`.

  Two schema constraints show through the API. `a:bevel` is required, so `cell3D: {}`
  still emits a bevel rather than an empty (invalid) `a:cell3D`. `a:lightRig` requires
  **both** `rig` and `dir`, so a half-specified rig is reported and dropped whole. Any
  value outside its `ST_` union is reported as the new **`table/invalid-cell3d`** and
  dropped. The four enums are exported as `BevelPresetType`, `PresetMaterialType`,
  `LightRigType` and `LightRigDirection`.

- **`TableCell.id` / `TableCell.headerIds`** (read only) — a cell's `a:tc/@id` and the
  header cells associated with it (`a:tcPr/a:headers/a:header/@val`), which is how a
  complex table tells a screen reader what a value means.

  **There is deliberately no write-API counterpart.** PowerPoint opens a deck carrying
  both without complaint and then strips them on the first save, so an emitter would
  ship a feature that dies as soon as anyone edits the deck. The measurement is
  `test/read/fixtures/authoring/probe-table-cell-a11y-and-3d.ps1`, and it is a
  controlled one: `a:cell3D` and `a:headers` were injected into the *same* `a:tcPr`,
  and PowerPoint kept the first and discarded the second. `TableProps.hasHeader`
  (`a:tblPr/@firstRow`) remains the header marker PowerPoint keeps — and the one its
  own accessibility checker reads. The accessors exist because a deck from another
  producer may still carry the association; `pptx-to-script` records its loss as
  `table.cell.headers`.

- **`BorderProps.dashType`** — the exact `a:prstDash` preset for a border, using the
  same vocabulary as `ShapeLineProps.dashType`. `BorderProps.type` is only a coarse
  three-way switch, so every dashed border it can express — dotted, long-dash,
  dash-dot — collapsed onto the single `sysDash` preset on write *and* on read-back.
  A deck whose table borders are `dot` or `lgDashDot` could not be authored or
  replicated. `dashType` names the preset directly and wins over `type` when both are
  set; `type: 'none'` still suppresses the border before any dash is chosen. Honored
  by table cell borders.

  `ShapeLineProps.dashType` (and therefore `BorderProps.dashType` and
  `ConnectorProps.dashType`) now spans the **whole** `ST_PresetLineDashVal` set: the
  three values it was missing — `dot`, `sysDashDot`, `sysDashDotDot` — are accepted.
  This is a widening, so no existing value changes meaning. `pptx-to-script` maps a
  read dash straight through instead of flattening it, and a dash outside the enum is
  recorded as `table.cell.borders.dash` rather than silently approximated.

  An unrecognized `dashType` is reported as the new **`border/invalid-dash-type`**
  diagnostic and falls back to what `type` implies, rather than being written — a
  value outside `ST_PresetLineDashVal` would make the part schema-invalid, which
  PowerPoint reports as a corrupt file.

- **`TableProps.outerBorder`** — a border for the table's **perimeter** only: the top
  edge of the first row, the bottom edge of the last row, the left edge of the first
  column and the right edge of the last column. A single `BorderProps` boxes the
  table; a TRBL array with holes rules only the sides it names, leaving the others to
  whatever `border` already drew. Applied after every other border source, so
  "outline the table, no interior grid" is `outerBorder` with no `border` at all.

  The perimeter is decided by **grid position**, not by authored cell, so merges
  work: PowerPoint defines a merged region's outer edges on the *covered* cells, and
  a colspan reaching the last column gets that column's rule on its `hMerge` dummy.
  Leaving the option unset emits nothing — existing decks are byte-identical.

- **Vertical table cell text now survives `pptx-to-script`.** `TableCellProps`
  inherits `textDirection` from `TextBaseProps` and the emitter has always written it
  to `a:tcPr/@vert`, but the replication mapper recorded the attribute as unwritable
  and dropped it, so a replicated deck lost every vertical cell label for no reason.
  The four directions the option spells (`horz`/`vert`/`vert270`/`wordArtVert`) now
  round-trip; the `table.cell.vert` note is narrowed to the East-Asian
  `ST_TextVerticalType` modes that genuinely have no spelling.

- **`TableCellProps.horzOverflow`** (`'clip' | 'overflow'`) — controls what a table
  cell does with a **single glyph** wider than its text width, emitted as
  `a:tcPr/@horzOverflow`. `'clip'` (PowerPoint's default) cuts the glyph at the cell
  edge; `'overflow'` lets it draw past. It matters for oversized display type, wide
  CJK/emoji glyphs, and icon fonts in a narrow column. Read back via the new
  `TableCell.horzOverflow` accessor on `ts-pptx/read`, and carried through
  `pptx-to-script`.

  **It is not a text-wrap switch, despite where it sits.** That distinction is the
  reason this landed: the attribute had long been filed as the route to per-cell
  no-wrap, and it is not. PowerPoint has no per-cell no-wrap at all — `wrap="none"`
  on a cell's `a:bodyPr` renders inert and is stripped on the next save, and
  `TextFrame.WordWrap` is read-only on a cell over COM. Cell text always wraps to
  the column width. `test/read/fixtures/authoring/probe-table-cell-wrap.ps1`
  reproduces every part of that, and `table-cell-horzoverflow.pptx` is the
  PowerPoint-authored oracle for what does work.

  Writing `'clip'` explicitly is honored but redundant: it is the schema default, so
  PowerPoint drops the attribute the first time it saves the deck. Leaving the option
  unset emits nothing, so no existing deck's bytes change. An unrecognized value is
  reported as **`table/invalid-horz-overflow`** and dropped rather than written —
  `ST_TextHorzOverflowType` admits only those two, and PowerPoint reports a
  schema-invalid slide part as a corrupt file rather than as a mis-set option.

- **`border/unknown-key`** — a new diagnostic reporting a key that is not part of
  `BorderProps`, which was previously discarded without a sound. The thickness
  field is `width`, in points; a border authored with any other name for it lost
  the value and rendered at the 1pt default, and a `.pptx` gives no second signal
  — nothing throws and the deck opens, only heavier than asked for.

  TypeScript's excess-property check already rejects a stray key on a border
  written inline at the call site. It deliberately does not fire when the border
  is built as a variable first (`const b = {...}` then `{ border: b }`), because
  a variable may legitimately be a supertype — and that is precisely the reuse
  pattern a shared grid style encourages. The check closes that gap at runtime.

  It sits in `resolveBorderWidth`, the one function every emitted border resolves
  its width through, so it covers table cell borders, table-style regions, and
  chart borders alike. One exception: `chartArea.border` is rebuilt from a fixed
  key list during normalization, so an unknown key on it (and `cap` with it) is
  stripped before generation and cannot be reported. Emitted bytes are unchanged
  in every case — this only adds a report where a value was already being lost.

### Changed

- **`docs/tables.md`** — the table guide, which did not exist. Tables were the
  most-used object after text and the only major one with no prose doc. Covers the
  cell model, the styling precedence chain, borders (per-cell default vs. perimeter,
  dash styles, diagonals), fills, merges, sizing and auto-paging, the read/edit
  surface, and a "not authorable" section for the constructs PowerPoint discards.

- **`TableProps.border` is documented as what it is: a per-cell default, not the
  table's perimeter.** The old wording ("single value applied to all 4 sides / array
  in TRBL for individual sides") read like the outside of the table. It never was:
  `normalizeTableRows` broadcasts it to *every* cell, so
  `border: [solid, none, solid, none]` gives each cell a top and bottom rule — a full
  set of horizontal grid lines — rather than a rule above and below the table.
  Nothing about the behaviour changed; the doc now says so plainly and points at the
  new `TableProps.outerBorder` for the perimeter case.

- **A table-level `border` side with `width: 0` is now emitted as the hairline it
  asks for**, instead of being replaced by the 1pt default. Per-*cell* borders
  already treated `0` as a real width; only the table-level path used a truthiness
  test, so the two disagreed for exactly one value. Both now share one helper. A
  border that sets no `width` at all is unaffected.

## [1.0.0] - 2026-07-29

Initial public release of ts-pptx — an ESM-first, TypeScript-first library for
generating PowerPoint `.pptx` files from Node.js and modern JavaScript
toolchains.

ts-pptx descends from [gitbrent/PptxGenJS](https://github.com/gitbrent/PptxGenJS)
(MIT) and has been developed independently since; see the README for lineage and
the [project target](docs/project-target.md) for scope. It ships its own API and
makes no backwards-compatibility guarantee with the original project.

### Added

- Slide authoring: slides, layouts, masters, sections, speaker notes, and
  presentation metadata.
- Content: text, tables, shapes, connectors, groups, images, SVGs, charts,
  media, and OLE objects.
- Outputs: file, stream, buffer, Blob, base64, and browser download, depending
  on the runtime.
- A `.pptx` read model for opening, inspecting, and round-tripping existing
  decks (`@shbernal/ts-pptx/read`).
- Standalone text measurement and table-fit helpers
  (`@shbernal/ts-pptx/measure`).
- Native equation authoring from LaTeX or MathML (`@shbernal/ts-pptx/math`).
- Explicit ESM package boundary with typed subpath exports for `inspect`,
  `measure`, `read`, `math`, `node`, and `browser`.

- **`Run.charSpacingPt`** on the read model — character spacing (tracking) in
  points from `a:rPr/@spc`, the read counterpart of the write-side `charSpacing`
  option. It sat beside `caps`, `strike`, and `baselinePct` as the one run
  property the flat inspect surface could read and the deep model could not.

- **An error taxonomy: `TsPptxError` and five subclasses.** The library threw
  ~160 bare `Error`s and shipped no error classes at all, so a consumer wanting to
  tell *"you passed a bad coordinate"* from *"this font file is corrupt"* from
  *"these bytes are not a package"* had to match on message substrings — text the
  project is explicitly free to reword in any release.

  Every failure is now a `TsPptxError` carrying a stable `code`:

  ```ts
  import { MediaError, PackageReadError, InvalidOptionError } from '@shbernal/ts-pptx'

  try {
  	await buildDeck(spec)
  } catch (err) {
  	if (err instanceof MediaError) return retryWithPlaceholderAsset(spec)
  	if (err instanceof PackageReadError) return rejectUpload(err.code)
  	if (err instanceof InvalidOptionError) throw err // our bug — fail loudly
  	throw err
  }
  ```

  The classes are a deliberately flat set of five — `InvalidOptionError`,
  `UnsupportedFeatureError`, `PackageReadError`, `MediaError`, `InternalError` —
  answering *whose problem is this?*; the `code` carries the specificity.

  **The class and the `code` are API; the `message` is not.** Codes draw on the
  same vocabulary as diagnostics, so a condition reads the same whichever way it
  surfaces: `coord/non-finite` means the same thing thrown as an
  `InvalidOptionError` or warned as a `Diagnostic`. Each code belongs to exactly
  one class and the pairing is type-enforced —
  `new MediaError('coord/non-finite', …)` does not compile.

  Everything remains `instanceof Error`, so existing `catch` blocks are unaffected.
  The classes are re-exported from every entry point and resolve to one shared
  module, so `instanceof` works regardless of which subpath you imported from and
  which subpath threw. Where the library wraps a lower-level failure, the original
  is preserved on `cause` rather than flattened into the message. See
  [docs/errors.md](docs/errors.md).

- **A diagnostics seam: `setDiagnosticHandler`.** Library warnings were hardwired
  to `console.warn` across ~100 call sites, so a consumer generating decks in a
  batch job could neither silence nor route them, and nothing downstream could
  react to a *specific* condition without matching on message substrings.

  Every warning is now a structured `Diagnostic { code, message, detail? }`
  delivered to a handler you can install:

  ```ts
  import { setDiagnosticHandler } from '@shbernal/ts-pptx'

  setDiagnosticHandler((d) => logger.warn({ code: d.code }, d.message))
  setDiagnosticHandler(() => {}) // silence
  setDiagnosticHandler(null) // restore the console default
  ```

  **The `code` is API; the `message` is not.** A code identifies a *condition* in
  `area/condition` form (`'chart/non-finite-value'`, `'coord/bare-number-is-inches'`)
  and is stable: adding one is back-compatible, removing or renaming one is
  breaking. The wording behind it is free to improve in any release — do not parse
  it. `DiagnosticCode` is a closed union, so codes complete in an editor and a
  typo is a compile error.

  There is no separate strict mode: a handler that throws is one, and it composes
  with whatever policy you want.

  ```ts
  setDiagnosticHandler((d) => {
  	if (d.code === 'coord/bare-number-is-inches') throw new Error(d.message)
  })
  ```

  The handler is process-global rather than per-presentation. That is deliberate
  and documented — the emitting code is a tree of free functions with no
  presentation in scope — with the trade-off (concurrent builds cannot be told
  apart) written down in [docs/diagnostics.md](docs/diagnostics.md).

### Changed

- **Breaking (non-Node, non-browser runtimes): the bare `@shbernal/ts-pptx`
  import no longer resolves to the browser build.** `exports["."]` carries
  `node`, `browser`, and `default` conditions. The first two were right; the
  third pointed at `dist/index.js`, which did nothing but re-export
  `./browser.js`. So a runtime that sets neither condition — Deno, Bun, an edge
  worker — got the DOM adapter, and `writeFile()` tried to create an anchor
  element and click it in an environment with no `document`. For a documented
  Node-first project, the neutral fallback should not have been the browser.

  `dist/index.js` is now its own runtime-agnostic entry. Authoring is unchanged
  and everything that hands bytes back to you still works — `write()`,
  `stream()`, `toParts()` — which is the shape a worker wants anyway. Remote
  media and fonts still load, over `fetch`/`btoa`/`TextEncoder`, which every one
  of these runtimes has. Two things differ:

  - **`writeFile()` throws** an `UnsupportedFeatureError`
    (`runtime/file-output-unavailable`) naming `@shbernal/ts-pptx/node` and
    `@shbernal/ts-pptx/browser`, rather than failing on a missing `document`
    deep inside the call. There is no filesystem and no DOM here, so there is no
    destination to write to; take the bytes from `write()` and place them
    yourself.
  - **`tableToSlides()` is absent.** It resolves an element id against the
    global `document`, so it is defined on the browser entry alone. The
    DOM-agnostic form is the free `tableToSlides` on `@shbernal/ts-pptx/html`,
    which takes the element directly.

  Nothing changes for Node or browser consumers at runtime: those conditions
  already resolved to `dist/node.js` and `dist/browser.js` and still do.

  **Types now resolve through the same condition as the code.** `.` previously
  served one `dist/index.d.ts` to every condition, so a Node consumer was typed
  against the browser class and TypeScript accepted `pptx.tableToSlides(…)` on a
  build where it was `undefined` at runtime. Each condition now carries its own
  `types`, so what the compiler shows matches what the runtime has.

- **`ts-pptx/inspect` is now a projection over `ts-pptx/read`, and
  `fast-xml-parser` is no longer a dependency.** The library shipped *two*
  independent readers of a `.pptx` over two different XML parsers: the deep,
  navigable `read` model on `@xmldom/xmldom`, and the flat `inspect` snapshot on
  `fast-xml-parser` with its own hand-rolled, JSZip-shaped package facade. The
  overlap was near-total — boxes, rotation, group composition, runs, paragraphs,
  font size, colour, fill, line, shape type, wrap, autofit, body insets, part
  listing — so every read-side fix had to be made twice, and a divergence between
  the two was invisible to every test in the repo.

  The *shape* of `inspect` is unchanged: a cheap, flat, allocation-light snapshot
  is a genuinely different use case from a navigable model, and every exported
  type and function keeps its name and meaning. The *implementation* is gone.
  `inspect` now reaches the package through `OpcPackage` and every field through
  the read model's own getters, so the two surfaces cannot disagree about what a
  deck says.

  **`fast-xml-parser` is dropped from `dependencies`** (it stays a devDependency
  for one maintenance script), which is 1.4 MB less installed for every consumer
  of the package, whichever entry they import.

  Four behaviour changes come with it:

  - **Slides are reported in presentation order (`p:sldIdLst`), not part-name
    order.** Dragging a slide in PowerPoint rewrites that list and leaves
    `slideN.xml` named as it was, so the old directory-order enumeration reported
    the authoring history rather than the deck. `slides[].index` is now the deck
    position; `slides[].path` still names the part.
  - **`textRuns[].text` is verbatim.** `fast-xml-parser` trims text nodes by
    default, so every run came back stripped of the whitespace an
    `xml:space="preserve"` run carries. That both lost the leading/trailing space
    that widens a line and welded adjacent runs together — a slide reading
    `"This is test content."` inspected as `"Thisis testcontent."`, with a word
    count to match. Element-level `text` (runs joined, whitespace collapsed,
    trimmed) is unaffected in shape but now has the right word boundaries.
  - **`loadPptxPackage()` returns an `OpcPackage`**, not the JSZip-shaped
    `{ files, file(path) }` facade; the `PptxPackage` / `PptxPackageFile` types are
    removed. `listPptxParts()` / `readPptxTextPart()` / `readPptxBinaryPart()` keep
    their signatures and still speak zip paths. Migration: a caller that reached
    into `pptxPackage.file(path).async('string')` calls `readPptxTextPart(pkg,
    path)`, and one that wants more can now use the `OpcPackage` directly or hand
    it to `Presentation.fromPackage()` without re-reading the bytes.
  - **The input must be a real OPC package.** A zip holding slide XML but no
    `[Content_Types].xml` used to inspect fine; it now throws a `PackageReadError`
    (`package/not-an-opc-package`), the same bar `ts-pptx/read` applies.

  A run highlight authored as a theme token now resolves to a literal hex against
  the slide's theme instead of reading `null`.

  The cost is bundle size: a consumer importing only `/inspect` now pulls the read
  model's chunks (~440 KB of library code before their own tree-shaking) where it
  used to pull ~20 KB plus `fast-xml-parser`. This project is Node-first, and one
  reader that is right beats two that drift.

  Every other field of every element across all 43 fixture decks is byte-identical
  to the old implementation, pinned by a new characterization snapshot
  (`test/read/fixtures/inspect-surface.snapshot.json`).

- **`Shape.presetGeometry` moved from `AutoShape` to the `Shape` base**, so a
  picture or connector reports its preset geometry too. PowerPoint gives both one
  — a `p:pic` is `rect` unless cropped to a shape — and `Shape.adjustValues` was
  already on the base documenting itself as the companion to a member only
  auto-shapes had.

- **Nine reality-checks that wrote to the console now throw or warn.** A handful of
  validation sites in `gen/define/` reported a problem with a direct
  `console.error` / `console.log` and then carried on. That output could not be
  captured, silenced, or branched on — it predated the diagnostics seam and was
  never a `warn()` call, so the migration to `setDiagnosticHandler` did not see
  it. Each site now takes whichever surface fits what the library actually does
  next:

  - **`addImage()` throws** `InvalidOptionError` when the source is unusable —
    `image/missing-source`, `image/path-not-a-string`, `image/data-not-a-string`,
    `image/missing-base64-header`. **This is a behaviour change:**
    `addImage({})` used to print a line and silently omit the image, leaving a
    deck that opened fine and was missing content. There is nothing to draw, so
    it now rejects, matching `addMedia()` and the `hyperlink` check in the same
    function.
  - **`addTable()` throws** `InvalidOptionError` (`table/rows-not-nested`) for a
    row that is not an array of cells. It already rejected exactly this for row
    0; later rows only logged and pushed an empty row, quietly dropping content.
    **This is a behaviour change** for the later-row case.
  - **A picture bullet whose `data` lacks a base64 header warns**
    (`bullet/image-missing-base64-header`) rather than throwing: the run emitter
    falls back to a default glyph, so the deck is still valid. Behaviour is
    unchanged apart from the output now being routable.
  - **A malformed `hyperlink` reports once instead of twice.** Registration
    logged a line and declined to mint a relationship; the emitter then threw
    `hyperlink/not-an-object` / `hyperlink/missing-target` for the same input.
    The log is gone — the throw was always the real report.

  The `verbose: true` table tracer still prints to the console. It is a DEV-ONLY
  flag whose output reports no condition, and not passing it silences it.

- **Error messages no longer label themselves.** Fourteen messages carried an
  `ERROR: ` / `ERROR! ` prefix, four an `addMedia() error: ` one, and
  `coordToEmu`'s carried a literal `ts-pptx: `. The class name already labels the
  failure in every stack trace and console rendering, so all of them are gone.
  `presentation.layout = 'nope'` threw the literal string `UNKNOWN-LAYOUT`; it now
  names the value and says what to pass instead. The conditions and their codes
  are unchanged — only code matching on message text is affected, which the
  contract has never supported.

- **Warning output is now the default handler's job, not the message's.** Two
  messages carried their own prefix (`[WARNING] `, `Warning: `) and one carried a
  literal `ts-pptx: ` inside the text, which the console handler then doubled.
  All three are gone; the prefix is applied in exactly one place. Only code that
  scrapes stderr for those exact strings is affected — the conditions, and now
  their codes, are unchanged.

- **`@shbernal/ts-pptx/html`: HTML `<table>` → slides, anywhere there is a DOM.**
  `tableToSlides` was reachable only as a method on the browser build, which made
  converting an existing HTML table a browser-only capability. It is now also a
  free function on a new `/html` subpath that runs under Node with any DOM
  implementation and in the browser, from one artifact — deliberately no
  `browser`/`node` condition split.

  ```ts
  import { tableToSlides } from '@shbernal/ts-pptx/html'

  tableToSlides(pptx, win.document.getElementById('report'))
  tableToSlides(pptx, 'report', { document: win.document }) // by id
  ```

  Pass the element and no global DOM is consulted at all — the document and its
  view come from the element's own `ownerDocument`/`defaultView`. Pass a string
  id and it resolves against the new `TableToSlidesProps.document`, defaulting to
  the global `document` as before. `pptx` is structural (`addSlide` +
  `presLayout`), so any presentation instance works, including the Node build;
  `masterTitle` resolves on both forms.

  Strictly additive: `TsPptx.prototype.tableToSlides(eleId, options)` keeps its
  exact signature and behavior, and its body is now a delegation to the shared
  implementation, so the two forms cannot drift.

  **Column widths degrade without a layout engine.** In a browser the columns are
  sized from each cell's rendered `offsetWidth`. Nothing outside a browser lays a
  table out, so `offsetWidth` is `0` there — which previously made the
  proportional calc a `0/0` divide and emitted a table with zero-width columns.
  The basis now falls back in two steps: the computed CSS `width`s when the
  stylesheet states them for every column in one unit (all `px` or all `%`), then
  an equal split. `data-pptx-width` / `data-pptx-min-width` on the `<thead>`
  cells still win outright on every path, and are the way to pin widths
  regardless of runtime.

  Cell text is read the same way: `innerText` where the DOM genuinely renders,
  and otherwise a `childNodes` walk that keeps `<br>` as a line break. jsdom does
  not implement `innerText` at all (every cell would have come out empty), and
  happy-dom implements it as `textContent` rather than rendered text (every cell
  would have come out on one line).

  New public types: `TableToSlidesElement`, `TableToSlidesDocument`,
  `TableToSlidesHost`. These are structural rather than `lib.dom`'s
  `Element`/`Document`, because a non-browser DOM's types do not satisfy those and
  demanding them would reject from TypeScript exactly the implementations the
  entry exists to accept.

- **`pptx-to-script` re-embeds a surface's picture fill.** An image-filled shape
  or table cell (`p:spPr/a:blipFill`, `a:tcPr/a:blipFill`) converted to a script
  and came back unfilled: the read model saw the blip, the write API could author
  one (`fill: { type: 'image', image: { data } }`), and only the converter's fill
  mappers were missing an `AssetResolver` to join them with. They have one now,
  so the bytes are carried as an asset exactly as an `addImage`'s are — deduped
  against every other reference to the same part — and the blip's
  `a:alphaModFix` opacity carries as `transparency`.

  The write path emits every picture fill as a plain stretched blip, so a tiled,
  cropped or inset fill is re-embedded and then *flattened*: that is the new
  `fill.picture.geometry` / `table.cell.fill.picture.geometry` note, recorded
  only when the source uses one of them. The older `fill.picture` /
  `table.cell.fill.picture` notes narrow to the surfaces whose bytes cannot be
  carried at all — a blip embedding no part, a part missing from the package, and
  an SVG, which `addImage` accepts but a fill does not.

- **The read model can now see a picture fill.** `TableCell.pictureFill` and
  `AutoShape.pictureFill` decode an `a:blipFill` on a *surface* — the cell's
  `a:tcPr`, the shape's `p:spPr` — into `{ relId, partName, mode, srcRect,
  fillRect, tile, alpha, dpi, rotWithShape }`. Rect edges are per-edge fractions
  (`÷ 100000`, so `0.1` is 10 %, and a negative `fillRect` edge stays negative);
  tile offsets stay in EMU and tile scales become fractions. All three call sites
  share one reader (`readPictureFill`), so a slide background's `image` variant
  now carries the same decoded `picture` alongside the `relId`/`partName` it
  always had.

  This closes the asymmetry the cell picture-fill writer opened: the library could
  author an image-filled cell and then read it back as unfilled, because
  `resolvedFill` decodes solid colours only. It also settles what a `Picture` is
  and is not — a `p:pic` whose image is its sibling `p:blipFill` — as against a
  shape or cell whose *surface* happens to be an image.

  Gated on `test/read/fixtures/table-cell-image-fill.pptx` for the cell half
  (stretched, tiled, bordered and merged picture cells, plus a solid control) and
  on `math-omml.pptx` for the shape half, whose `p:spPr/a:blipFill` carries a
  negative `<a:fillRect b="-6667"/>` bleed. That shape sits in an `mc:Fallback`
  branch, which the read model does not walk, so the test unwraps the
  `mc:AlternateContent` — the shape XML and its relationships are PowerPoint's own
  bytes either way.

### Fixed

- **`tableToSlides` cell padding was parsed with a regex that deleted the decimal
  point.** Computed `padding-*` went through `.replace(/\D/gi, '')`, which strips
  every non-digit — including the `.` — so a `1.5px` padding became the number
  `15`, a ten-fold inset, and `0.5px` became `5`. Fractional computed paddings are
  ordinary: any `em`/`%` padding, or a `rem` on a non-integer root size, resolves
  to one. It is parsed as a CSS length now, keeping the fraction and rounding
  once. A value that is not an absolute px length (a `%` padding, a keyword) insets
  by nothing rather than by whatever digits it contained.

  This is visible to existing browser callers: a table whose cells have
  fractionally-computed padding gets a correct inset where it previously got a
  roughly 10× one.

- **`tableToSlides` produced an empty table for an id that is not a valid CSS
  identifier.** The table id was interpolated raw into a CSS selector
  (`#${id} tr:first-child th`), so an id starting with a digit, or containing `.`
  or `:`, matched nothing — *after* passing the `getElementById` reality-check,
  which made it look like the table had simply been read as empty. Every query is
  now scoped to the element, so the id is never parsed as a selector. Selector
  semantics are otherwise unchanged.

- **`tableToSlides` emitted the literal color `NANNANNAN` for a computed color
  that was not `rgb()`.** Computed colors were parsed by stripping `rgb(`/`rgba(`
  and splitting on commas. A browser always normalizes to `rgb()`, so this held
  there — but nothing outside a browser normalizes, and a DOM that returns the
  authored `#ff0000` produced `Number('#ff0000')` → `NaN` → that string, emitted
  without complaint. Colors now go through a CSS parser that handles
  `rgb()`/`rgba()` and `#rgb`/`#rrggbb`, clamps and rounds channels, and falls
  back to the caller's default rather than guessing for anything else. Browser
  output is unchanged: `rgb()` parses exactly as before, and a fully transparent
  background still becomes white.

  A dead condition next to it was removed:
  `getComputedStyle(cell).getPropertyValue('transparent')` tested a CSS *keyword*
  as if it were a property, so it returned `''` for every cell and never fired.

- **`tableToSlides` dropped every row that was not inside a `<thead>`/`<tbody>`/
  `<tfoot>`, then threw `addTable: Array expected!`.** Rows were collected with
  three descendant queries (`thead tr`, `tbody tr`, `tfoot tr`), which see only
  sectioned rows. `<table><tr>…` is valid authored markup, and a table assembled
  with `createElement`/`appendChild` has no `<tbody>` at all — the HTML *parser*
  inserts one, the DOM API does not. Such a table lost all of its rows and
  reached `addTable` empty, which is the reported failure in upstream
  gitbrent/PptxGenJS#1005. Rows now come from the table's own row list, and a row
  in no section is treated as a body row. The same change stops a table nested
  inside a cell from having its rows folded into the outer table — the descendant
  queries matched those too (as does happy-dom's non-conformant `rows`, so each
  row's ownership is checked rather than assumed).

- **`tableToSlides` applied `data-pptx-width` to the wrong column when a header
  cell spanned.** The width overrides were looked up by *column* index
  (`thead tr:first-child th:nth-child(n)`) against a row indexed by *cell* — the
  two part ways the moment a `colspan` is involved. A 2-span header with
  `data-pptx-width="4"` followed by a header with `data-pptx-width="2"` sized the
  span's second column with the *next* header's 2in and left the last column with
  no override at all. This is upstream gitbrent/PptxGenJS#1244. A spanning cell's
  `data-pptx-width` / `data-pptx-min-width` now divides across the columns it
  covers, matching what its `offsetWidth` and its computed CSS width already did.

  Two related mismatches went with it. The overrides were read from `<thead>`
  `<th>` cells while the widths themselves came from the first row *anywhere*, so
  a table with no `<thead>` measured one row and took overrides from a row that
  did not exist — both now come from the same cells. And the width-source query
  was a descendant selector, so a table with `<th>` in both `<thead>` and
  `<tfoot>` derived twice as many columns as it had; `addTable` then rejected the
  column count and discarded `colW` wholesale, taking every override with it.

- **`tableToSlides` emitted ragged tables, and crashed on an empty one.** An HTML
  row states only the cells it starts, so a short row is ordinary markup; pptx has
  no such model — `<a:tblGrid>` declares a column count and a row carrying fewer
  `<a:tc>` is a table PowerPoint has to repair. Rows are now measured against the
  grid the table actually occupies (`colspan` widening a row, a `rowspan` from
  above filling one it never mentions) and padded with blank cells to the width of
  the widest row. A table with no cells at all no longer fails deep in the
  auto-pager with `Reduce of empty array with no initial value`; it throws a
  `tableToSlides:` error naming what is missing.

- **A line's `cap`, and a stroke's `pattern`/`image` paint, were dropped before
  reaching the emitter.** `ShapeLineProps extends ShapeFillProps`, so a stroke
  accepts `gradient`/`pattern`/`image` as well as a solid `color`, plus its own
  `cap` — and `drawingml/line.ts` reads all of them. But the define pass rebuilt
  the caller's `line` object from a fixed list of keys, so any key added to the
  type without also being added to that list never survived normalization.

  `cap` was silently ignored everywhere: `line: { cap: 'round' }` on a shape and
  `border: { cap: 'round' }` on a table both emitted `cap="flat"`. `pattern` was
  worse than ignored — `line: { type: 'pattern', pattern: {…} }` reached
  `genXmlPatternFill` with no pattern object and threw *"Pattern fill requires a
  pattern object."* A gradient stroke was dropped on the `addText({ shape:
  'line' })` path specifically, which carried its own near-duplicate rebuild.

  All four rebuilds (two for shape/text lines, two for table borders) now spread
  the caller's object and override only the keys they actually default, so they
  cannot fall out of sync with the type again. Output for any deck that did not
  set the dropped options is byte-identical. `addBackground()` / `defineSlideMaster({ background })` derived
  the media extension from `path` only. With no path it substituted the
  `preencoded.png` placeholder, so `background: { data:
  'data:image/svg+xml;base64,…' }` embedded SVG bytes in a `.png` part that
  `[Content_Types].xml` announced as `image/png` — the Default/payload mismatch
  PowerPoint offers to "repair" — and the same held for any non-PNG format. The
  `data:` mime is now read first and wins over `path`, as it always has for
  `addImage()`.

  The sniff itself moves to one shared `imageExtensionForSource(path, data)` in
  `src/media/content-type.ts`, replacing the four near-copies that had drifted
  apart (image objects, image fills, picture bullets, OLE preview covers) —
  which is how the background copy came to be the one missing the mime branch.
  Two small consequences of the single implementation: a mime is now lower-cased
  the way a path extension already was (`data:image/PNG;` names its part `.png`,
  not `.PNG`), and an OLE object's `cover` follows the same bytes-win precedence
  as everything else.

- **A combo chart's per-subchart options are validated like the chart-level
  ones.** `addChart` normalizes its options once, but a `ChartMulti` entry's own
  `options` were merged over them only at emit time — after every clamp and enum
  correction had run — so they reached the part verbatim. A subchart
  `barOverlapPct: 250` emitted `<c:overlap val="250"/>` where `ST_Overlap` is
  -100..100, `barGapWidthPct: 9999` blew past `ST_GapAmount`'s 500 and
  `barGrouping: 'sideways'` failed the `ST_Grouping` enumeration: three
  PowerPoint-repair prompts reachable only through the combo API, and silent —
  the same values passed at chart level have always been clamped with a warning.

  Each subchart's options now go through the same pass, keyed to that subchart's
  own plot type, covering `barDir`, `barGrouping`, `barGapWidthPct`,
  `barGapDepthPct`, `barOverlapPct`, `bar3DShape`, `holeSize`, `firstSliceAng`,
  `lineDataSymbol`, `lineDataSymbolSize`, `lineDataSymbolLineSize` and
  `dataLabelPosition`. What it validates is the merged value the emitter actually
  reads, writing back only what a correction changed, so per-subchart options stay
  a sparse override of the chart-level ones.

  That also closes the wider half of the same hole: a combo chart's internal
  `_type` is a `ChartMulti[]`, so the chart-level corrections that key off the
  chart *type* — `barGrouping`, `dataLabelPosition` — previously matched no branch
  and never ran either. They now resolve per subchart, which is why one bad
  chart-level `barGrouping` correctly lands as `clustered` for a bar group and
  `standard` for a line group.

  One behaviour change comes with it: a **stacked bar subchart now emits
  `gapWidth 50`**, the narrower default a chart-level stacked bar already got,
  where it previously inherited the clustered default of 150. An explicit
  `barGapWidthPct` on either the chart or the subchart still wins.

- **A table cell with a non-solid fill no longer reports the table style's
  colour.** `TableCell.resolvedFill` fell through to the style graph whenever the
  cell's own fill was not a solid one, so an image-, gradient- or pattern-filled
  cell under a shading style reported a colour PowerPoint never paints, and an
  explicit `a:noFill` cell reported the shading it was suppressing. A cell that
  declares any `EG_FillProperties` choice of its own now overrides the style, the
  same guard `AutoShape.resolvedFill` already applied to the theme style matrix.
  A cell with no fill choice at all still inherits its banding/header shading as
  before.

### Added

- **A table cell can now be filled with a picture.** `addTable` accepts an image
  fill on a cell exactly as a shape or text box already did — `fill: { type:
  'image', image: { path } }` (or `{ data }`, or a bare `image:` with no `type`)
  — and emits `<a:blipFill>` inside `<a:tcPr>`, after the cell's borders. It
  works through every route a fill reaches a cell: per-cell `options.fill`, the
  `headerRow` and `columns[i]` sugar, and the table-level `fill`. Merged regions
  fill uniformly across the whole span, and an auto-paged table registers each
  overflow slide's media on that slide rather than piling every relationship onto
  the first one.

  This is a cell *fill*, not a picture nested in a cell — `CT_TableCell` accepts
  only `a:txBody`, `a:tcPr` and `a:extLst`, so no `<p:pic>` can live inside an
  `<a:tc>` at any effort. To float a real picture over a cell instead, use
  `pptx.tableLayout()` to get that cell's computed rect and `addImage()` at those
  coordinates — which is also how PowerPoint itself fakes it.

  No new API surface: `TableCellProps.fill` was already `ShapeFillProps`, so an
  image fill always type-checked. What was missing was the registration step —
  nothing walked a table's cells to allocate the media relationship, which is why
  it degraded at write time (see Fixed, below). Registration is keyed on fill
  *object identity*, so the `headerRow`/`columns` sugar — which shares one fill
  object across every cell it styles — mints a single relationship rather than one
  per cell.

  Gated on `test/read/fixtures/table-cell-image-fill.pptx`, authored by desktop
  PowerPoint: stretched, tiled, bordered and merged picture-fill cells in one
  table. Two findings from it are worth keeping. PowerPoint writes a bare
  `<a:tcPr/>` on a merged region's *covered* cells and repeats no fill there,
  while this library copies the origin's fill onto them — kept deliberately, since
  a covered cell is never rendered and the copy keeps image fills uniform with the
  solid case. And PowerPoint's *stretched* cell fill omits `dpi="0"
  rotWithShape="1"` and `<a:srcRect/>`, which our shared `genXmlImageFill` always
  writes; that shared emitter was left alone, because both attributes are optional
  with no schema default and PowerPoint authors exactly that attribute set for its
  own *tiled* cell in the same table.

- **A deck IR (`@shbernal/ts-pptx/script`), the read half of turning an existing
  `.pptx` back into source.** `readModelToIr(presentation)` walks a deck read
  through `ts-pptx/read` and returns a serializable description of the write-API
  calls that would rebuild it — `{ slideSize, props, slides, assets, fidelity }`,
  where each slide holds `{ method, args }` calls whose `args` are literal
  write-API option objects. Geometry is carried as exact `"<n>emu"` strings
  wherever the option is `Coord`-typed, and as six-decimal inches for the three
  that are not (`colW`, `rowH`, `margin`) — the proven minimum for an EMU-exact
  round-trip.

  It is a new subsystem rather than part of `ts-pptx/read` because it needs both
  the read model and the write option types, and because the read subpath is
  documented as isomorphic (bytes in, bytes out); a converter whose output is
  source text would break that guarantee for its consumers.

  Every construct that cannot survive is a `FidelityNote` on the IR rather than a
  log line, carrying `{ slideNumber, shapeName, construct, disposition, cause,
  detail }`. `cause` distinguishes a missing read accessor (`unread`) from a
  missing write option (`unwritable`) from a structural limit (`unsupported`),
  which is what makes a note actionable. Because the notes are data, a round-trip
  check can exclude exactly the declared losses and treat any other difference as
  a defect — an undeclared loss fails, and a declared loss that actually survives
  is a stale note. Read `DeckIr.fidelity` before trusting a conversion: notable
  entries include theme-referenced outline width (`p:style/a:lnRef` resolves a
  colour but not a width or dash), embedded audio/video (only the poster frame is
  readable), and OMML equations.

- **`printScript(ir, options)` turns that IR into a runnable TypeScript module.**
  Returns `{ code, assets, notes }`: the module source, the image bytes it
  expects beside it, and the losses that apply to *this* output. The emitted
  script anchors on a template, and the template is the **source deck itself,
  unmodified** — `Presentation.fromTemplate` already strips a package's slides
  while leaving its masters, layouts, theme, and document properties
  byte-identical, so only slide content is regenerated and the deck's whole
  design survives untouched. Slides are emitted in source order; contiguous
  slides sharing a layout share one generator, since `appendSlides` binds one
  layout per call.

  `notes` is not simply `ir.fidelity`. A template-anchored output *rescues* some
  declared losses — all twelve document properties ride in the template, so the
  IR's `deck.docProps` note does not apply and is omitted rather than left to
  teach readers to skim. It also *adds* losses that belong to the tier rather
  than to the conversion: a slide's `p:cSld@name` reads fine and would survive a
  byte copy, but has no public write-API setter. The applicable set is
  reproduced as a header comment in `code`, so the artifact carries its own
  caveats, and it is the set a round-trip check should exclude from its diff.

  Binding is by layout name where that is unambiguous, because a name survives
  being re-pointed at a different template; a deck whose masters repeat a layout
  name falls back to gallery position, since `appendSlides` throws on an
  ambiguous name rather than choosing.

- **`printStandaloneScript(ir, options)` prints the same IR with no template at
  all.** The emitted module depends on nothing but this package: the theme, one
  `defineSlideMaster` per source layout, and every slide, all re-authored through
  the public write API and therefore all editable. It is the second printer over
  one IR, and the only thing the two differ in is where the deck's chrome comes
  from.

  The trade is worth stating before choosing a tier. Template-anchored output
  gets the deck's entire design back byte for byte, at the cost of shipping the
  source deck alongside the script and of leaving that design uneditable.
  Standalone output is one file, and the parts of the original design the read
  model cannot see are gone. Three of them are unreachable from *both*
  directions, so no amount of converter work recovers them: `a:fmtScheme` (the
  fill, line and effect style lists a shape's `p:style` indexes into — no reader,
  and the write path emits a hardcoded Office one), `p:txStyles` (the master's
  per-level text styles — no reader, though `SlideMasterProps.textStyles` could
  author them), and master/layout decoration. A fourth, `p:clrMap`, is readable
  with no setter. Each is a fidelity note in this tier and rides across untouched
  in the other, which is why the other shipped first.

  A `defineSlideMaster` here carries a title and a background and nothing else,
  and that is a write-path constraint rather than a shortcut:
  `addPlaceholdersToSlideLayouts` seeds every slide with each layout placeholder
  the slide did not populate, as an empty text shape. Since this converter
  authors every source shape as concrete absolute-positioned content and binds
  none of them to a placeholder, re-declaring a layout's placeholders would add a
  ghost shape to every slide for each one.

  The second consumer moved the IR twice, both still within this unreleased
  window. `DeckIr` gained `chrome` (`{ theme, masters[] }`), which the
  template-anchored printer ignores entirely since the source deck *is* its
  chrome. And `SlideIr.calls` is now populated for a `carried` slide too: marking
  a slide carried was an erasure and is now a recommendation, because a printer
  with no source package to copy from had the whole slide erased rather than only
  the unwritable construct that made it uncarryable — and that construct already
  declares its own loss.

  Two notes describe the write path rather than the source deck, because both are
  permanent properties of any standalone output: a presentation always carries a
  blank layout named `DEFAULT`, seeded in the constructor with no way to remove
  it (`master.default`), and all five settable document properties are stamped in
  the constructor and cannot be unset — assigning `''` writes an empty element
  rather than removing it (`deck.docPropsDefault`).

- **A guide for the whole subsystem: [PPTX To Script](docs/reference/pptx-to-script.md).**
  The two tiers and how to choose, the chrome cliff that forces the split, the
  fidelity-note contract, the measured loss list across the fixture corpus, and
  what a clean round-trip run does and does not prove.

- **A round-trip check for generated scripts: `canonicalDeckIr` + `diffDeckIr`,
  and `pnpm run script:roundtrip`.** Reads a deck, prints a script, runs it,
  reads the deck that came out, and diffs the two IRs using the printer's
  fidelity notes as the exclusion list — so a difference no note predicted is a
  defect. `diffDeckIr` returns `{ differences, undeclared, declared, added,
  unmatchedNotes }`; `undeclared` is the number to gate on.

  The comparison is on IRs rather than packages because the output can never be
  byte-identical (fresh rel ids, regenerated shape ids), so comparing bytes would
  fail for every deck and measure nothing. `canonicalDeckIr` removes the noise
  that is not loss — a value spelled out that means what its absence means
  (`bold: false`, `wrap: true`, the default `a:bodyPr` insets), and asset
  identity by content digest rather than by generated filename — and each such
  rule cites the OOXML default that makes it an equivalence rather than a
  convenience.

  Its reach is bounded, and knowing how bounded is the point. Both IRs come from
  the same reader, so a construct the read path cannot see is absent from both
  and compares equal; and the converter need not be injective, so two source
  constructs mapping to one call also compare equal. It detects **asymmetry**.
  Mutation testing says so concretely: of twelve deliberate converter defects it
  catches six, and the six that survive are exactly the symmetric ones. Read a
  clean run as "nothing the converter can distinguish was lost", and pair it with
  `pnpm run read:census` and the IR unit tests, whose expectations come from
  `src/types/*.ts` rather than from the converter.

- **`ts-pptx/script` now transcribes a slide's show transition, in both tiers.**
  A `SlideIr` gained a `transition` field and the printers emit it as
  `slideN.transition = { … }` — a property assignment rather than a call, which
  is how the write API models it. Speed bucket, exact `p14:dur` duration,
  `advClick`/`advTm` advance behaviour and the type-specific variant attributes
  (`{ dir: 'd' }`, `{ spokes: '2' }`) all carry across, each omitted from the
  emitted literal when the source left it at its OOXML default.

  The type is filtered against the write API's closed vocabulary. The read model
  reports `TransitionInfo.type` as an *open string*, because it also decodes
  PowerPoint's modern effects (Morph, Vortex, Ripple, …) and tells them apart by
  namespace, while `TransitionType` names the 21 base ECMA-376 transitions and
  nothing else. A name that does not survive the filter files a
  `slide.transition` note instead of producing a script that does not compile.
  PowerPoint's own probed effect table lists exactly those 21 base effects, so the
  filter is checked against ground truth rather than against a transcribed list.

  Transition **sounds** map in both OOXML forms: the stop-previous `p:endSnd`, and
  an embedded start sound whose WAV is resolved through the slide's own `r:embed`
  and carried as an asset. The embedded form survives the standalone tier only —
  `extractSlides` does not surface a transition's audio part, so the append path
  the template-anchored tier rides has nothing to wire and drops the sound
  (silently, with no dangling reference). That tier declares it as
  `slide.transitionSound`.

  The round-trip oracle was widened to match: `CanonicalSlide` now carries
  `transition`, compared structurally rather than as one opaque value, so a note
  can declare a lost *sound* without also excusing a wrong transition *type*.
  Without that, a printer that stopped emitting transitions entirely would have
  produced identical calls and reported a clean round trip.

  Assets are now numbered per media kind (`image1.png`, `audio1.wav`) instead of
  over one shared counter, so a generated script does not bind a transition sound
  to a `const image7`.

### Fixed

- **`addChart` mutated the arrays and options object it was handed.** Series
  normalization was applied in place, so `pptx.addChart(data, opts)` rewrote
  `data[0].labels` from `['A','B','C']` to `[['A','B','C']]` — the nested form
  the multi-level category serializer wants — and stamped an internal
  `_dataIndex` onto every series. Any caller that reused its own data afterwards
  (to build a legend, a table, or a second chart) silently got one nested array
  where it had passed three strings, and the failure surfaced far from the chart
  call. `addChart` now normalizes into copies; the caller's series objects are
  never written back to.

  The options object had the same problem and is fixed with it: defaults
  (`chartColors`, `barGapWidthPct`, `plotArea`, …) were written onto the caller's
  object, and invalid entries were *deleted* from it — an out-of-range
  `layout.x`, a bad `catGridLine.size`, a `dataLabelPosition` illegal for the
  chart type. Sharing one options object across two charts therefore meant the
  second chart saw the first chart's normalization. Options are now copied before
  anything is applied, so `addChart` treats both of its arguments as read-only
  inputs.

  Two consequences worth noting. `_dataIndex` is gone from the public
  `OptsChartData` type — it was only ever there because the normalization wrote
  it onto the caller's object, and it remains on the internal series shape the
  emitters read. And code that *relied* on reading the normalized values back off
  its own options object after the call (the filled-in defaults, the clamped
  percentages) no longer sees them; pass the values explicitly instead. Emitted
  OOXML is byte-identical.

- **A theme color on a chart gridline or series line emitted invalid XML.**
  `valGridLine`, `catGridLine`, `serGridLine` and `barSeriesLine` built their
  color by hand as `<a:srgbClr val="…"/>`, bypassing the shared color emitter
  every other color in the library goes through. A scheme token therefore landed
  verbatim in the `val` attribute — `<a:srgbClr val="accent1"/>`, where the
  attribute is `ST_HexColorRGB` — so gridlines could not follow the deck's theme.
  Both emitters now route through `createColorElement`, which picks
  `<a:schemeClr val="accent1"/>` for a scheme token and `<a:srgbClr>` for hex.
  `OptsChartGridLine.color` widens from `HexColor` to `Color` accordingly, so
  `valGridLine: { color: SchemeColor.accent1 }` now type-checks and renders.

  Sharing that emitter also brings the gridline path in line with every other
  color site: a leading `#` is stripped, an 8-digit RGBA value splits its alpha
  byte into a sibling `<a:alpha>`, and an unparseable color warns and falls back
  to the default instead of being written out as-is. One cosmetic consequence:
  hex is now normalized to uppercase, so a gridline authored as `'d9d9d9'` emits
  `val="D9D9D9"`. Rendering is identical and the built-in defaults are unchanged,
  but a test that pins chart XML bytes for a lowercase gridline color will see a
  one-time diff.

- **A negative `w`/`h` produced a deck PowerPoint refused to open at all.** The
  signed value went straight into `<a:ext cx=… cy=…>`, and both attributes are
  `ST_PositiveCoordinate` — so one negative extent anywhere cost the *whole*
  presentation: *"The file or directory is corrupted and unreadable"* (0x80070570),
  naming no shape, no part, and no slide. LibreOffice rendered the same package
  happily, so a pipeline that previews with LibreOffice saw nothing wrong until
  the deck reached PowerPoint.

  It was easy to hit, because a signed delta is the natural way to write "draw
  from A to B": `addShape('line', { x: x0, y: y0, w: x1 - x0, h: y1 - y0 })`
  works fine until the line happens to run leftward or upward. A negative
  extent is now normalized to the box PowerPoint itself would write for that
  geometry — origin at the min corner, absolute extent, and a flip on the
  mirrored axis — which is the encoding `addConnector` has always derived from
  its endpoints. `{ x: 1, y: 3, w: 1.5, h: -2 }` emits
  `<a:xfrm flipV="1"><a:off x="914400" y="914400"/><a:ext cx="1371600" cy="1828800"/></a:xfrm>`.

  It applies to every object kind, not just `addShape`, because it happens at the
  one point where all of them share a placement path — after each `Coord` form
  has resolved to EMU, so `'-25%'` and `'-2in'` normalize alongside a plain
  negative number. A derived flip XOR-composes onto an explicit one, so
  `{ w: -2, flipH: true }` is mirrored twice and therefore not at all. Group
  auto-bounds normalize each child before taking the bounding box; previously a
  child with a negative extent reported a `maxX` left of the group's `minX` and
  collapsed the group frame.

  No warning is emitted, and `Math.min`/`Math.abs` at the call site is no longer
  needed — the signed form is now a supported spelling rather than a trap.

- **`addAnimation()` and `groupObjects()` could not find a shape whose
  `objectName` contained `&`, `<`, `>`, `"`, `'`, a tab or a newline.** A shape
  added as `objectName: 'Q&A'` is stored attribute-escaped (`Q&amp;A`) so it can
  be written into `<p:cNvPr name>` as-is, and the two lookups compared the
  caller's raw string against that escaped text. Neither ever matched:
  `addAnimation({ preset: 'fadeIn', objectName: 'Q&A' })` warned *"no object
  named "Q&A" on the slide"* and dropped the effect — leaving the deck with no
  `<p:timing>` at all — and `groupObjects(['Q&A'])` threw *"no top-level object
  on this slide has that objectName"*, both naming a shape that was plainly
  there. Renaming to `QandA` was the only way through. Any `objectName` now
  works as a lookup key, including for a shape inside a group and for the
  "already inside a group" hint that tells a grouped name apart from a typo.
  Escaping is now done once, where the comparison happens
  (`resolveObjectNameToId`), rather than re-derived at each call site — which is
  how the animation lookup came to disagree with the connector one, the only
  caller that had it right.

- **An image fill on a table cell type-checked, then silently rendered as no
  fill.** `TableCellProps.fill` has always been `ShapeFillProps`, so `fill: {
  type: 'image', image: { path } }` on a cell compiled and flowed all the way to
  the emitter — but nothing on the table path ever called
  `registerImageFillMedia`, so no media relationship was allocated. The fill
  reached `genXmlColorSelection`'s `case 'image'` with no resolved `_imgRid`,
  warned *"image fill is missing its resolved media reference"*, and emitted
  `<a:noFill/>`. The package contained no `ppt/media/` entry and the cell rendered
  blank. Shapes (`gen/define/shape.ts`) and text boxes (`gen/define/text.ts`) had
  been wired since image fills were introduced; tables were the one major object
  kind left out. Cell image fills now resolve properly — see Added, above.

- **A transition sound supplied as `data:audio/x-wav;…` was written to a media
  part named `.x-wav`.** The media filename was taken from the data URI's mime
  *subtype* verbatim, so the package ended up with `ppt/media/audio-1-1.x-wav`
  and a `<Default Extension="x-wav"/>` declaring a file type that exists nowhere
  else. The content type was right and PowerPoint opened it, but nothing else
  would recognise the file. The subtype now maps to a real extension
  (`audioExtensionForSubtype`), so the same bytes land on `audio-1-1.wav` under
  `<Default Extension="wav" ContentType="audio/x-wav"/>` — what PowerPoint itself
  authors. `audio/x-wav` is not an exotic input: it is exactly the content type
  PowerPoint writes for an embedded transition sound, so it arrives on every deck
  read back in and handed to `ts-pptx/script`.

- **A tab, carriage return or line feed inside an XML attribute value was emitted
  literally, so it read back as a space** (`dn-xml-attr-whitespace`). XML 1.0
  §3.3.3 requires a parser to normalise those three characters to a single space
  inside an attribute value *before any consumer sees them*; carrying one across
  needs a character reference. Every caller-supplied string that lands in an
  attribute was affected — `objectName` (`p:cNvPr/@name`), alt text
  (`p:cNvPr/@descr`), layout and slide titles (`p:cSld/@name`), section titles
  (`p14:section/@name`) and hyperlink tooltips — so a two-line layout title came
  back as one line. This is not theoretical: PowerPoint's built-in German layout
  set ships a layout named across two lines ("Abschnitts-\<LF\>überschrift").

  The fix is a new `encodeXmlAttrValue` used by the attribute-emitting paths only
  (the element builder in `src/gen/oxml/el.ts`, plus the few emitters that write
  attributes with template strings). It is deliberately *not* a widening of
  `encodeXmlEntities`, which also escapes element text — there a literal newline
  is meaningful content, and escaping it would change bytes across every
  text-bearing part in the package.

  **This changes emitted bytes** for any deck whose attribute values contain a
  tab, CR or LF; every other deck is byte-identical (verified against the full
  1637-part demo deck). Consumers that string-matched the emitted XML for such an
  attribute must now match `&#9;`/`&#10;`/`&#13;`. Reading is unaffected: any
  conforming parser resolves the references back to the original characters.

- **Nine converter defects that no static check and no execution check could
  catch**, all found by the round-trip comparison above and all producing a
  script that typechecked, ran, and wrote a plausible deck.
  - **Every paragraph bullet was wrong.** `Paragraph.bullet` reports a *tagged*
    string (`'none'`, `'char:<glyph>'`, `'autoNum:<type>'`) and the mapper read
    it as a bare glyph — so an explicit `a:buNone` became a literal `n` bullet
    (`'none'.codePointAt(0)`), a real character bullet became `c`, and a numbered
    list became `a`. The numbering scheme was also passed as `style` rather than
    `numberType`, so every numbered list fell back to the default scheme, and the
    glyph was emitted with fewer than the four hex digits the write path requires,
    which made it substitute its own default.
  - **Placeholders were emitted with no geometry at all.** A shape with no
    transform of its own reported no frame, and omitting `x`/`y`/`w`/`h` does not
    leave the geometry to be inherited — an appended slide inherits nothing — it
    produces a zero-height box in the corner. Now resolved through the
    layout/master chain via `resolvedFrame`.
  - **A group's rotation and flips were applied twice**, because the children were
    already emitted in composed slide-absolute coordinates *and* the transform was
    repeated on the group. A 30° group came back at 60°; a flipped group came back
    unflipped, the double flip having cancelled.
  - **Image crops were sent to `sizing`**, which crops in displayed inches against
    the image's measured natural size, instead of to `crop`, which is `a:srcRect`
    emitted verbatim. Fractions read as inches shrank every cropped picture.
  - **Every text body was re-anchored to centre.** An unset `a:bodyPr/@anchor`
    means top in OOXML and centre in `addText`, so the converter now spells the
    anchor out rather than leaving it off.
  - **Every uncoloured run was repainted black**, since `addText` fills a run with
    no colour using `DEF_FONT_COLOR`. The inherited colour is now resolved and
    emitted (with a note that it is frozen against later theme edits), or, where
    nothing resolves it, declared as a colour that may be wrong.
  - **PowerPoint text boxes were demoted to auto shapes.** Text-box-ness is
    `p:cNvSpPr/@txBox`, not "has no preset geometry" — PowerPoint gives every text
    box an explicit `prstGeom rect`, so the old test misclassified all of them.
  - **An SVG picture was reduced to its raster fallback.** The vector part is now
    preferred, since `addImage` accepts SVG bytes and regenerates a fallback.
  - **A bulleted paragraph with more than one run was split in two**, because the
    write path treats a bullet on a run that does not open a line as a request for
    a new paragraph. Paragraph-level `bullet` now rides on the first run only.
  - Also declared rather than fixed, each being a real limit rather than a
    mistake: an automatic field (`a:fld` — slide number, date, footer) has no
    accessor and no `addText` expression; an explicit zero `a:spcBef`/`a:spcAft`
    is indistinguishable from unset on the write side; an auto-height table row
    comes back pinned to its content's height; and a paragraph that inherits its
    bullet cannot say "inherit" through the write API.

- **`line: { type: 'none' }` did nothing.** The value is documented on
  `ShapeFillProps.type` and accepted by the type checker, but `genXmlLineFill` had
  no branch for it and emitted no paint child — which is how a stroke says
  "inherit", not "none". A shape authored with an explicitly suppressed outline
  therefore *grew* the theme's border instead of losing it. It now emits
  `<a:noFill/>`.

- **`readModelToIr` mapped three constructs onto write-API shapes that do not
  exist.** All three produced an IR that typechecked and a script that failed at
  run time, because `IrValue` is deliberately loose enough that `tsc` cannot
  check an argument against the signature it is meant to satisfy.
  - `addChart` received the chart type as a third positional argument; the
    signature is `addChart(data, options & { type })`, so the type now rides in
    the options object.
  - `addConnector` received a bounding box (`x`/`y`/`w`/`h`) and a nested `line`
    object. It takes two endpoints and flat stroke options, so connectors now
    emit `x1`/`y1`/`x2`/`y2` derived from the box **and its `a:flipH`/`a:flipV`
    flags** — without the flips every up- or leftward connector is silently
    mirrored — plus `color`/`width`/`dashType`/arrowheads. A gradient or
    translucent connector stroke has no flat spelling and is now noted.
  - A slide bound to a layout whose name is shared by another layout in the deck
    made `appendSlides` throw. `SlideIr.layout` now carries the gallery index and
    whether the name is unique.

- **Three shapes could vanish from a conversion with no fidelity note**, against
  the contract that a dropped shape is never silent: an auto shape with neither
  text nor geometry of its own (an unfilled placeholder, whose outline comes from
  the layout), and a group whose every child was dropped. Both are now declared
  (`shape.empty`, `group.empty`). The third path — an unrecognised shape kind —
  turned out to be unreachable, which the type checker proves.

- **`Presentation.appendSlides` now carries speaker notes.** A generator slide
  authored with `addNotes` previously lost its notes entirely when spliced onto a
  loaded deck — `extractSlides` emitted no notes part, so the append path had
  nothing to wire. It now emits a `notesSlide` per notes-bearing slide, wired back
  to the slide it annotates, with the notes body's own hyperlink relationships
  preserved (`rId1` = notes master, `rId2` = slide, hyperlinks from `rId3`).

  A notes slide must bind to a notes master, and a template usually has none — a
  deck authored without speaker notes carries no `notesMaster` part at all — so the
  generator's own notes master (and the theme its `.rels` requires) rides along in
  `ExtractedSlides.notesMaster` and is installed **only** when the destination deck
  has none. A destination that already has a notes master keeps it, so its notes
  styling wins; this matches the existing `importNotes` policy. `ExtractedSlide`
  gains an optional `notes` field for the same reason.

- **`markDirty()` on every read-model class that exposes `element_`.** `element_`
  hands out the live DOM node, but an edit through it was silently discarded on
  `save()` unless the caller reached the owning part themselves
  (`shape.slide.part.markDirty()` — three hops, undocumented). The obligation now
  sits on the same object as the hatch: `Slide`, `Shape`, `TextFrame`,
  `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `Placeholder`,
  `NotesPlaceholder`, `Theme`, `Chart`, `ChartAxis`, `ChartSeries`, `ChartEx`,
  `ChartExAxis`, `ChartExSeries`, and `ResolvedTableStyle`. `Shape.markDirty()`
  was `protected` and is now public.
- **`Slide.element_`** — the `p:sld` root, filling the one missing rung in the
  `element_` ladder. Slide-level DOM access was previously only reachable as
  `slide.part.dom`.
- Guide formulas passed to a `custGeom` shape's `guides` option now have their
  leading operation checked against the 17 operations ECMA-376 §20.1.9.11
  defines. An unknown operation (e.g. `{ name: 'w2', formula: 'bogus 1 2' }`)
  previously emitted schema-shaped but semantically dead geometry whose first
  feedback was a PowerPoint repair prompt; it now warns and skips the guide,
  matching the existing degenerate-entry behaviour. Operands are still passed
  through uninterpreted.
- The project's escape-hatch policy is now written down in
  [project target](docs/project-target.md) — the convenience-vs-guarantee rule
  and why the read path gets a deep raw hatch while the write path does not.
- **Four aggregate checks**, replacing the practice of hand-composing four or
  five scripts per iteration. `verify` (~45s) is the three typechecks +
  `backlog:validate` + the whole test suite; `verify:full` (~65s) adds
  `package:lint` and `test:package`. `check:static` and
  `check:package` are the two halves CI runs as separate jobs. `verify` and
  `verify:full` omit `lint`/`format:check` by design — the git hooks own those.
- **A `dist/` freshness guard** (`scripts/ensure-dist.mjs`) that every test,
  typecheck, and package script now starts with. It rebuilds only when `src/` or
  a build config is newer than `dist/`, and is a ~0.1s no-op otherwise. This
  replaces both halves of the old pattern — the unconditional
  `pnpm run build &&` prefix and the `:fast` twins that skipped it — so there is
  no longer a stale-`dist/` footgun to reason about.

### Changed

- **`demos/` are showcases now, and nothing there is a test.** Two flagship decks
  live in `demos/showcases/` and build with `pnpm demos:build [slug]` — a
  corporate quarterly review (themed colour scheme, five masters, native
  gradients, grouped KPI cards, three chart types, a styled table, speaker notes)
  and an image-led photo essay (full-bleed photography, gradient scrims, a
  duotone picture effect, an embedded video, a live hyperlink). They replace
  `demos/modules/`, 7,100 lines of feature-enumerating builders that made a deck
  nobody would show anyone.

  The demo smoke test (`test:demos`, `test:demo:node`, `test:demo:vite`,
  `scripts/demo-smoke.mjs`) is gone with it, and `check:package` and
  `verify:full` no longer chain it. Its actual job — proving the built package
  works for a consumer — belongs to `test:package`, which now imports all nine
  export subpaths out of an installed tarball under both npm and pnpm and forces
  the `browser` condition. One signal did not survive and is recorded as an
  accepted gap in [testing](docs/testing.md#demos-are-not-tests): the Vite build
  was the only check that put a real bundler in front of the package.

  The review deck imports nothing from `node:`, so `demos/vite-demo` imports that
  same module and builds the identical deck in a browser rather than keeping a
  second copy of demo code.

- **Breaking (internal constructors):** `ChartAxis`, `ChartSeries`, and
  `ChartExAxis` now take the owning chart part (respectively the owning `ChartEx`)
  as a second constructor argument, so `markDirty()` can reach it. These are
  obtained from `Chart.axes` / `Chart.series` / `ChartEx.axes`; only code that
  hand-constructed them is affected.
- **Breaking (structural type):** `ResolvedTableStyle` gained a required
  `markDirty(): void` member. Only code that builds the object literal itself
  (rather than reading it from `table.resolvedStyle`) is affected.
- **Breaking (internal type):** the internal `ShadowPropsInternal.opacity`
  field (the derived shadow alpha) is renamed to `_alpha`, clearing it of the
  removed public `opacity` shadow input. A stray `opacity` from an untyped/legacy
  caller is now inert (it lands on a field nothing reads) instead of being
  actively stripped — no behaviour change for supported inputs; use
  `transparency` (0–100). Only code reading the internal shape off a corrected
  shadow is affected.
- `pnpm run byte-identity:baseline` now refuses to run when `src/gen/` has
  uncommitted changes (override: `--allow-dirty`). A baseline frozen after the
  refactor has begun records the very bytes it exists to detect, so every later
  `check` passes trivially. The error names the workaround it is closing —
  `git stash` on a dirty tree, which risks unrelated work to a pop conflict.
- **The byte-identity gate builds the showcase decks.** It had built its corpus
  by importing `demos/modules/demos.mjs`, removed in the demos-to-showcases move
  above, so both subcommands died on `MODULE_NOT_FOUND` — worse than no gate,
  because a harness nobody can run still gets cited as one. It now builds every
  deck registered in `demos/showcases/lib/showcases.mjs` (a registry `build.mjs`
  reads too, so a new showcase is gated the day it lands), writes them under
  `.tmp/byte-identity/decks/` rather than over the artifacts `pnpm demos:build`
  leaves for a human, and explodes each under its own slug so a diff names the
  deck that moved. `Math.random` is reseeded per deck: on a single stream,
  editing the first deck shifts every GUID in the second, and the gate reports a
  diff in a deck nobody touched.

  The corpus is 177 parts against the old deck's 1637. Every part *kind* survives
  — charts with their embedded workbooks, media, notes slides, masters, layouts,
  themes — but it is a narrower slice of the emitters, so AGENTS.md now says to
  confirm the part you touched is in the baseline before reading a PASS as proof
  of anything.

- The three `tsc` projects are now `incremental`, keeping their build state in
  the gitignored `.tmp/` (a distinct `tsBuildInfoFile` per project). Warm
  `typecheck` drops from ~3.4s to ~1.3s; cold runs are no slower, so CI is
  unaffected.
- The OOXML schema fixtures now run concurrently, taking that suite from ~50s to
  ~10s — cheap enough that `verify` runs it on every iteration instead of
  reserving it for `verify:full`.
- The root build configs (`eslint.config.mjs`, `vitest.config.ts`,
  `tsdown.config.ts`, `tsdown.dev.config.ts`) are now linted and typechecked.
  They previously matched no ESLint `files` block and no tsconfig `include`, so
  nothing checked them at all.
- A missing OOXML validator now **fails** the read suite under `CI` instead of
  silently skipping a few hundred schema assertions. Locally it still skips, but
  prints a notice — a green local run no longer reads as a complete one.
- CI runs the static checks once rather than once per Node version, and adds a
  Windows leg for the package and demo scripts, which are the only exercise the
  Windows-specific subprocess handling in `scripts/script-utils.mjs` gets. The
  publish workflow now reuses the CI workflow instead of keeping its own copy of
  the gate, which had already drifted out of sync.

### Fixed

- **A theme-indexed picture background reported the wrong part.** When a
  `p:bgRef`'s `fmtScheme` entry is an `a:blipFill` — the third `bgFillStyleLst`
  slot in several stock Office themes (Ion, Facet, …) — the fill element comes out
  of the *theme* part, so its `r:embed` is scoped to the theme's relationships.
  `SlideBackground.resolvedFill` resolved it against the owning
  slide/layout/master's relationships instead, which does not fail loudly: the
  same id usually exists there and points at something else entirely. On an
  Ion-themed deck `resolvedFill.partName` read `/ppt/slideLayouts/slideLayout1.xml`
  where the image is `/ppt/media/image1.jpeg`. Affects `Slide.background`,
  `SlideMaster.background`, and `SlideLayout.background`.

[3.6.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.6.0
[3.5.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.5.0
[3.4.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.4.0
[3.3.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.3.0
[3.2.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.2.0
[3.1.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.1.0
[3.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v3.0.0
[2.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v2.0.0
[1.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v1.0.0
