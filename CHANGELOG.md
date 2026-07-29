# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

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

## [1.0.0] - 2026-07-24

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

[1.0.0]: https://github.com/shbernal/ts-pptx/releases/tag/v1.0.0
