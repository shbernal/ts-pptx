---
doc-schema-version: 1
title: "Measured Text Fit"
summary: "How fit:'shrink'/'resize' compute and bake an autofit result so overflowing text self-corrects in headless renders, and the PowerPoint-authored oracle that calibrates it."
read_when:
  - Changing the autofit / text-fit solvers or font metrics
  - Touching registerFontMetrics or the measured-fit export pass
  - Regenerating or interpreting the autofit calibration fixtures
  - Understanding why fit:'shrink'/'resize' now bake a value instead of a bare flag
  - Measuring Chinese, Japanese or Korean text, or changing where lines may break
doc_type: "decision"
---

# Measured Text Fit

## Status

**Shipped (2026-06-21).** Measured fit for `fit:'shrink'` and `fit:'resize'`
(text boxes) and `TableCellProps.fit:'shrink'` (table cells) is implemented and
calibrated against PowerPoint-authored fixtures. Source:

- `src/measure/font-metrics.ts`: opentype.js metrics provider (+ heuristic
  fallback), the `FontMetricsRegistry`, and `makeRegistryResolver`.
- `src/measure/text-fit.ts`: `wrap=square` simulator and the shrink/resize solvers.
- `src/measure/paragraphs.ts`: turns an authored slide object into the
  simulator's inputs (runs → `FitParagraph`s, insets, box, anchor share).
- `src/measure/table-fit.ts`: `computeTableLayout` and the cell-grid walk.
- `src/measure/fit.ts`: the export-time pass that measures text boxes and table
  cells and rewrites the slide object before the sync XML build.
- `pptx.registerFontMetrics()`: the public registration API.

Held conservative by `test/read/autofit-calibration-oracle.test.js` (computed
shrink `fontScale` ≤ PowerPoint's; computed resize `cy` ≥ PowerPoint's *and* ≥
LibreOffice's). See `CHANGELOG.md` for the change history.

Downstream driver: a consumer hitting text overflow (text spilling out of
cards/components) in a headless render pipeline.

## The problem

The `fit` autofit markup already existed in the project and downstream consumers
already used `fit:'shrink'` heavily, but the project emitted the **bare flag with no
baked result** (`<a:normAutofit/>` / `<a:spAutoFit/>`). A bare flag defers the fit
computation to an interactive edit/resize event that a headless render never
fires, so in a headless LibreOffice → PNG pipeline (and on plain file-open)
nothing recomputes and the text still overflows.

This is **not** an inherent renderer limitation: both PowerPoint and LibreOffice
honor a fit that is already baked into the file. The catch is that the two
mechanisms store the fitted answer in **different places**, so the fix is not
symmetric:

| `fit` value | bare emit | where the fitted answer lives | renderer on plain open |
| --- | --- | --- | --- |
| `'shrink'` | `<a:normAutofit/>` | `fontScale`/`lnSpcReduction` **attrs on the element** | applies stored scale; bare = no attrs = 100% = no shrink |
| `'resize'` | `<a:spAutoFit/>` | the shape's `a:ext/@cy` (+ `a:off/@y`) in `xfrm` | draws the box at whatever `cy` is already written |

The missing capability was **measurement**: the library had no font metrics, so it
could not compute the value to bake. This feature adds serialize-time measured fit
so overflow self-corrects in headless renders without a human round-trip through
PowerPoint. It is generic ts-pptx value (every consumer that renders or ships
pptx headlessly hits it), so it lives upstream, not as a downstream workaround.

## Design

### Font metrics provider (`src/measure/font-metrics.ts`)

- Loads a font file (TTF/OTF) and exposes per-glyph advance widths +
  ascent/descent/line-gap. **`opentype.js`** (new dependency, lazily imported,
  Node/web only). A **TrueType Collection (`.ttc`) is not loadable**:
  `parseFontMetrics` throws `Unsupported OpenType signature ttcf`. That is worth
  knowing before registering an East Asian face on Windows, where MS Gothic, Yu
  Gothic, SimSun and Microsoft YaHei all ship as `.ttc`; Malgun Gothic is a plain
  `.ttf` and is what the CJK fixture uses.
- Width is summed from **raw `charToGlyph` advances**: deliberately **no**
  GPOS/GSUB shaping. Kerning almost always narrows a line, so summing raw advances
  over-estimates width, the conservative direction (shrink a touch too much, never
  overflow). (`getAdvanceWidth()` also runs shaping and throws on unsupported
  lookups, so it is avoided.)
- `FontMetricsRegistry` is keyed by `(face, bold, italic)` with a regular
  fallback, since variant advances differ. Parsed fonts are cached.

### Registration (`pptx.registerFontMetrics`)

Consumers tell the library where a face's file lives, since the project does not
embed/register fonts on the write side:

```ts
await pptx.registerFontMetrics('Aptos', '/path/to/Aptos.ttf')
await pptx.registerFontMetrics('Aptos', aptosBoldBytes, { bold: true })

slide.addText(runs, { x, y, w, h, fontFace: 'Aptos', fit: 'shrink' })
// → with metrics registered, exports <a:normAutofit fontScale="83000"/> computed
//   to fit; without metrics, exports bare <a:normAutofit/> + warns once.
```

`source` is a path/URL or raw `Uint8Array`/`ArrayBuffer`; font bytes are loaded
via `RuntimeAdapter.loadFontData` (node `fs` / browser `fetch`).

### Wrap simulator + solvers (`src/measure/text-fit.ts`)

- Greedy `wrap=square` line breaking that mirrors PowerPoint: break on whitespace
  or either side of a per-character breaking code point (below),
  hard-break over-long tokens, honor `\n`, `charSpacing`, bold/italic metrics, and
  multi-run paragraphs. Output: wrapped line count + laid-out height for a font
  scale.
- Single-spacing line pitch is **1.2117 × fontSize** (an oracle finding, below).
- **shrink** solver: search `fontScale` on PowerPoint's 2.5% grid down to a 25%
  floor until laid-out height ≤ inner height. `lnSpcReduction` is left at 0
  (fontScale-only is provably ≤ PowerPoint's, the conservative direction).
- **resize** solver: reuse the wrap simulator at `fontScale=100` and return the
  needed inner height. Resize has **no safety net** (under-estimate → overflow,
  there is no text-scaling fallback), so it **must err tall**.
- Conservative WIDTH/HEIGHT safety factors (`1.03`/`1.04`) approximate
  PowerPoint's device-DPI advance rounding, so the simulator wraps slightly early.

### East Asian line breaking

Chinese and Japanese text has no spaces to break at: PowerPoint breaks it between
any two characters. Tokenizing such a run as one unbreakable "word" moves the whole
run to the next line, wastes the rest of the current one, and over-reports the line
count, so `fit:'shrink'` bakes a `fontScale` PowerPoint would never have chosen and
`measureText` reports a vertical overflow that is not there. `isCjkBreakCharacter`
in `src/measure/text-fit.ts` therefore makes each such code point its own wrap
opportunity (a break is allowed either side of it, whitespace or not).

**Hangul is deliberately excluded**, and this is the part that is easy to get wrong.
UAX #14 permits breaking Korean between syllables, but PowerPoint does not do it:
Korean is written with spaces between words and the app breaks it like Latin. A
Hangul run longer than a line still breaks between syllables, but that is the
over-long-token character-wrap fallback every unbreakable word gets, not a break
class. Adding Hangul to the break set would *under*-report the line count, which is
the direction that overflows, the one thing the resize path has no safety net for.

Both halves are pinned by `autofit-cjk-wrap.pptx` and
`test/read/cjk-line-breaking-oracle.test.js` (see [Calibration oracle](#calibration-oracle)).
Two limitations are recorded rather than fixed:

- **No kinsoku.** PowerPoint will not start a line with `。` or `、` and hangs them
  past the right inset instead; this model breaks before them. Same line count,
  narrower widest line, so the height stays conservative.
- **No font fallback.** PowerPoint silently substitutes another face for a code
  point the named font lacks; the registry charges the named font's default advance
  instead. That is a metrics gap rather than a break-class one, and it is why the
  two fixture cases whose glyphs Malgun Gothic lacks are skipped by the oracle test.

### Integration (`src/measure/fit.ts`)

- A measured-fit pass runs **during async export**, before the `gen/` emitter's sync body
  build. It extracts paragraphs/runs (mirroring `gen/slide/objects/` grouping + inheritance),
  computes the inner box from `w`/`h`/insets/margin, then:
  - `'shrink'` → rewrites `fit:'shrink'` to the object form with the computed
    `fontScale`, so `gen/drawingml/text-run.ts` emits `<a:normAutofit fontScale=…/>`.
  - `'resize'` → rewrites `options.h` (and `options.y` per the resolved vertical
    anchor) as `"<emu>emu"` strings, so `gen/slide/objects/text.ts` emits the baked `ext.cy`/`off.y`
    while keeping the `<a:spAutoFit/>` marker. `off.y` shifts by 0 / half / full of
    the height delta for anchor `t` / `ctr` / `b`.
- The `gen/` emitter consumes pre-computed values only.
- The feature is driven off the existing `fit` value, so consumers opt in by what
  they already write: `fit:'shrink'`/`'resize'` measure automatically **when
  metrics are registered** (zero API churn). Per the project's no-back-compat-obligation
  policy, changing what `'shrink'` emits when metrics are present is acceptable;
  the no-metrics path is unchanged.

### Table cells (`TableCellProps.fit:'shrink'`)

PowerPoint has **no** text-autofit for table cells: `a:tcPr`
(`CT_TableCellProperties`) carries no autofit child and the app ignores
`normAutofit` inside a cell txBody (rows auto-grow instead). So a cell's
`fit:'shrink'` (also cascades from a table-level `fit:'shrink'`) cannot bake a
`fontScale`. Instead, `measure-fit.ts` walks the cell grid (colspan/rowspan via an
occupancy sweep; column widths from the shared `resolveTableColWidthsEmu`; row
heights from `rowH`/table `h`; cell margins + table→cell inheritance), runs the
**same** shrink solver, and bakes a **reduced literal font size** (floored to
0.1pt) onto the cell runs: which both PowerPoint and LibreOffice render
identically with no edit/resize. Only fixed-height rows are touched; auto-height
rows are skipped (they grow). Options objects are cloned before mutation because
plain-string cells share the table's single `opt` object. `'resize'`/object forms
are ignored for cells (a row already auto-grows ≈ spAutoFit).

A precondition bug was fixed along the way: auto-width tables (`w` without `colW`)
emitted ~0-EMU `gridCol` widths; `gen/slide/objects/table.ts` now divides the resolved EMU width via
`resolveTableColWidthsEmu`.

### Unregistered-font heuristic

When a deck registered **some** metrics, a `fit:'shrink'`/`'resize'` box or cell
whose **named** face has no exact metrics falls back to a conservative
average-advance table (`getHeuristicFontMetrics`) and still bakes an approximate
result + warns once, instead of degrading to the bare flag. A deck that registers
no metrics at all is unaffected (measured fit stays off); an unnamed
(theme-default) face stays unmeasurable (the face cannot be guessed).

The **layout-time** side (`measureText`) resolves runs identically, with one
deliberate difference: the empty-registry case. `applyMeasuredFit` reads "no metrics"
as "this deck never opted into measured fit" and bakes nothing, but `measureText` is
a read-only query with nothing to opt into, so it returns heuristic numbers and stays
useful with zero setup. The two therefore diverge for an empty registry (the query
predicts a shrink the export will not bake) and that is intentional, not a bug:

| | `measureText` | `applyMeasuredFit` |
|---|---|---|
| kind | read-only query | mutates the deck at export |
| empty registry | heuristic, useful with zero setup | no-op: the deck never opted in |

The library cannot tell "never intended to register" from "intended to register and
the load silently failed"; detecting the latter is the caller's job, at font-load
time. `measureText` reports every named face it had to guess at in
`approximatedFaces`, so a caller that needs exact numbers can check rather than
assume. `measured-fit-integration.test.js` pins this divergence so it cannot change
silently.

## Layout-time measurement (public API)

The same calibrated engine is also exposed for **layout-time** use, so a consumer
can size its own geometry *before* export (grow a card to fit its text, reflow a
grid, detect overflow) instead of relying only on the export-time bake. For any deck
that opted into measured fit (i.e. registered at least one face), a layout-time
prediction must never disagree with what the export then bakes, so both paths share
one converter (`buildFitParagraphs`), one resolver (`makeRegistryResolver`), and one
layout function (`measureLayout`): there is no second wrap model. The one deliberate
exception is an **empty** registry (see below). Source: `src/measure.ts` (subpath entry), `measureText` +
`buildFitParagraphs` in `src/measure/paragraphs.ts`, `makeRegistryResolver` in
`src/measure/font-metrics.ts`, and `measureLayout` in `src/measure/text-fit.ts`.

### Instance methods (inches/points, reuse registered metrics)

```ts
await pptx.registerFontMetrics('Aptos', '/path/Aptos.ttf')

const m = pptx.measureText('A long heading…', { wIn: 3, fontSize: 18, fontFace: 'Aptos' })
// m.heightIn   → laid-out height (conservative/tall — matches the resize bake)
// m.lineCount  → wrapped line count
// m.widestLineIn → width of the widest laid-out line (natural width when wIn is
//                  unconstrained; widest wrapped line otherwise; errs slightly wide)
// m.measurable → false only for an unnamed theme-default face
// m.approximatedFaces → named faces guessed via the heuristic ([] when all exact)
// m.fitsBox(hIn)         → does it fit a box of inner height hIn?
// m.shrinkScaleFor(hIn)  → the fontScale (%) that fits hIn (100 if it already does)

if (pptx.overflowsBox(text, { wIn, hIn, fontSize, fontFace })) warn() // conservative
```

`measureText` is **synchronous** and assumes metrics are pre-registered (the async
`registerFontMetrics` runs ahead of time; lookup is sync). Resolver semantics match
the export pass run-for-run: exact metrics → conservative heuristic for any **named**
face without exact metrics (reported in `approximatedFaces`) → `measurable:false`
only for an unnamed theme-default face. The one place the two paths differ by design
is an empty registry: see "Unregistered-font heuristic" above. Units are inches
(width/height) + points (type/spacing); `insetIn` is subtracted from `wIn` on both
sides if a raw box width is passed.

Because the model errs **tall** (the same `WIDTH_SAFETY`/`HEIGHT_SAFETY` factors as
the resize bake), `heightIn` is ≥ what PowerPoint/LibreOffice render: right for
"grow a container", and why `overflowsBox` is a *conservative* (slightly
over-reporting) check suited to a build-time **warning**, not a hard gate. An
unmeasurable face makes `overflowsBox` return `false` (no false positive).

### Standalone primitives (`ts-pptx/measure`)

For a consumer that lays out without a `TsPptx` instance, the subpath re-exports
the pure pieces so it can build its own resolver/registry and measure directly:
`measureLayout`/`measureHeightPt`/`solveShrink`/`solveResize`, the
`FitParagraph`/`FitBox`/`MetricsResolver`/`Shrink-`/`ResizeOutcome` types, the
calibration constants, and `parseFontMetrics`/`getHeuristicFontMetrics`/
`FontMetricsRegistry`. `opentype.js` stays lazily imported (only `parseFontMetrics`
pulls it in), keeping the subpath cheap to import.

A regression (`test/regression/text/measure-text-api.test.js`) asserts the no-drift
contract: `measureText`'s height equals the height `solveResize` (the export bake)
computes for the same input.

## Calibration oracle

The solvers reproduce PowerPoint's own layout decisions closely enough that shrink
never under-shrinks and resize never under-grows. That cannot be validated against
prose or self-generated XML, so the model is calibrated against **PowerPoint-authored
fixtures where PowerPoint itself baked the fit value**, held as a conservative
regression target.

The hard problem the fixtures pin down is **vertical line metrics**: laid-out
height = `lineCount × lineHeight`, and `lineHeight` is where renderers diverge ("single"
spacing is font-metric-derived; a font carries hhea vs OS/2 win-* vs typo-* pairs,
gated by the `USE_TYPO_METRICS` bit; PowerPoint and LibreOffice can pick different
pairs). A consumer may render through headless LibreOffice, but the file must also
be correct in PowerPoint, so the solver is conservative against the **taller** of
the two: which is why the fixtures measure both engines.

The axes are split to avoid a combinatorial explosion: **per-font metric
calibration** sweeps all 5 fonts on a small core (this pins each family's advance
widths and effective line height), while **policy calibration** (discrete
fontScale steps, lnSpcReduction trade, per-anchor `off.y`) is font-independent and
swept on a single anchor font.

### Fixture decks

Four desktop-PowerPoint-authored decks live in `test/read/fixtures/` (inspection
only: not loaded by `test:read`). PowerPoint baked every fit value
non-interactively on `SaveAs`. Fonts: **Aptos, Aptos SemiBold, Calibri, Tahoma,
Arial**.

- `autofit-line-metrics.pptx` (90 cases): per-font single-line height and advance
  widths straight from XML, plus a LibreOffice cross-measure column.
- `autofit-shrink.pptx` (29 cases): `normAutofit` calibration: per-font core +
  Aptos policy sweep (overflow ladder, lnSpcReduction onset, line spacing,
  space-before/after, multi-run, insets, charSpacing, anchor).
- `autofit-resize.pptx` (19 cases): `spAutoFit`/baked-`cy` calibration: per-font
  core + Aptos anchor/under-fill/spacing/inset sweep.
- `autofit-edge.pptx` (11 cases): over-long unbreakable tokens, trailing spaces,
  empty paragraphs, tabs, whitespace-only runs, mixed sizes, plus a documented
  unsupported RTL box and a single-line CJK box kept from when CJK was a gap.

A fifth deck sits outside that matrix because it calibrates a different axis, and
needs a font none of the four use:

- `autofit-cjk-wrap.pptx` (11 cases, **Malgun Gothic**): where PowerPoint breaks
  **lines** in East Asian text. One `spAutoFit` box per case at a fixed width, so
  the app's own answer is baked into `a:ext/@cy`. Han, Kana, fullwidth Latin,
  halfwidth Katakana and Plane 2 all break per character; Hangul does not, and the
  Hangul box spends a third line where the Han box spends two. Its sidecar,
  `autofit-cjk-wrap.oracle.json`, is written by `authoring/author-cjk-wrap.ps1`
  rather than derived from the deck, because **nothing in the package records where
  a line broke**: the `lines` column is `TextRange.Lines()` read over COM at
  authoring time. The `bakedHeightPt` column *is* in the package, and
  `test/read/cjk-line-breaking-oracle.test.js` re-derives it from the committed deck
  on every run so a hand-edited sidecar stops matching its own fixture.

The decks are the source of truth; `test/read/fixtures/autofit-calibration.json`
(with the LibreOffice cross-measure column) is the derived, regenerable table the
Node-only/Linux test suite reads. The whole regeneration chain lives in
`test/read/fixtures/authoring/`: `gen-cases.mjs` writes the manifests `author-all.ps1`
authors from, `measure-lo.py` beside it adds the LibreOffice column, and
`extract-autofit-calibration.mjs` derives the table from the committed decks. Only
the first two need Windows; the extractor runs anywhere.
Provenance, SHA-256 hashes, and the case-id scheme are in
`test/read/fixtures/README.md`.

### Findings that parameterize the solver

- Single-spacing line pitch ≈ **1.2117 × fontSize**, and **font-independent**
  across the five fonts (the per-font axis is advance **width**, not line height).
- `fontScale` sits on a **2.5% grid** (…, 85, 77.5, 70, 62.5, 55, …, 40, …)%.
- `lnSpcReduction` ramps 0 → 10 → 20% and caps at 20% (recorded for a future
  refinement; the solver currently leaves it at 0).
- Vertical anchor does **not** change `fontScale`.
- On `spAutoFit` growth, `off.y` shifts by **0 / half / full** of the height delta
  for anchor **t / ctr / b**.
- PowerPoint vs LibreOffice autofit height agree to **≤ 0.05pt** for this font set.
- Han, Kana, fullwidth Latin, halfwidth Katakana and Plane 2 ideographs break
  **per character**; **Hangul does not** (it breaks at spaces, like Latin).
- Kinsoku is on: PowerPoint hangs `。`/`、` past the right inset rather than start a
  line with one.

## Standing caveats

- **resize ≠ "extend the card".** Baking `ext.cy` grows only the *text box*. A card
  background rectangle and an adjacent icon are separate shapes the library does not
  know are related, so resize alone will not "extend the card" or un-overlap the
  icon: that layout coordination lives in the consumer's component. This makes
  **shrink** the higher-leverage fix for the actual driver; resize is a partial
  answer for grouped card components.
- Metric fidelity vs PowerPoint's layout engine (kerning, ligatures, GPOS) and vs
  LibreOffice's line metrics is mitigated by erring conservative (raw advances,
  taller-of-two line height) plus the calibration regression target.
- **CJK line breaking is modeled** (see [East Asian line breaking](#east-asian-line-breaking));
  kinsoku, font fallback for uncovered code points, RTL and complex shaping are not.
- Keep `opentype.js` on the Node path / lazy-loaded to bound browser bundle size.
