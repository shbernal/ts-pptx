---
doc-schema-version: 1
title: "Text that fits: design"
summary: "How measured fit is built: the font metrics provider and registry, font collection unwrapping, the wrap simulator and the shrink and resize solvers with the PowerPoint findings behind their constants, East Asian line breaking, the export-time pass, table cells and the estimate for unregistered faces."
read_when:
  - Changing the wrap simulator, the shrink or resize solvers, or their constants
  - Changing font metrics, the registry, or how one font is picked out of a .ttc or .otc collection
  - Touching the export-time fit pass or registerFontMetrics
  - Changing where lines may break in Chinese, Japanese or Korean text
  - Changing how table cells shrink, or how row heights are read for measuring
doc_type: "architecture"
---

# Text that fits: design

The user guide is [Text that fits](../../text-fit.md). This page is how the code behind it is shaped.

## Where the code lives

| Module | Holds |
| --- | --- |
| `src/measure/font-metrics.ts` | `parseFontMetrics`, `FontMetricsRegistry`, the estimate for unregistered faces, `makeRegistryResolver`, `collectUncoveredCodepoints` |
| `src/measure/font-collection.ts` | `listFontFaces`, `resolveFontFace`, `extractFontFace` |
| `src/measure/text-fit.ts` | the tokenizer, the line counter, `measureLayout`, both solvers and their constants |
| `src/measure/paragraphs.ts` | an authored text object turned into `FitParagraph`s, with its insets, box and anchor share |
| `src/measure/table-fit.ts` | the resolved table grid, cell options and insets, `computeTableLayout` |
| `src/measure/fit.ts` | `applyMeasuredFit`, the export-time pass, and `measureText` |
| `src/measure.ts` | the `pptx-ts/measure` entry |
| `src/families/measure.ts` | `measureText`, `overflowsBox` and `tableLayout` on the presentation |
| `src/font-source.ts` | a font source read into bytes, shared with `embedFont` |

## Where the result is baked

A bare autofit element carries no result, and each mode keeps its result in a different place:

| `fit` | Bare element | Where the result goes |
| --- | --- | --- |
| `'shrink'` on a text box | `<a:normAutofit/>` | `fontScale`, and `lnSpcReduction` when set, on the element. An absent attribute means 100% and 0%. |
| `'resize'` on a text box | `<a:spAutoFit/>` | the shape's `a:ext/@cy`, and `a:off/@y` for the anchor |
| `'shrink'` on a table cell | none: `a:tcPr` has no autofit child, and PowerPoint ignores a `normAutofit` in a cell's `txBody` | the run font sizes |

## Font metrics provider

- `parseFontMetrics` imports opentype.js dynamically, so the dependency loads on the first parse and not when `pptx-ts/measure` is imported.
- `advanceWidthPt` sums `charToGlyph(ch).advanceWidth` over the code points and adds `charSpacingPt` once per code point. It does not call `getAdvanceWidth`, which runs GPOS and GSUB shaping and throws on some lookups. Without kerning, a line measures at least as wide as PowerPoint sets it.
- `hasCodepoint` asks `hasChar`. `charToGlyph` returns `.notdef` for a missing code point, so it cannot report coverage.
- Parser errors are rethrown as `MediaError` with `font/parse-failed`.
- `FontMetricsRegistry` keys metrics by lowercased face, bold and italic. `get` falls back from the exact variant to regular, then to any variant of the face, and returns `undefined` for a missing or empty face. `hasCodepoint(face, cp)` returns `undefined` for an unregistered face, which a coverage audit must not read as covered.
- `registerFontMetrics` reads the source through `resolveFontBytes`, where a string is a path or URL. It passes `face` as the font selector only when `isFontCollection` is true. A plain font's internal family name need not match the deck's, and registering one under another name has to keep working.

## Font collections

opentype.js does not read the `ttcf` wrapper: `parseBuffer` throws `Unsupported OpenType signature ttcf`. The gap is tracked upstream as opentypejs/opentype.js#379, with #866 an unmerged implementation. `src/measure/font-collection.ts` unwraps one member into a standalone sfnt before parsing.

- A member's table records hold offsets from the start of the file, which is how members share one `glyf`. A member therefore does not have to be copied out. `extractFontFace` copies the file once and writes the member's table directory over the header at offset 0, leaving every table where it is.
- The directory fits because a well-formed collection places its first table after the header and every member's directory. `extractFontFace` checks this and throws `font/parse-failed` when the directory would overwrite a table.
- `head.checkSumAdjustment` then describes the collection rather than the buffer. Nothing reads it, so it is not recomputed.
- `listFontFaces` reads name IDs 1, 2, 4 and 6 from each member's own `name` table. It prefers Windows English records, then other Windows or Unicode records, then Macintosh ones. `msgothic.ttc` carries a Japanese name for the same IDs, and callers select by the English one.
- `resolveFontFace` never falls back to the first font. The wrong member measures with plausible widths, and nothing later can tell. A plain font is a one-entry list, so the selector is checked there too.

`test/regression/text/font-collection.test.js` has two suites. The first builds a collection from the committed Silkscreen fonts and asserts that a member's metrics equal the same font parsed on its own; it runs everywhere. The second checks the genuine Windows collections against `test/read/fixtures/fonts/windows-collections.oracle.json`, 38 faces across 15 collections, with advances read by WPF's `GlyphTypeface`, which shares no code with this repo. That suite catches a builder and a reader that agree on a layout real files do not use. It skips off Windows.

## Wrap simulator and solvers

`measureLayout(paragraphs, innerWidthPt, resolve, fontScalePct, lnSpcReductionPct, widthSafety)` is the one layout. The export pass, `measureText` and `computeTableLayout` all call it. It returns `null` when any run has no metrics or the width is not above 0.

1. `tokenizeParagraph` resolves each run's metrics, scales its size by `fontScalePct`, and emits tokens: a newline, a space, or a word. A tab counts as one space. A word continues across run boundaries. A code point in the per-character break set, [below](#east-asian-line-breaking), is a word of its own. Every advance is multiplied by `widthSafety`.
2. `countLines` places the tokens greedily, as `wrap=square` does. A space that would overflow ends the line. A word that fits the box but not the rest of the line starts a new line. A word wider than the box wraps character by character. Trailing spaces count toward the line width. It returns the line count and the widest line.
3. A paragraph's line height is `lineSpacingPts` when set, otherwise `SINGLE_LINE_PITCH` times its largest run size times `lineSpacingPct` / 100, then reduced by `lnSpcReductionPct`. The paragraph adds its lines times that height, plus its space before and after.

`solveShrink(paragraphs, box, resolve)`:

- lays the text out with `WIDTH_SAFETY_FACTOR` and compares the height times `HEIGHT_SAFETY_FACTOR` with the inner height. For `wrap: false` the layout width is unbounded, and the widest line must also fit the inner width;
- returns `fits` at 100%, otherwise steps down by `FONT_SCALE_STEP_PCT` to `MIN_FONT_SCALE_PCT` and returns the first scale that fits, or the floor when none does;
- leaves `lnSpcReductionPct` at 0. PowerPoint reduces line spacing to keep a larger font, so a scale found without the reduction is never above PowerPoint's;
- returns `unmeasurable` when a run has no metrics.

`solveResize` lays the text out once at 100%, with an unbounded width for `wrap: false`, and returns the height times `HEIGHT_SAFETY_FACTOR`. A short result has no fallback: the text overflows the baked box. So resize errs tall.

### Findings behind the constants

Measured from the PowerPoint-authored decks listed under [Font oracles](../testing.md#font-oracles). The first eight rows cover Aptos, Aptos SemiBold, Calibri, Tahoma and Arial:

| Finding | Where it is used |
| --- | --- |
| Single-spacing line pitch is 1.2117 times the font size at 12, 18 and 32 pt, the same for all five fonts. The fonts differ in advance width, not in line height. | `SINGLE_LINE_PITCH` |
| PowerPoint and LibreOffice line heights agree within 519 EMU, 0.041 pt, across the 90 line-metric cases. | resize is held against both |
| `fontScale` sits on a 2.5% grid. | `FONT_SCALE_STEP_PCT` |
| `lnSpcReduction` goes from 0 to 10% to 20% and stops at 20%, applied before `fontScale` drops further. | not used: the solver writes 0 |
| The vertical anchor does not change `fontScale`. | the shrink solver ignores the anchor |
| When `spAutoFit` grows a box, `off.y` moves by none, half or all of the height change for anchor `t`, `ctr` and `b`. | `anchorTopShareOfDelta` in `src/measure/paragraphs.ts` |
| Leading and trailing spaces count toward line width. An empty paragraph takes a full line. Line height follows the largest run. | `countLines`, `measureLayout` |
| An unbreakable word wraps by character at `wrap=square`, and widens the box at `wrap=none`. | `countLines`. Resize does not widen a box. |
| Raw advances pack a line slightly tighter than PowerPoint, which rounds each advance at device resolution. Width times 1.03 and height times 1.04 keep the computed scale at or below PowerPoint's on the calibration cases. | `WIDTH_SAFETY_FACTOR`, `HEIGHT_SAFETY_FACTOR` |
| Han, Kana, fullwidth Latin, halfwidth Katakana and Plane 2 ideographs break per character. Hangul does not. PowerPoint applies kinsoku. | `isCjkBreakCharacter`, below |

`MIN_FONT_SCALE_PCT`, 25, is the solver's own floor, not a finding.

## East Asian line breaking

Chinese and Japanese text breaks between any two characters. A tokenizer that treated a run of it as one word would move the whole run to the next line, count too many lines, and shrink text that fits. `isCjkBreakCharacter` makes each such code point its own break opportunity.

| Script | Breaks per character in PowerPoint | Kinsoku in PowerPoint | The model | Cases in `autofit-cjk-wrap.pptx` |
| --- | --- | --- | --- | --- |
| Chinese: Han, Extensions A to I, Bopomofo | yes | not pinned | per character, no kinsoku | `han_between_words`, `han_run_alone`, `han_latin_no_space`, `ext_b_astral` |
| Japanese: Hiragana, Katakana, halfwidth Katakana | yes | yes: `、` hangs past the right inset | per character, breaks before `、` | `kana_between_words`, `halfwidth_kana`, `kinsoku_hanging_comma` |
| Fullwidth Latin | yes | not pinned | per character | `fullwidth_latin` |
| Korean: Hangul syllables and jamo | no, it breaks at spaces like Latin | not applicable | at spaces; a run longer than the line wraps by character | `hangul_between_words`, `hangul_run_alone` |

- The code point ranges are in `isCjkBreakCharacter`. The fixture pins one case per row. The other ranges, such as CJK punctuation, Kanbun, compatibility forms and Extensions C to I, are sibling blocks generalized from those cases.
- Hangul is left out on purpose. UAX #14 allows a break between Hangul syllables, but PowerPoint moves the run down whole, so the Hangul box takes three lines where the Han box takes two. Breaking per syllable would count too few lines, the direction that overflows. A Hangul run longer than a line still wraps, through the over-long word fallback in `countLines`.
- The two Bopomofo blocks sit on either side of Hangul Compatibility Jamo, and they are in the break set. They are Chinese phonetic notation.
- U+3000 IDEOGRAPHIC SPACE is whitespace, so it stays a space token.
- Without kinsoku, the model's line count matches and its widest line is narrower, so the height stays on the safe side. `test/read/cjk-line-breaking-oracle.test.js` pins the gap and fails when kinsoku is implemented.

### Font fallback

PowerPoint draws a code point the named font lacks from a substitute font, at that font's advance. The model has no fallback. It charges the registered font's `.notdef` advance, one flat number unrelated to the glyph that paints, so this is the one approximation with no fixed direction. Malgun Gothic's `.notdef` advances 0.663 em. It lacks halfwidth Katakana, normally half an em wide, so that case gains a line. It also lacks Plane 2 ideographs, normally a full em wide, so 24 of them in a 150 pt box measure 2 lines where full-em advances need 3.

The model reports the gap rather than absorbing it. `collectUncoveredCodepoints` audits each run against its registered face before layout. `measureText` returns the result as `uncoveredCodepoints`, and the export pass warns `measure/uncovered-codepoints`. A face without metrics is not audited, because the estimate has no cmap. The CJK oracle skips the two cases whose glyphs Malgun Gothic lacks, after asserting that the audit reports them. `test/regression/text/measure-text-api.test.js` pins the short direction with synthetic metrics.

## Integration

`src/gen/prepare.ts` orders the steps that run before the synchronous XML build: placeholder backfill, media, then `bakeMeasuredFit`. The bake is last because it measures text the backfill may add. `src/package/assemble.ts` and `src/gen/extract-slides.ts` both call it.

`applyMeasuredFit(slides, registry)`:

- returns at once on an empty registry, because a deck that registered nothing has not asked for measuring;
- walks each slide's objects and descends groups at every depth. `computeBox` reads the object's own `w` and `h`, which is correct because `addGroup` writes an identity child coordinate space. A scaled group would need the ancestor transform the read side composes in `Shape.absoluteFrame`;
- measures text objects whose `fit` is the string `'shrink'` or `'resize'`. The object form is already a result, and is skipped;
- for `'shrink'`, rewrites `options.fit` to `{ type: 'shrink', fontScale }`, adding `lnSpcReduction` only when it is not 0;
- for `'resize'`, rewrites `options.h`, and `options.y` when it moves, as `"<emu>emu"` strings. The emitter keeps `<a:spAutoFit/>`;
- collects the faces it could not measure, the faces it estimated and the uncovered code points across all slides, and warns once per code at the end.

The emitters read only the rewritten options. The rewrite lands on the stored slide model, so a second write finds baked values: a text box's object-form `fit` is skipped, and a resized box or a shrunk cell measures again from its new size and comes out the same.

The rule is that a layout-time prediction must never disagree with what the export then bakes. `measureText` and the pass share `buildFitParagraphs`, `makeRegistryResolver` and `measureLayout`, and `test/regression/text/measure-text-api.test.js` asserts that `measureText`'s height equals `solveResize`'s for the same input. The one case where they differ, a deck with no registered face, is stated in the guide under [Measure text before export](../../text-fit.md#measure-text-before-export), and `test/regression/text/measured-fit-integration.test.js` pins it.

The bake is on the core write path. The `measure` family holds only the three presentation methods, so a deck composed without it still gets its fit.

## Table cells

A cell's `fit: 'shrink'` bakes a smaller literal font size, which PowerPoint and LibreOffice draw with no edit.

- `resolveTableGridEmu` resolves the column count, column widths through `resolveTableColWidthsEmu`, and row heights through `resolveTableRowHeightEmu`, from `h` or the resolved `cy`. The emitter and `pptx.tableLayout()` use the same resolver.
- `walkTableGrid` yields each origin cell with its spans. A cell is measured across the columns and rows it covers, and skipped when any row it covers has no fixed height.
- The cell's own `fit` wins, otherwise a table-level `'shrink'` applies. Only the string `'shrink'` is acted on.
- `effectiveCellOpts` fills the cell's unset text options from the table through `CELL_INHERITED_TEXT_KEYS`, the emitter's own list. `resolveCellInsetsEmu` resolves margins the way the emitter does.
- `solveShrink` runs as for a text box. Below 100%, `scaleCellFontSizes` multiplies the cell's size and every run's explicit size by the scale and floors each to 0.1 pt. It assigns new option objects rather than editing objects that plain-string cells share with the table.
- `'resize'` and the object form are ignored, because a row with no fixed height already grows.
- `resolveTableRowHeightEmu` in `src/units-internal.ts` is the one reading of `rowH`. An entry fixes its row only when it is a number above 0. Any other present value warns `table/invalid-row-height` and falls back to an even split of `h`. A missing slot or `null` is silent, since that is how the auto-pager marks a row that grows. `test/regression/table/table-row-height-agreement.test.js` checks that the emitted `<a:tr h>`, `computeTableLayout` and the fit pass agree.
- `computeTableLayout` estimates a row with no fixed height with the same layout at full size and flags it `heightExact: false`. A cell with a rowspan does not drive a row's estimate.

## Estimate for unregistered faces

`getHeuristicFontMetrics()` returns one shared `FontMetrics` with no font behind it. `heuristicCharRatio` gives each character a fixed advance, biased wide for a proportional Latin face:

| Characters | Advance |
| --- | --- |
| U+2E80 and above | 1.0 em |
| U+0100 to U+2E7F | 0.62 em |
| space, tab and narrow characters such as `i`, `l` and `.` | 0.3 em |
| semi-narrow characters such as `f`, `t`, `r` and brackets | 0.4 em |
| `m`, `w` / `M`, `W` / `@`, `%` | 0.9 / 0.98 / 1.0 em |
| other capitals A to Z | 0.72 em |
| everything else: lowercase, digits, other punctuation | 0.58 em |

- `hasCodepoint` always returns `true`, so the estimate never reports a missing glyph. The coverage audit skips estimated faces through the registry instead.
- `makeRegistryResolver` returns exact metrics, then the estimate for a named face, calling `onHeuristic` with that face, and `undefined` for a run with no face. The export pass uses it only when the registry is not empty.

## Calibration

The oracles that hold these constants against PowerPoint, the decks behind them, and how to regenerate them are under [Font oracles](../testing.md#font-oracles). Fixture provenance, hashes and case ids are in `test/read/fixtures/README.md`.
