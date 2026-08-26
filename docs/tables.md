---
doc-schema-version: 1
title: "Tables"
summary: "The table cell model, the styling precedence chain, borders (per-cell, perimeter, dash styles, diagonals), merges, auto-paging, measured fit, and the constructs PowerPoint will not keep."
read_when:
  - Adding a table with addTable() and deciding how to style its cells
  - Working out why a border landed on every cell instead of the table's outside edge
  - Merging cells, or debugging a colspan/rowspan that renders wrong
  - Reading or editing a table in an existing deck through ts-pptx/read
  - Checking whether a table construct is authorable before implementing it
doc_type: "guide"
---

# Tables

```js
const s = pptx.addSlide()
s.addTable(
  [
    ['Region', 'Q1', 'Q2'],
    ['North', '120', '145'],
    ['South', '98', '110'],
  ],
  { x: 1, y: 1, w: 8, hasHeader: true }
)
```

`slide.addTable(rows, options)` takes a **row-major array of rows**, each an array of
cells, and emits a `<p:graphicFrame>` wrapping an `<a:tbl>`. A table is an ordinary slide
object: it takes `x`/`y`/`w`/`h`, an `objectName`, an `altText`, and a `placeholder`.

## The cell model

A cell is a string, a number, or a `TableCell` object:

```js
'plain'                                    // shorthand for { text: 'plain' }
{ text: 'styled', options: { bold: true } }
{ text: [{ text: 'mixed ' }, { text: 'runs', options: { bold: true } }] }
```

`options` is a `TableCellProps`, which extends `TextBaseProps`, so everything that styles
text (`bold`, `color`, `fontFace`, `fontSize`, `align`, `valign`, `textDirection`, …) works
on a cell, alongside the table-specific `fill`, `border`, `diagonal`, `margin`, `colspan`,
`rowspan`, `anchorCtr`, `horzOverflow`, `cell3D` and `fit`.

Rows may be **lopsided**. A row under a `rowspan` from above holds only the cells that
actually start in it, and the emitter builds the rectangular merge grid itself. This is the
one place authoring and reading differ in shape: see [Merging cells](#merging-cells).

## Styling precedence

Highest wins, matching how PowerPoint resolves styling: direct cell formatting always beats
a style region:

| # | Source | Applied |
| --- | --- | --- |
| 1 | the cell's own `options` | as authored |
| 2 | `headerRow` | to every cell of row 0 |
| 3 | `columns[i]` | to every cell starting in column `i` |
| 4 | table-level options | to every cell that set none |
| 5 | **library defaults** | stamped onto every cell as direct formatting |
| 6 | `tableStyle` | by the built-in style's own region rules |

Exactly eleven table-level options inherit down to cells (row 4): `align`, `bold`, `border`,
`color`, `fill`, `fontFace`, `fontSize`, `margin`, `textDirection`, `underline`, `valign`.
The list is closed: `italic`, for instance, is **not** on it, so a table-level `italic: true`
styles nothing.

### The defaults tier

Row 5 is the one that surprises people, so it is worth stating plainly: for four properties
the library stamps a default onto **every cell as direct formatting**.

| Property | Default stamped on every cell |
| --- | --- |
| `border` | `{ type: 'none' }` on all four sides |
| `color` | `'000000'` |
| `fontSize` | `12` |
| `margin` | `[0.05, 0.1, 0.05, 0.1]` in |

This is deliberate rather than accidental. Direct formatting outranks a style region, so the
explicit four-side `none` is what keeps an unstyled table free of grid lines, and it is the
tier every option above it overrides. Set any of the four on the table, on `headerRow`, on
`columns[i]`, or on a cell, and yours wins.

One carve-out: a **hyperlink anywhere in the grid** stands the `color` default down for the
whole table. Black paints the whole run, so without this the words *after* a link would come
out black instead of following the link colour. Text then falls through to the theme's `tx1`.

### How to brand a table

Direct formatting is the mechanism: `headerRow`, `columns[i]`, the table-level `border` /
`fill` / `color`, and per-cell `options`:

```js
s.addTable(rows, {
  hasHeader: true,
  border: { type: 'solid', color: 'D9D9D9', width: 0.5 },
  headerRow: { fill: { color: '1A2B3C' }, color: 'FFFFFF', bold: true },
})
```

For banded rows, set each row's `fill` as you build the data: a table style cannot supply
brand colours (see [Table styles](#table-styles) below for why).

The merge is **property-level**, not object-level. A header cell keeps `headerRow`'s
typography *and* takes its column's `fill` when the two set different properties: which is
what makes a graduated header band expressible without writing a fill onto every cell:

```js
s.addTable(rows, {
  headerRow: { color: 'FFFFFF', bold: true, align: 'center' },   // no fill here
  columns: [{}, { fill: { color: 'BBD3FB' } }, { fill: { color: '4B7BE5' } }],
})
```

`columns[i]` counts each cell's `colspan` within a row, so merged cells map to the right
column. It does not track a `rowspan` inherited from an earlier row.

Setting `headerRow` implies `hasHeader: true` unless you set `hasHeader` explicitly.

### Table styles

`tableStyle` takes a built-in `TableStyle` member. Which regions activate is controlled by the
flags: `hasHeader` → `firstRow`, `hasFooter` → `lastRow`, `hasBandedRows` → `band1H`/`band2H`,
and so on.

```js
s.addTable(rows, { tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1, hasHeader: true })
```

`hasHeader` does one thing beyond activating a region: it emits the `firstRow="1"` marker
PowerPoint's accessibility checker reads, so it is worth setting on any table with a header
row regardless of styling.

#### Only a built-in style renders, and there is no custom-style escape hatch

**PowerPoint resolves `<a:tableStyleId>` against its own table-style gallery. It never reads a
style definition out of the package.** A GUID it recognises paints even when the deck defines
nothing; a GUID it does not recognise paints nothing however complete the definition: the
table falls back to PowerPoint's no-style look, a black hairline grid on white.

This was measured by rendering, not inferred from the schema (PowerPoint desktop 16.0):

| how the style is offered | built-in GUID | custom GUID |
| --- | --- | --- |
| `<a:tableStyleId>` reference | **renders** | never |
| inline `<a:tableStyle>` in `<a:tblPr>` | not tested | never |
| `def=` default on `tableStyles.xml` | never | never |

The decisive control: take a PowerPoint-authored deck and rewrite one style's GUID to a novel
value in *both* the styles part and the slide, bytes otherwise identical. That table drops to
the black grid while its untouched neighbours keep their styling. Lifting a genuine
PowerPoint-authored `<a:tblStyle>` block under a custom GUID does not help either: the markup
was never the problem. (This is consistent with PowerPoint having no "New Table Style" command,
unlike Word and Excel.)

So `ppt/tableStyles.xml` ships as a bare stub naming a default style id and defining nothing.
`Presentation.defineTableStyle()` and `TableProps.styleDrivenCells` were **removed**: they
emitted well-formed, schema-valid, permanently invisible markup, and `styleDrivenCells` made it
worse by standing down the direct formatting that was carrying the render. Use
[direct formatting](#how-to-brand-a-table) for brand colours.

The read side is unaffected and still resolves style graphs out of *imported* decks:
`Table.resolvedStyle`, `TableCell.resolvedFill`, and `importSlideMasters({ tableStyles })` all
work against the definitions PowerPoint itself wrote.

## Fills

Two table-level fill options, and the difference is where the paint lands rather than how it
looks:

- **`fill`** is *stamped onto every cell*. Each cell that sets none of its own gets this as
  its own `a:tcPr` fill.
- **`tableFill`** is the table's *own* background: one `a:tblPr` fill that the cells sit on
  top of.

They usually render alike. The difference matters when a cell is meant to be transparent:
with `tableFill`, a cell with no fill shows the background through; with `fill` there is no
such thing as a cell without a fill, so nothing can fall back to it. `tableFill` is also
what a deck read back from PowerPoint actually carries, so it is the right choice when
reproducing a source deck.

A cell fill is a `ShapeFillProps`, so it takes the whole DrawingML fill group: solid,
gradient, pattern, or a picture:

```js
{ fill: { color: '0088CC', transparency: 50 } }
{ fill: { color: SchemeColor.accent1 } }
{ fill: { type: 'gradient', gradient: { kind: 'linear', angle: 90, stops: [...] } } }
{ fill: { type: 'pattern', pattern: { preset: 'diagCross', fgColor: '1A2B3C' } } }
{ fill: { type: 'image', image: { path: 'logo.png' } } }
```

A picture fill is embedded once and shared by every cell that inherits it. Raster only: an
SVG warns and is ignored, matching shape fills.

## Borders

### `border` is a per-cell default, not the table's perimeter

This is the single most misread option in the table surface. `TableProps.border` is
broadcast to **every cell**:

```js
// Every cell gets a top and bottom rule -- i.e. a full set of horizontal grid lines,
// NOT a rule above and below the table.
s.addTable(rows, { border: [{ type: 'solid' }, { type: 'none' }, { type: 'solid' }, { type: 'none' }] })
```

An array is read in **TRBL** order (`[top, right, bottom, left]`); a single `BorderProps` is
broadcast to all four sides. A cell's own `options.border` overrides the table default
entirely: the two do not merge per side.

A `null` entry is a **hole**, not a rule. The element is left out of `<a:tcPr>` altogether, so
that edge keeps inheriting from the built-in table style, the theme banding and the master
chain. `{ type: 'none' }` is the other state: an explicit "no line" that overrides the
inheritance. So `[rule, null, rule, null]` draws horizontal rules and leaves the style's
vertical ones alone, while `[rule, { type: 'none' }, rule, { type: 'none' }]` draws the same
two and erases the other two. A cell with no `border` authored at all is the third case, and
it depends on `tableStyle`: a table that named a style leaves every edge absent so the style
paints its own grid, while a table with no style takes the four-side no-fill default, which is
what keeps it free of PowerPoint's no-style black hairline grid. To erase a styled table's
grid, say so: `border: { type: 'none' }`.

### `outerBorder` is the perimeter

For the table's outside edge (the top of the first row, the bottom of the last, the left of
the first column and the right of the last) use `outerBorder`:

```js
s.addTable(rows, { outerBorder: { type: 'solid', color: '1A2B3C', width: 1 } })   // box it
s.addTable(rows, { outerBorder: [rule, undefined, rule, undefined] })              // rules above and below
```

An omitted entry leaves that side to whatever `border` (or the cell's own option) already
drew, so the two compose: `border` draws the interior grid and `outerBorder` overrides the
edges it reaches. "Outline the table, no interior grid" is `outerBorder` with no `border`.

The perimeter is decided by **grid position**, so merges work: PowerPoint defines a merged
region's outer edges on the *covered* cells, and a colspan reaching the last column gets that
column's rule on its covered half.

### Dash styles

`BorderProps.type` is a coarse three-way switch (`'none' | 'solid' | 'dash'`). For a specific
dash, set `dashType`, which takes the whole `ST_PresetLineDashVal` set:

```js
{ border: { type: 'solid', color: '999999', dashType: 'lgDashDot' } }
```

`dashType` wins over `type` when both are set, except that `type: 'none'` suppresses the
border before any dash is chosen. An unrecognized value is reported as
`border/invalid-dash-type` and falls back to what `type` implies: a value outside the enum
would make the slide part schema-invalid, which PowerPoint reports as a corrupt file rather
than as a mis-set option.

### Diagonals

The two corner-to-corner rules (PowerPoint's "Diagonal Down/Up Border") are a separate
option, not two more entries on `border`'s tuple:

```js
{ diagonal: { tlToBr: { type: 'solid', color: 'C00000' } } }              // strike a cell out
{ diagonal: { tlToBr: { type: 'solid' }, blToTr: { type: 'solid' } } }    // an X
```

On a merged cell the diagonal is drawn **once**, on the span origin: it is a single stroke
across the whole region, and repeating it per covered cell would draw a sawtooth. Covered
cells inherit the origin's *edges* but never its diagonals.

## Merging cells

`colspan` and `rowspan` go on the cell that starts the span:

```js
s.addTable([
  [{ text: 'wide', options: { colspan: 3 } }],
  [{ text: 'tall', options: { rowspan: 2 } }, 'B2', 'C2'],
  ['B3', 'C3'],                                   // no cell for column 0 — the rowspan covers it
], { x: 1, y: 1, w: 9 })
```

Rows are authored **lopsided** (a row covered by a `rowspan` from above simply omits that
cell) and the emitter expands them into the rectangular grid OOXML requires, inserting the
covered cells with their `hMerge`/`vMerge` flags.

A covered cell inherits the origin's border and fill. That is a deliberate divergence from
PowerPoint, which writes a bare `<a:tcPr/>` there: a covered cell is never rendered (the
origin spans over it), so copying the fill is invisible either way, and keeping it uniform
avoids a branch that would change nothing on screen. It also puts the merged region's outer
edges where PowerPoint expects to find them.

## Sizing and overflow

**Columns.** `colW` takes inches, either one number for every column or an array. Without it
the table's `w` is split evenly. `fitColumns: 'shrink'` scales every column down by the same
factor when the total exceeds the space available from `x`; it never grows a column and
enforces no minimum width, so a very high column count can still end up thin.

**Rows.** `rowH` takes inches, one number or an array. A row is auto-height (as tall as its
content needs) only when the table sets **neither `rowH` nor `h`**. A table-level `h` is
divided evenly across the rows, so `h` makes every row fixed just as surely as `rowH` does.

**Text that does not fit.** `fit: 'shrink'`, on a cell or on the table, measures the wrapped
text and bakes a reduced literal font size onto the runs. It requires the font registered via
`registerFontMetrics` and only triggers when a row's height is **fixed** and the text exceeds
it: which, per the previous paragraph, means a table with `rowH` *or* `h`. With neither, the
row simply grows and nothing shrinks. See [Measured text fit](measured-text-fit.md).

**Cell text always wraps.** PowerPoint has no per-cell no-wrap: `wrap="none"` on a cell's
`a:bodyPr` renders inert and is stripped on the next save. `horzOverflow` is *not* that
switch: it decides whether a single glyph too wide for the line is clipped at the cell edge
or draws past it, which matters for oversized display type and wide CJK/emoji glyphs.

**Auto-paging.** `autoPage: true` shreds a table too tall for one slide across as many as it
needs. `autoPageRepeatHeader` (with `autoPageHeaderRows`) repeats the header on each,
`autoPagePlaceholder` copies the source slide's populated placeholders onto the overflow
slides, and `autoPageSlideStartY` sets where the continuation starts.

How much fits is an **estimate**, not a measurement: the pager counts wrapped lines from a
characters-per-line approximation and prices each at the font size times a fixed line-height
factor, plus the cell's top and bottom margins. It never asks a renderer. So the page break
lands where the arithmetic says, and a font whose real metrics are far from the approximation
can put it a row out: `autoPageCharWeight` (characters) and `autoPageLineWeight` (line
height) nudge the two constants when it does. What the pager does guarantee is
self-consistency: pages of equal usable height get equal row budgets. Continuation slides
start at `autoPageSlideStartY`, or at the top margin when it is unset, so a first page
placed lower with `y` is the one page that legitimately holds fewer rows.

## Reading and editing an existing table

`ts-pptx/read` exposes `Table → TableRow[] → TableCell[]`, each wrapping a live DOM element.
Reading covers the cell model, the six borders, both fills, the spans and the style graph;
`TableCell.resolvedFill` reports the colour a cell *renders* as, folding in the style's
banding, and `TableCell.hasOwnFill` says whether that colour is the cell's own: which is
the distinction anything reproducing a table needs, since baking an inherited banding colour
into a copy makes it stop responding to its own style.

Editing covers cell properties (`setAnchor`, `setVerticalText`, `setHorzOverflow`,
`setAnchorCtr`, `setMarginsEmu`, `setBorder`, `setFillColor`, `setFillSchemeColor`,
`noFill`) and structure (`addRow`, `removeRow`, `addColumn`, `removeColumn`, `mergeCells`,
`unmergeCell`). Each mutates in place and marks the part dirty.

Unlike the write path, an invalid value **throws** here rather than warning and being
dropped: a caller editing one attribute would otherwise be left looking at an unchanged deck
with nothing to explain it.

A **stored** table's grid is already rectangular: unlike the authoring side, every `a:tr`
holds one `a:tc` per column with the covered halves present and flagged. The structural
editors keep it that way: inserting a row or column *through* a merge extends it rather than
splitting it, removing a merge origin promotes its first continuation so the region survives
one shorter, and `mergeCells` rejects a rectangle that cuts through an existing merge rather
than silently widening it to fit.

## Not authorable

Constructs measured against desktop PowerPoint and found not to be worth an emitter. Each is
recorded so it is not re-attempted.

| Construct | Why |
| --- | --- |
| **Per-cell no-wrap** | PowerPoint has none. `TextFrame.WordWrap` is read-only on a cell over COM, and `<a:bodyPr wrap="none"/>` in a cell renders inert and is stripped on the next save. Probe: `test/read/fixtures/authoring/probe-table-cell-wrap.ps1`. |
| **`a:tc/@id` and `a:tcPr/a:headers`** | The screen-reader header association. PowerPoint opens a deck carrying both without complaint and then **strips them on the first save**, so an emitter would ship a feature that dies as soon as anyone edits the deck. Read accessors (`TableCell.id`, `.headerIds`) exist for decks from other producers. `hasHeader` is the header marker PowerPoint keeps, and the one its own accessibility checker reads. Probe: `test/read/fixtures/authoring/probe-table-cell-a11y-and-3d.ps1`. |
| **Table-level effects** (`a:tblPr` `EG_EffectProperties`) | Schema-legal, but PowerPoint's UI exposes no table-level effect, so a source deck will not contain one and there is nothing to reproduce. |

`a:tcPr/a:cell3D` is the counter-example that makes the header-association result
trustworthy rather than circumstantial: it was injected into the **same** `a:tcPr` by the
same probe, and PowerPoint preserved it verbatim while discarding `a:headers`. So it is a
deliberate normalization, not a failed patch, and `cell3D` is authorable, via
`TableCellProps.cell3D`.
