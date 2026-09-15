---
doc-schema-version: 1
title: "Tables: design"
summary: "Why addTable() writes what it writes where PowerPoint's own output or behavior decided it: built-in style GUIDs only and the bare tableStyles.xml stub, covered merge cells that copy their origin, a pager that never breaks inside a word, and the cell constructs measured and found not worth an emitter."
read_when:
  - Changing how a table style is written, or proposing a custom table style
  - Changing what a covered merge cell carries
  - Changing how the auto-pager breaks cell text into lines
  - Proposing a per-cell no-wrap, a header association or a table-level effect
doc_type: "architecture"
---

# Tables: design

The user guide is [Tables](../../tables.md). This page records the decisions behind it that were settled against PowerPoint, so they are not re-derived.

## Where the code lives

| Concern | File |
| --- | --- |
| the table-styles part | `src/gen/pres/table-styles.ts` |
| the rectangular grid, covered cells and `a:tcPr` | `src/gen/slide/objects/table.ts` |
| line breaking and auto-paging | `src/gen/table/autopage.ts` |
| the table-cell probes | `test/read/fixtures/authoring/probe-table-cell-wrap.ps1`, `test/read/fixtures/authoring/probe-table-cell-a11y-and-3d.ps1` |

## Only a built-in style GUID renders

PowerPoint resolves `<a:tableStyleId>` against its own table-style gallery. It never reads a style definition out of the package. A GUID it recognises paints even when the deck defines nothing, and a GUID it does not recognise paints nothing, however complete the definition. The table falls back to PowerPoint's no-style look, a black hairline grid on white.

This was measured by rendering in PowerPoint desktop 16.0, not inferred from the schema:

| How the style is offered | Built-in GUID | Custom GUID |
| --- | --- | --- |
| an `<a:tableStyleId>` reference | renders | never |
| an inline `<a:tableStyle>` in `<a:tblPr>` | not tested | never |
| the `def` default on `tableStyles.xml` | never | never |

The method is a rendered pair. Take a PowerPoint-authored deck and rewrite one style's GUID to a novel value in both the styles part and the slide, bytes otherwise identical. That table drops to the black grid while its untouched neighbours keep their styling. Lifting a genuine PowerPoint-authored `<a:tblStyle>` block under a custom GUID does not help either, so the markup was never the problem. PowerPoint also has no command to create a table style, unlike Word and Excel.

What follows from it:

- `ppt/tableStyles.xml` ships as a bare stub that names a default style id and defines nothing. The part still ships because PowerPoint expects its relationship and content-type override.
- `Presentation.defineTableStyle()` and `TableProps.styleDrivenCells` were removed. They emitted well-formed, schema-valid markup that never painted, and `styleDrivenCells` also stood down the direct formatting that was carrying the render.
- The read side still resolves style graphs out of imported decks: `Table.resolvedStyle`, `TableCell.resolvedFill` and `importSlideMasters({ tableStyles })` work against the definitions PowerPoint itself wrote.

## Covered merge cells copy their origin

The emitter expands lopsided authored rows into the rectangular grid OOXML requires, inserting each covered cell with its `hMerge` or `vMerge` flag. A covered cell carries the origin's border and fill. PowerPoint writes a bare `<a:tcPr/>` there instead.

The divergence is deliberate. A covered cell never renders, because the origin spans over it, so the copied fill is invisible either way. Copying keeps every cell on one code path rather than adding a branch that changes nothing on screen, and it puts the merged region's outer edges on the covered cells, which is where PowerPoint reads them.

## The pager never breaks inside a word

The auto-pager fills a line word by word and starts a new one when the next word will not fit. A word wider than its column is placed whole and overflows the column in PowerPoint.

A hard break at the column width, as a browser's `overflow-wrap: break-word` does, was weighed and rejected. It would break against the same characters-per-line estimate every row height here is priced with, so it would land mid-word at a position the renderer disagrees with. The broken text would then have to keep the line count and the emitted text in step, or the row is priced for lines it does not have. Overflowing is predictable and visible, and the caller can widen the column, lower the font size or insert the break with `breakLine`.

## Measured and not authorable

| Construct | Why there is no emitter |
| --- | --- |
| Per-cell no-wrap | PowerPoint has none. `TextFrame.WordWrap` is read-only on a cell over COM, and `<a:bodyPr wrap="none"/>` in a cell renders inert and is stripped on the next save. Probe: `probe-table-cell-wrap.ps1`. |
| `a:tc/@id` and `a:tcPr/a:headers` | The screen-reader header association. PowerPoint opens a deck carrying both without complaint and strips them on the first save. `TableCell.id` and `TableCell.headerIds` read them from other producers' decks, and `hasHeader` is the header marker PowerPoint keeps. Probe: `probe-table-cell-a11y-and-3d.ps1`. |
| Table-level effects on `a:tblPr` | Schema-legal, but PowerPoint's UI has no table-level effect, so no source deck carries one. |

`a:tcPr/a:cell3D` is the control that makes the header result trustworthy. The same probe injected it into the same `a:tcPr`, and PowerPoint kept it verbatim while it discarded `a:headers`. Stripping the header association is a deliberate normalization, not a failed patch, and `cell3D` is authorable through `TableCellProps.cell3D`.
