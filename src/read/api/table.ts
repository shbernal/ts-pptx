/**
 * The `Table` read/edit proxy (`a:tbl`) hosted in a `p:graphicFrame`, and the barrel the rest
 * of the tree imports the table surface through.
 *
 * `Table -> TableRow[] -> TableCell[]`, each wrapping a live DOM element and holding the
 * owning slide `Part` so edits mutate the node in place and call `part.markDirty()`. Cell text
 * reuses the `TextFrame`/`Paragraph`/`Run` proxies, so per-run formatting edits work exactly
 * as on a shape's text.
 *
 * The three classes share no state, so they are three files; this one re-exports the other two
 * and the cell value types, keeping `read.ts`'s import surface exactly as it was.
 */
import type { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { attr, boolAttr, type Element, firstChild, getElements, numberValue } from '../oxml/dom.js'
import { FILL_CHOICES, solidFillColor } from '../oxml/fill.js'
import { insertColumn, insertRow, mergeCells, removeColumn, removeRow, rowsOf, unmergeCell } from './table-structure.js'
import type { ThemeContext } from '../oxml/theme.js'
import { readPictureFill, type PictureFill } from './picture-fill.js'
import { readGradientFill, type GradientFill } from './gradient.js'
import { readPatternFill, type PatternFill } from './pattern-fill.js'
import { resolveTableStyle, type ResolvedTableStyle } from './table-style-resolve.js'
import { resolveSolidFillColor, type ResolvedColor } from './theme-context.js'
import { TableRow } from './table-row.js'
import { TableCell, type TableCellStyleContext } from './table-cell.js'
/** A table: a grid of rows and cells inside a graphic frame. */
export class Table {
	constructor(
		private readonly tbl: Element,
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to each cell's text for `Run.resolvedColor`. */
		private readonly themeContext?: ThemeContext,
		/** The deck package, for resolving `a:tableStyleId` against `tableStyles.xml` (style-graph cell fills). */
		private readonly opc?: OpcPackage,
		/** The owning slide's relationships, for resolving a cell picture fill's `r:embed` to a partname. */
		private readonly rels?: Relationships
	) {}

	/** The table's rows (`a:tr`) in document (top-to-bottom) order. */
	get rows(): TableRow[] {
		const style = this.#styleContext()
		return getElements(this.tbl, 'a:tr').map(
			(tr, rowIndex) => new TableRow(tr, this.part, this.themeContext, style, rowIndex, this.rels)
		)
	}

	/**
	 * The table's style graph entry (`styleId`, `name`, raw `a:tblStyle`) resolved
	 * from the deck's `tableStyles.xml`, or `null` when the table names no
	 * `a:tableStyleId`, the deck has no `tableStyles.xml`, or the id is a built-in
	 * `[MS-OE376]` style the deck does not materialise. This is what supplies the
	 * banded-row / header shading a cell with no own fill inherits — read the
	 * resolved per-cell colour off {@link TableCell.resolvedFill}.
	 */
	get resolvedStyle(): ResolvedTableStyle | null {
		return this.opc ? resolveTableStyle(this.opc, this.styleId) : null
	}

	/**
	 * The table's entry in the deck's `tableStyles.xml`, resolved once per instance.
	 *
	 * Resolving it means scanning the package for the table-styles part and then its
	 * `a:tblStyle` list, and every cell needs the answer -- so reading a 20x8 table cell by
	 * cell did it 160 times for one unchanging answer. The memo is scoped to this instance,
	 * which is rebuilt on every `GraphicFrame.table` access, so it cannot outlive the deck
	 * state it was resolved against.
	 *
	 * A flag rather than a `??=` on a nullable field: "this table names no style" is the
	 * common answer and the expensive one to recompute, and `??=` would memoize every answer
	 * except that one.
	 */
	#resolvedStyle: ResolvedTableStyle | null = null
	#styleAsked = false

	/** The per-cell style-resolution context, or `null` when no style resolves or there is no theme context. */
	#styleContext(): TableCellStyleContext | null {
		if (!this.opc || !this.themeContext) return null
		if (!this.#styleAsked) {
			this.#resolvedStyle = resolveTableStyle(this.opc, this.styleId)
			this.#styleAsked = true
		}
		const style = this.#resolvedStyle
		if (!style) return null
		return {
			style,
			flags: {
				firstRow: this.#tblPrFlag('firstRow'),
				lastRow: this.#tblPrFlag('lastRow'),
				firstCol: this.#tblPrFlag('firstCol'),
				lastCol: this.#tblPrFlag('lastCol'),
				bandRow: this.#tblPrFlag('bandRow'),
				bandCol: this.#tblPrFlag('bandCol'),
			},
			rowCount: this.rowCount,
			colCount: this.columnCount,
			ctx: this.themeContext,
		}
	}

	/** Number of rows (`a:tr`). */
	get rowCount(): number {
		return getElements(this.tbl, 'a:tr').length
	}

	/** Number of grid columns (`a:tblGrid/a:gridCol`). */
	get columnCount(): number {
		const grid = firstChild(this.tbl, 'a:tblGrid')
		return grid ? getElements(grid, 'a:gridCol').length : 0
	}

	/** Column widths in EMU (`a:gridCol/@w`), one per grid column. */
	get columnWidths(): (number | null)[] {
		const grid = firstChild(this.tbl, 'a:tblGrid')
		if (!grid) return []
		return getElements(grid, 'a:gridCol').map((col) => numberValue(attr(col, 'w')))
	}

	/**
	 * The table's style GUID (`a:tblPr/a:tableStyleId` text), e.g.
	 * `{5940675A-B579-460E-94D1-54222C63F5DA}`, or `null` when the table carries no
	 * `a:tableStyleId`. This is the reference into `ppt/tableStyles.xml` (or a
	 * built-in `[MS-OE376]` style) that supplies the banded-row / header shading the
	 * `firstRow`/`bandRow` flags activate — without it a replica loses the whole
	 * table style, so it is the read counterpart of the writer's `tableStyle` option.
	 */
	get styleId(): string | null {
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		const idEl = tblPr && firstChild(tblPr, 'a:tableStyleId')
		const id = idEl?.textContent?.trim()
		return id ? id : null
	}

	/**
	 * The table's own background (`a:tblPr` fill) resolved against the slide's theme colour
	 * context to a literal hex, or `null` when the table carries no solid background.
	 *
	 * This is the table-level counterpart of {@link TableCell.resolvedFill}, and it is a
	 * genuinely different thing from a cell fill: a `a:tblPr` fill sits *behind* the grid, so
	 * a cell with no fill of its own shows it through. Reports `null` for a non-solid choice
	 * (`a:blipFill`/`a:gradFill`/`a:pattFill`/`a:noFill`) — read {@link pictureFill} for an
	 * image background — and `null` with no theme context.
	 */
	get resolvedFill(): ResolvedColor | null {
		if (!this.themeContext) return null
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		if (!tblPr || !FILL_CHOICES.some((q) => firstChild(tblPr, q))) return null
		return resolveSolidFillColor(tblPr, this.themeContext)
	}

	/**
	 * The table's picture background (`a:tblPr/a:blipFill`), or `null` when the table is not
	 * image-backed. The table-level twin of {@link TableCell.pictureFill}, and needed for the
	 * same reason: {@link resolvedFill} decodes only solid colours, so without this an
	 * image-backed table is indistinguishable from an unfilled one.
	 */
	get pictureFill(): PictureFill | null {
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		return tblPr ? readPictureFill(tblPr, this.rels ?? null) : null
	}

	/**
	 * The table's gradient background (`a:tblPr/a:gradFill`), or `null` when it is not
	 * gradient-filled. Needed for the same reason {@link pictureFill} is:
	 * {@link resolvedFill} decodes only solid colours.
	 */
	get gradientFill(): GradientFill | null {
		if (!this.themeContext) return null
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		return tblPr ? readGradientFill(tblPr, this.themeContext) : null
	}

	/**
	 * The table's pattern (hatch) background (`a:tblPr/a:pattFill`), or `null` when it is not
	 * pattern-filled.
	 */
	get patternFill(): PatternFill | null {
		if (!this.themeContext) return null
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		return tblPr ? readPatternFill(tblPr, this.themeContext) : null
	}

	/**
	 * The raw `schemeClr` token of the table's own background
	 * (`a:tblPr/a:solidFill/a:schemeClr/@val`), or `null` for an absent or `srgbClr` fill.
	 * {@link resolvedFill} is the literal it resolves to; this is the unresolved reference,
	 * which is what a replica should carry so the copy still tracks its theme.
	 */
	get fillSchemeColor(): string | null {
		return solidFillColor(firstChild(this.tbl, 'a:tblPr'), 'a:schemeClr')
	}

	/** Whether the first row is styled as a header (`a:tblPr/@firstRow`). */
	get firstRowHeader(): boolean {
		return this.#tblPrFlag('firstRow')
	}

	/** Whether rows are banded (`a:tblPr/@bandRow`). */
	get bandedRows(): boolean {
		return this.#tblPrFlag('bandRow')
	}

	/**
	 * The cell at `(rowIndex, columnIndex)` (both zero-based), or `null` when out
	 * of range. Column index counts `a:tc` elements in the row, so a cell that
	 * spans columns (`gridSpan`) occupies a single index here.
	 */
	cell(rowIndex: number, columnIndex: number): TableCell | null {
		// Indexed rather than `this.rows[r]?.cells[c]`: that built a `TableRow` for every row
		// and a `TableCell` for every cell of the one row asked for, so the natural
		// `for (r) for (c) cell(r, c)` loop over a 20x8 table allocated 3,000 objects and
		// resolved the table's style 160 times to hand back 160 cells.
		const tr = getElements(this.tbl, 'a:tr')[rowIndex]
		const tc = tr && getElements(tr, 'a:tc')[columnIndex]
		if (!tc) return null
		return new TableCell(tc, this.part, this.themeContext, this.#styleContext(), rowIndex, columnIndex, this.rels)
	}

	/**
	 * Insert a row, defaulting to the bottom of the table. Returns the new {@link TableRow}.
	 *
	 * The new row is auto-height (`a:tr/@h="0"`) and its cells are empty. Inserting *through*
	 * a vertical merge extends that merge rather than interrupting it: the span's origin
	 * grows by a row and the new cell joins as a continuation, because an origin claiming
	 * more rows than it has continuations is a table PowerPoint reports as corrupt.
	 * @param {number} [index] - where to insert, `0`..`rowCount`; appended when omitted
	 * @throws {InvalidOptionError} when `index` is outside that range
	 */
	addRow(index?: number): TableRow {
		const tr = insertRow(this.tbl, index)
		this.part.markDirty()
		return new TableRow(tr, this.part, this.themeContext, this.#styleContext(), rowsOf(this.tbl).indexOf(tr), this.rels)
	}

	/**
	 * Remove the row at `index`, with its content.
	 *
	 * A cell in the row that *continues* a vertical merge shortens that merge. A cell that
	 * *starts* one hands the region to its first continuation, which becomes the new origin —
	 * the merged region survives one row shorter, and only the removed row's own text is lost.
	 * @throws {InvalidOptionError} when `index` is out of range
	 */
	removeRow(index: number): void {
		removeRow(this.tbl, index)
		this.part.markDirty()
	}

	/**
	 * Insert a column, defaulting to the right-hand end. Updates `a:tblGrid` **and** every
	 * row, which is the pair that has to stay in step.
	 *
	 * The mirror of {@link addRow}'s merge case: inserting inside a horizontal merge widens
	 * it instead of splitting it.
	 * @param {number} [index] - where to insert, `0`..`columnCount`; appended when omitted
	 * @param {number} [widthEmu] - the new column's width; defaults to 1 inch
	 * @throws {InvalidOptionError} when `index` is outside that range
	 */
	addColumn(index?: number, widthEmu?: number): void {
		insertColumn(this.tbl, index, widthEmu)
		this.part.markDirty()
	}

	/**
	 * Remove the column at `index`, from `a:tblGrid` and from every row.
	 *
	 * Inside a horizontal merge the region narrows by one and keeps its content: a covered
	 * cell is dropped rather than the origin. Elsewhere the column's cells go with it.
	 * @throws {InvalidOptionError} when `index` is out of range
	 */
	removeColumn(index: number): void {
		removeColumn(this.tbl, index)
		this.part.markDirty()
	}

	/**
	 * Merge the rectangle between two cell positions into one cell. The top-left cell keeps
	 * its content; the rest become covered cells and are emptied, since a covered cell is
	 * never rendered.
	 *
	 * A rectangle that cuts through an existing merge is **rejected**, not silently widened
	 * to fit — the caller asked for a specific region, and quietly producing a different one
	 * is how a layout ends up wrong with nothing to point at. Unmerge first.
	 * @throws {InvalidOptionError} when an index is out of range, the range covers one cell,
	 *   or it partially overlaps an existing merge
	 */
	mergeCells(row1: number, col1: number, row2: number, col2: number): void {
		mergeCells(this.tbl, row1, col1, row2, col2)
		this.part.markDirty()
	}

	/**
	 * Split the merged cell whose **origin** is `(row, col)` back into individual cells.
	 * A no-op on a cell that is not merged; addressing a covered cell instead of the origin
	 * throws, and the message names the origin to use.
	 * @throws {InvalidOptionError} when an index is out of range or the cell is a covered cell
	 */
	unmergeCell(row: number, col: number): void {
		unmergeCell(this.tbl, row, col)
		this.part.markDirty()
	}

	/** Escape hatch: the underlying `a:tbl` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.tbl
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	#tblPrFlag(name: string): boolean {
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		return tblPr ? boolAttr(tblPr, name) === true : false
	}
}

export { TableRow } from './table-row.js'
export {
	TableCell,
	type CellBorder,
	type CellBorders,
	type CellThreeD,
	type TableCellBorderEdit,
} from './table-cell.js'
