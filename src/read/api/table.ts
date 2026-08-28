/**
 * Read/edit proxies for a table (`a:tbl`) hosted in a `p:graphicFrame`.
 *
 * `Table → TableRow[] → TableCell[]`, each wrapping a live DOM element and
 * holding the owning slide `Part` so edits mutate the node in place and call
 * `part.markDirty()`. Cell text reuses the `TextFrame`/`Paragraph`/`Run`
 * proxies, so per-run formatting edits work exactly as on a shape's text.
 */
import type { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import {
	attr,
	createElement,
	firstChild,
	getElements,
	getOrAddChild,
	intValue,
	ownerDocumentOf,
	removeAttr,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import { FILL_CHOICES, normalizeHex, setSolidFill } from '../oxml/fill.js'
import {
	ANCHOR_VALUES,
	checkEnum,
	checkFiniteEmu,
	EDGE_QNAMES,
	HORZ_OVERFLOW_VALUES,
	insertTcPrChild,
	TCPR_AFTER,
	tcPrChild,
	VERT_VALUES,
	type TableCellEdge,
} from './table-edit.js'
import { insertColumn, insertRow, mergeCells, removeColumn, removeRow, rowsOf, unmergeCell } from './table-structure.js'
import type { InvalidOptionErrorCode } from '../../codes.js'
import type { ThemeContext } from '../oxml/theme.js'
import { readPictureFill, type PictureFill } from './picture-fill.js'
import { readGradientFill, type GradientFill } from './gradient.js'
import { readPatternFill, type PatternFill } from './pattern-fill.js'
import {
	resolveTableCellStyleFill,
	resolveTableStyle,
	type ResolvedTableStyle,
	type TableConditionFlags,
} from './table-style-resolve.js'
import { resolveSolidFillColor, type ResolvedColor } from './theme-context.js'
import { setTextBodyText, TextFrame } from './text.js'
import { InvalidOptionError, PackageReadError } from '../../errors.js'
import { EMU_PER_POINT } from '../../units.js'

/**
 * One border as {@link TableCell.setBorder} takes it — the write-side mirror of
 * {@link CellBorder}, which is what reading one gives back.
 *
 * Deliberately not `CellBorder` itself: that type reports every field, `null` included, so
 * passing one back would make "leave this alone" unsayable. Here an omitted field is simply
 * not written, which is what an edit usually means.
 */
export interface TableCellBorderEdit {
	/** Border width in points; omitted leaves `@w` unset, which renders as a hairline. */
	widthPt?: number | null
	/** Dash preset (`a:prstDash/@val`, e.g. `sysDash`); omitted leaves the border solid. */
	dash?: string
	/** Explicit colour as 6-hex (`#` optional). Ignored when {@link schemeColor} is given. */
	color?: string
	/** Theme colour token (e.g. `accent1`); preferred over {@link color} when both are set. */
	schemeColor?: string
	/** Write an explicit `a:noFill` — a deliberately suppressed edge, not an absent one. */
	noFill?: boolean
}

/**
 * The context a cell needs to resolve its style-graph fill: the table's resolved
 * `a:tblStyle`, its condition flags, and the grid dimensions. Threaded from
 * {@link Table} down to each {@link TableCell} alongside the cell's own position.
 * `null` when the table references no resolvable style.
 */
interface TableCellStyleContext {
	style: ResolvedTableStyle
	flags: TableConditionFlags
	rowCount: number
	colCount: number
	ctx: ThemeContext
}

/**
 * One edge border of a table cell (`a:tcPr/a:lnL|lnR|lnT|lnB|lnTlToBr|lnBlToTr`).
 * Mirrors the `a:ln`-style decode used for shape strokes: {@link widthPt} from
 * `@w` (EMU → points), {@link dash} from `a:prstDash/@val`, and the stroke colour
 * split into a resolved {@link color} (literal hex) and the raw {@link schemeColor}
 * token — the cell-border counterpart of a shape's line accessors.
 */
export interface CellBorder {
	/** Border width in points (`@w` is EMU; 12700 EMU = 1pt), or `null` when unset. */
	widthPt: number | null
	/** Dash style (`a:prstDash/@val`, e.g. `solid`/`sysDash`), or `null` when unset. */
	dash: string | null
	/**
	 * The stroke colour as a full {@link ResolvedColor} — base `hex`, the raw
	 * `transforms` list (`lumMod`/`shade`/…) in document order, and the
	 * `effectiveHex`/`alpha` after applying them — the same object a shape's
	 * `resolvedLine` gives. `null` when the edge carries no solid fill, or the
	 * colour cannot be made literal.
	 *
	 * This is the field to read when re-authoring a border against a *different*
	 * theme: {@link color} alone is one theme baked in, so a `lumMod`-darkened
	 * accent carried forward as a literal hex silently stops tracking the theme it
	 * came from. `transforms` is what says whether there was anything to track, and
	 * an empty list here means the edge stated none — not that the reader could not
	 * see them.
	 */
	resolvedColor: ResolvedColor | null
	/**
	 * The stroke colour resolved against the table theme to a literal hex — exactly
	 * `resolvedColor?.effectiveHex ?? null`, kept as a flat field because painting a
	 * rule is what most callers want. `null` when the edge has no resolvable colour.
	 */
	color: string | null
	/** Raw `schemeClr` token of the stroke (`a:solidFill/a:schemeClr/@val`), or `null` for an srgb/absent colour. */
	schemeColor: string | null
	/** `true` when the edge is an explicit no-border (`a:noFill`) — a deliberately suppressed side. */
	noFill: boolean
}

/**
 * A table cell's six possible borders, keyed by edge. Each side is `null` when the
 * cell defines no line for it; the diagonals ({@link tlToBr} `╲`, {@link blToTr} `╱`)
 * are rarely authored. The whole object is `null` when the cell carries no border
 * element at all (see {@link TableCell.borders}).
 */
export interface CellBorders {
	left: CellBorder | null
	right: CellBorder | null
	top: CellBorder | null
	bottom: CellBorder | null
	tlToBr: CellBorder | null
	blToTr: CellBorder | null
}

/**
 * A table cell's 3-D bevel (`a:tcPr/a:cell3D`), as read back. Every field is `null` when the
 * source leaves it to the schema default, so a reader can tell "PowerPoint wrote `circle`"
 * from "PowerPoint wrote nothing and `circle` is what it means".
 */
export interface CellThreeD {
	/** `a:cell3D/@prstMaterial` (`ST_PresetMaterialType`), or `null` when unset. */
	material: string | null
	/** The required `a:bevel`. Present whenever the cell has a `a:cell3D` at all. */
	bevel: {
		/** `a:bevel/@prst` (`ST_BevelPresetType`), or `null` when unset. */
		preset: string | null
		/** `a:bevel/@w` converted from EMU to points, or `null` when unset. */
		widthPt: number | null
		/** `a:bevel/@h` converted from EMU to points, or `null` when unset. */
		heightPt: number | null
	}
	/** `a:cell3D/a:lightRig`, or `null` when the cell leaves lighting to the renderer. */
	lightRig: { rig: string | null; dir: string | null } | null
}

/** A table: a grid of rows and cells inside a graphic frame. */
export class Table {
	constructor(
		private readonly tbl: Element,
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to each cell's text for `Run.resolvedColor`. */
		private readonly themeColors?: ThemeContext,
		/** The deck package, for resolving `a:tableStyleId` against `tableStyles.xml` (style-graph cell fills). */
		private readonly opc?: OpcPackage,
		/** The owning slide's relationships, for resolving a cell picture fill's `r:embed` to a partname. */
		private readonly rels?: Relationships
	) {}

	/** The table's rows (`a:tr`) in document (top-to-bottom) order. */
	get rows(): TableRow[] {
		const style = this.#styleContext()
		return getElements(this.tbl, 'a:tr').map(
			(tr, rowIndex) => new TableRow(tr, this.part, this.themeColors, style, rowIndex, this.rels)
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

	/** The per-cell style-resolution context, or `null` when no style resolves or there is no theme context. */
	#styleContext(): TableCellStyleContext | null {
		if (!this.opc || !this.themeColors) return null
		const style = resolveTableStyle(this.opc, this.styleId)
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
			ctx: this.themeColors,
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
		return getElements(grid, 'a:gridCol').map((col) => intValue(attr(col, 'w')))
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
		if (!this.themeColors) return null
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		if (!tblPr || !FILL_CHOICES.some((q) => firstChild(tblPr, q))) return null
		return resolveSolidFillColor(tblPr, this.themeColors)
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
		if (!this.themeColors) return null
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		return tblPr ? readGradientFill(tblPr, this.themeColors) : null
	}

	/**
	 * The table's pattern (hatch) background (`a:tblPr/a:pattFill`), or `null` when it is not
	 * pattern-filled.
	 */
	get patternFill(): PatternFill | null {
		if (!this.themeColors) return null
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		return tblPr ? readPatternFill(tblPr, this.themeColors) : null
	}

	/**
	 * The raw `schemeClr` token of the table's own background
	 * (`a:tblPr/a:solidFill/a:schemeClr/@val`), or `null` for an absent or `srgbClr` fill.
	 * {@link resolvedFill} is the literal it resolves to; this is the unresolved reference,
	 * which is what a replica should carry so the copy still tracks its theme.
	 */
	get fillSchemeColor(): string | null {
		const tblPr = firstChild(this.tbl, 'a:tblPr')
		const fill = tblPr && firstChild(tblPr, 'a:solidFill')
		const scheme = fill && firstChild(fill, 'a:schemeClr')
		return scheme ? attr(scheme, 'val') : null
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
		return this.rows[rowIndex]?.cells[columnIndex] ?? null
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
		return new TableRow(tr, this.part, this.themeColors, this.#styleContext(), rowsOf(this.tbl).indexOf(tr), this.rels)
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
		return tblPr ? attr(tblPr, name) === '1' : false
	}
}

/** One table row (`a:tr`). */
export class TableRow {
	constructor(
		private readonly tr: Element,
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to each {@link TableCell}. */
		private readonly themeColors?: ThemeContext,
		/** The table's style-resolution context, threaded to each cell for {@link TableCell.resolvedFill}. */
		private readonly style?: TableCellStyleContext | null,
		/** This row's zero-based index in the table, for style-graph banding/edge conditions. */
		private readonly rowIndex = 0,
		/** The owning slide's relationships, threaded to each cell for {@link TableCell.pictureFill}. */
		private readonly rels?: Relationships
	) {}

	/** The row's cells (`a:tc`) in left-to-right order. */
	get cells(): TableCell[] {
		return getElements(this.tr, 'a:tc').map(
			(tc, colIndex) => new TableCell(tc, this.part, this.themeColors, this.style, this.rowIndex, colIndex, this.rels)
		)
	}

	/** Row height in EMU (`a:tr/@h`), or `null` if unset. */
	get heightEmu(): number | null {
		return intValue(attr(this.tr, 'h'))
	}

	/** Escape hatch: the underlying `a:tr` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.tr
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}
}

/** One table cell (`a:tc`). */
export class TableCell {
	constructor(
		private readonly tc: Element,
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to the cell's text for `Run.resolvedColor`. */
		private readonly themeColors?: ThemeContext,
		/** The table's style-resolution context, for the {@link resolvedFill} style-graph fallback. */
		private readonly style?: TableCellStyleContext | null,
		/** This cell's zero-based row index in the table. */
		private readonly rowIndex = 0,
		/** This cell's zero-based column index in its row. */
		private readonly colIndex = 0,
		/** The owning slide's relationships, for resolving {@link pictureFill}'s `r:embed` to a partname. */
		private readonly rels?: Relationships
	) {}

	/** The cell's text frame (`a:txBody`); `null` only if the cell has none (non-conformant). */
	get textFrame(): TextFrame | null {
		const txBody = firstChild(this.tc, 'a:txBody')
		return txBody ? new TextFrame(txBody, this.part, this.themeColors) : null
	}

	/** The cell's text, paragraphs joined by `\n`. */
	get text(): string {
		return this.textFrame?.text ?? ''
	}

	/**
	 * Replace the cell's text with a single paragraph and run, preserving the
	 * formatting (`a:rPr`) of the cell's first existing run when there is one.
	 * For finer control (multiple runs, per-run formatting), edit
	 * `textFrame.paragraphs[].runs[]` directly.
	 */
	set text(value: string) {
		const txBody = firstChild(this.tc, 'a:txBody')
		if (!txBody) throw new PackageReadError('table/cell-has-no-text-body', 'Table cell has no a:txBody to set text on')
		setTextBodyText(txBody, value)
		this.part.markDirty()
	}

	/** The cell's properties element (`a:tcPr`), or `null` when the cell defines none. */
	#tcPr(): Element | null {
		return firstChild(this.tc, 'a:tcPr')
	}

	/**
	 * The cell's `a:tcPr`, creating it when absent.
	 *
	 * `CT_TableCell` sequences `a:txBody` then `a:tcPr` then `a:extLst`, so a newly created
	 * one is inserted before `a:extLst` — appending it blindly would put it after and make
	 * the part invalid. Every setter below goes through here for that reason.
	 */
	#getOrAddTcPr(): Element {
		return getOrAddChild(this.tc, 'a:tcPr', ['a:extLst'])
	}

	/**
	 * Set (or clear) the cell's vertical text anchor (`a:tcPr/@anchor`).
	 * `null` removes the attribute, leaving PowerPoint's default of top.
	 * @throws {InvalidOptionError} when the value is outside `ST_TextAnchoringType`
	 */
	setAnchor(value: string | null): void {
		this.#setTcPrAttr('anchor', value, ANCHOR_VALUES, 'table/invalid-cell-anchor')
	}

	/**
	 * Set (or clear) the cell's text direction (`a:tcPr/@vert`), e.g. `'vert270'`.
	 * `null` removes the attribute, leaving the default horizontal text.
	 * @throws {InvalidOptionError} when the value is outside `ST_TextVerticalType`
	 */
	setVerticalText(value: string | null): void {
		this.#setTcPrAttr('vert', value, VERT_VALUES, 'table/invalid-cell-vert')
	}

	/**
	 * Set (or clear) `a:tcPr/@horzOverflow` — what a single glyph too wide for the cell does.
	 * `null` removes the attribute; so does `'clip'` in effect, since it is the schema
	 * default, but the attribute is still written as asked.
	 * @throws {InvalidOptionError} when the value is outside `ST_TextHorzOverflowType`
	 */
	setHorzOverflow(value: string | null): void {
		this.#setTcPrAttr('horzOverflow', value, HORZ_OVERFLOW_VALUES, 'table/invalid-cell-overflow')
	}

	/**
	 * Centre the cell's text block horizontally, or stop doing so (`a:tcPr/@anchorCtr`).
	 * `false` removes the attribute rather than writing `"0"`, since `false` is the schema
	 * default and the two are indistinguishable to a renderer.
	 */
	setAnchorCtr(value: boolean): void {
		const tcPr = this.#getOrAddTcPr()
		if (value) setAttr(tcPr, 'anchorCtr', '1')
		else removeAttr(tcPr, 'anchorCtr')
		this.part.markDirty()
	}

	/**
	 * Set the cell's text insets in EMU (`a:tcPr/@marL`/`@marR`/`@marT`/`@marB`).
	 *
	 * Partial: only the sides named are touched, so `{ left: 0 }` flushes the text left and
	 * leaves the other three alone. A side given as `null` has its attribute removed, which
	 * returns it to the schema default (91440 EMU left/right, 45720 top/bottom) — not to
	 * zero. Pass `{}` to change nothing.
	 * @throws {InvalidOptionError} when a value is not a finite number
	 */
	setMarginsEmu(margins: {
		left?: number | null
		right?: number | null
		top?: number | null
		bottom?: number | null
	}): void {
		const tcPr = this.#getOrAddTcPr()
		const sides = [
			['marL', margins.left],
			['marR', margins.right],
			['marT', margins.top],
			['marB', margins.bottom],
		] as const
		for (const [name, value] of sides) {
			if (value === undefined) continue
			if (value === null) removeAttr(tcPr, name)
			else setAttr(tcPr, name, String(checkFiniteEmu(value, name, 'table/invalid-cell-margin')))
		}
		this.part.markDirty()
	}

	/**
	 * Set (or clear) one of the cell's six borders — the four edges and the two diagonals.
	 *
	 * `null` removes the element entirely, which is "inherit", and is a different thing from
	 * `{ noFill: true }`, which writes an explicit no-border and suppresses whatever would
	 * otherwise be inherited. The element is inserted at its schema position rather than
	 * appended: `CT_TableCellProperties` is a sequence, and an out-of-order `a:tcPr` is
	 * reported by PowerPoint as a corrupt file rather than as a bad edit.
	 *
	 * Colour is either `color` (6-hex, `#` optional) or `schemeColor` (a theme token); giving
	 * both prefers the token, matching how the read side reports them.
	 * @throws {InvalidOptionError} when the width or colour cannot be written
	 */
	setBorder(edge: TableCellEdge, border: TableCellBorderEdit | null): void {
		const qname = EDGE_QNAMES[edge]
		if (!qname) {
			throw new InvalidOptionError(
				'table/invalid-cell-border',
				`Unknown table cell edge: ${JSON.stringify(edge)}. Expected one of: ${Object.keys(EDGE_QNAMES).join(', ')}.`
			)
		}
		if (border === null) {
			const tcPr = this.#tcPr()
			if (!tcPr || !firstChild(tcPr, qname)) return
			removeChildrenByQName(tcPr, [qname])
			this.part.markDirty()
			return
		}

		const tcPr = this.#getOrAddTcPr()
		// Rebuilt rather than patched in place: a half-edited `a:ln` (say, a new colour left
		// beside a stale `a:noFill`) is a state neither the reader nor PowerPoint expects, and
		// the element is small enough that replacing it is simpler than reconciling it.
		removeChildrenByQName(tcPr, [qname])
		const doc = ownerDocumentOf(tcPr)
		const ln = createElement(doc, qname)
		if (border.widthPt !== undefined && border.widthPt !== null) {
			setAttr(ln, 'w', String(checkFiniteEmu(border.widthPt * EMU_PER_POINT, 'widthPt', 'table/invalid-cell-border')))
		}
		if (border.noFill) {
			ln.appendChild(createElement(doc, 'a:noFill'))
		} else if (border.schemeColor || border.color) {
			const fill = createElement(doc, 'a:solidFill')
			const scheme = border.schemeColor
			const clr = createElement(doc, scheme ? 'a:schemeClr' : 'a:srgbClr')
			setAttr(clr, 'val', scheme ? scheme : normalizeHex(border.color as string))
			fill.appendChild(clr)
			ln.appendChild(fill)
		}
		// `a:prstDash` follows the fill group in CT_LineProperties, so it is appended after.
		if (border.dash) {
			const dash = createElement(doc, 'a:prstDash')
			setAttr(dash, 'val', border.dash)
			ln.appendChild(dash)
		}
		insertTcPrChild(tcPr, qname, ln)
		this.part.markDirty()
	}

	/**
	 * Replace the cell's fill with a solid colour (`a:tcPr/a:solidFill`), or clear it.
	 *
	 * `null` removes the `a:solidFill` and lets the cell inherit from the table style again.
	 * That is not the same as {@link noFill}, which writes an explicit `a:noFill` and so
	 * suppresses the inherited shading. Any competing fill choice (`a:gradFill`,
	 * `a:blipFill`, …) is dropped first — `EG_FillProperties` admits one.
	 * @throws {InvalidOptionError} when `color` is not a 6-digit hex string
	 */
	setFillColor(color: string | null): void {
		this.#setFill(color === null ? null : { qname: 'a:srgbClr', val: normalizeHex(color) })
	}

	/**
	 * Replace the cell's fill with a theme colour token (`a:solidFill/a:schemeClr/@val`), or
	 * clear it. Preferred over {@link setFillColor} when the deck's theme should keep
	 * driving the colour.
	 */
	setFillSchemeColor(token: string | null): void {
		this.#setFill(token === null ? null : { qname: 'a:schemeClr', val: token })
	}

	/**
	 * Write an explicit `<a:noFill/>` on the cell — a transparent cell that shows the table
	 * background (or the slide) through. Distinct from `setFillColor(null)`, which removes
	 * the fill and lets the table style's shading apply again.
	 */
	noFill(): void {
		const tcPr = this.#getOrAddTcPr()
		removeChildrenByQName(tcPr, FILL_CHOICES)
		tcPrChild(tcPr, 'a:noFill')
		this.part.markDirty()
	}

	/** Set or clear one `a:tcPr` attribute, vetting it against its schema enum first. */
	#setTcPrAttr(name: string, value: string | null, valid: readonly string[], code: InvalidOptionErrorCode): void {
		if (value === null) {
			const tcPr = this.#tcPr()
			if (!tcPr) return
			removeAttr(tcPr, name)
			this.part.markDirty()
			return
		}
		const checked = checkEnum(value, valid, name, code)
		setAttr(this.#getOrAddTcPr(), name, checked)
		this.part.markDirty()
	}

	/** Replace (or remove) the cell's solid fill, keeping `EG_FillProperties` single-valued. */
	#setFill(color: { qname: string; val: string } | null): void {
		if (color === null) {
			const tcPr = this.#tcPr()
			if (!tcPr || !firstChild(tcPr, 'a:solidFill')) return
			removeChildrenByQName(tcPr, ['a:solidFill'])
			this.part.markDirty()
			return
		}
		setSolidFill(this.#getOrAddTcPr(), TCPR_AFTER['a:solidFill'] ?? [], color)
		this.part.markDirty()
	}

	/**
	 * The cell's solid fill resolved against the table's theme colour context to a
	 * literal hex — the table-cell counterpart of
	 * {@link import('./shapes.js').AutoShape.resolvedFill}. The cell's own
	 * `a:tcPr/a:solidFill` wins; when the cell defines none, this falls back to the
	 * table **style** graph (the `firstRow`/banded/`wholeTbl` shading the
	 * `a:tableStyleId` supplies — see {@link Table.resolvedStyle}), so a styled cell
	 * with an empty `a:tcPr` reports the colour PowerPoint actually renders rather
	 * than `null`.
	 *
	 * A cell that carries *some other* fill choice (`a:blipFill`/`a:gradFill`/
	 * `a:pattFill`/`a:noFill`) overrides the style graph in PowerPoint, so this
	 * reports `null` for one rather than falling through to the inherited shading —
	 * the same guard {@link import('./shapes.js').AutoShape.resolvedFill} applies to
	 * the style matrix. Read {@link pictureFill} for an image-filled cell. Also
	 * `null` with no theme context, or when neither source yields a solid colour (an
	 * unmapped token, an explicit style `a:noFill`). The returned
	 * {@link ResolvedColor} carries `effectiveHex` (the base colour with its
	 * `lumMod`/`lumOff`/… transforms applied) — read that for the final colour.
	 */
	get resolvedFill(): ResolvedColor | null {
		if (!this.themeColors) return null
		if (this.hasOwnFill) return resolveSolidFillColor(this.#tcPr(), this.themeColors)
		return this.style ? this.#styleFill() : null
	}

	/**
	 * Whether the cell carries a fill of its **own** (`a:tcPr` holds an `EG_FillProperties`
	 * child), as opposed to inheriting one from the table style's header/banding rules.
	 *
	 * This is the flag that disambiguates {@link resolvedFill}, which deliberately reports
	 * either — it answers "what colour does this cell render as", and both sources are valid
	 * answers to that. Anything that has to reproduce the cell rather than describe it needs
	 * to know which: baking an inherited banding colour into a copy makes the copy look right
	 * until someone changes its table style, and then nothing moves. `false` for a cell whose
	 * `a:tcPr` is empty or absent, whatever the style graph would render it as.
	 */
	get hasOwnFill(): boolean {
		const tcPr = this.#tcPr()
		return !!tcPr && FILL_CHOICES.some((q) => firstChild(tcPr, q))
	}

	/**
	 * `true` when the cell sets an explicit no-fill (`a:tcPr/a:noFill`) — a deliberately
	 * transparent cell showing the table background (or the slide) through. The cell-side
	 * counterpart of {@link import('./shapes.js').AutoShape.fillNoFill}, and what
	 * {@link noFill} writes.
	 *
	 * {@link hasOwnFill} is not this question: it is `true` for *any* `EG_FillProperties`
	 * child, so on its own it cannot separate a suppressed fill from a gradient or an image
	 * one — and every colour accessor ({@link resolvedFill}, {@link fillColor},
	 * {@link fillSchemeColor}) reports `null` for a no-fill cell exactly as it does for a
	 * cell that inherits its shading from the table style. Deriving it as "has a fill of its
	 * own, and no accessor recognises it" instead of reading it has two failure modes: it
	 * folds `a:grpFill` in with `a:noFill`, and its meaning changes silently the day a
	 * further fill kind gets an accessor. The two paint completely differently — an
	 * inherited banding colour versus nothing at all — so a consumer that cannot tell them
	 * apart paints a transparent cell in the style's shading.
	 */
	get fillNoFill(): boolean {
		const tcPr = this.#tcPr()
		return !!tcPr && !!firstChild(tcPr, 'a:noFill')
	}

	/**
	 * The cell's picture (image) fill (`a:tcPr/a:blipFill`), or `null` when the cell
	 * is not image-filled. The cell counterpart of
	 * {@link import('./shapes.js').AutoShape.pictureFill}: {@link resolvedFill}
	 * decodes only solid colours, so without this an image-filled cell is
	 * indistinguishable from an empty one. Carries the embedded image
	 * ({@link PictureFill.relId}/{@link PictureFill.partName}) plus the stretch/tile
	 * geometry; {@link PictureFill.partName} needs the owning slide's relationships,
	 * which a {@link Table} built without them cannot supply.
	 */
	get pictureFill(): PictureFill | null {
		const tcPr = this.#tcPr()
		return tcPr ? readPictureFill(tcPr, this.rels ?? null) : null
	}

	/** The fill this cell inherits from the table style graph, or `null` when the style defines none for it. */
	#styleFill(): ResolvedColor | null {
		const s = this.style
		if (!s) return null
		return resolveTableCellStyleFill(
			s.style.element_,
			s.flags,
			this.rowIndex,
			this.colIndex,
			s.rowCount,
			s.colCount,
			s.ctx
		)
	}

	/**
	 * The cell's gradient fill (`a:tcPr/a:gradFill`), or `null` when the cell is not
	 * gradient-filled. The cell twin of {@link Table.gradientFill}, and needed for the same
	 * reason {@link pictureFill} is: {@link resolvedFill} reports `null` for every non-solid
	 * choice, so without this a gradient cell is indistinguishable from an unfilled one.
	 */
	get gradientFill(): GradientFill | null {
		const tcPr = this.#tcPr()
		return tcPr && this.themeColors ? readGradientFill(tcPr, this.themeColors) : null
	}

	/**
	 * The cell's pattern (hatch) fill (`a:tcPr/a:pattFill`), or `null` when the cell is not
	 * pattern-filled.
	 */
	get patternFill(): PatternFill | null {
		const tcPr = this.#tcPr()
		return tcPr && this.themeColors ? readPatternFill(tcPr, this.themeColors) : null
	}

	/**
	 * The raw `schemeClr` token of the cell's solid fill (`a:tcPr/a:solidFill/a:schemeClr/@val`),
	 * e.g. `accent1`/`bg1`, or `null` when the fill is absent or an explicit `srgbClr`.
	 * The resolved literal is {@link resolvedFill}; this is the unresolved reference.
	 */
	get fillSchemeColor(): string | null {
		return this.#fillSchemeColorOf(this.#tcPr() ?? this.tc)
	}

	/**
	 * The cell's edge borders (`a:tcPr/a:lnL|lnR|lnT|lnB|lnTlToBr|lnBlToTr`), or
	 * `null` when the cell defines none. Each present edge decodes to a
	 * {@link CellBorder} (width / dash / resolved colour + its raw token and transform
	 * list / suppressed flag);
	 * absent edges are `null`. Cell borders are the biggest visible table gap — a
	 * replica built only from geometry and fill draws every cell edge-to-edge with
	 * no rule, so this surfaces the per-side stroke the writer's `border` option emits.
	 */
	get borders(): CellBorders | null {
		const tcPr = this.#tcPr()
		if (!tcPr) return null
		const decode = (qname: string): CellBorder | null => {
			const ln = firstChild(tcPr, qname)
			if (!ln) return null
			const w = intValue(attr(ln, 'w'))
			const dash = firstChild(ln, 'a:prstDash')
			const scheme = this.#fillSchemeColorOf(ln)
			const resolved = this.themeColors ? resolveSolidFillColor(ln, this.themeColors) : null
			return {
				widthPt: w === null ? null : w / EMU_PER_POINT,
				dash: dash ? (attr(dash, 'val') ?? null) : null,
				resolvedColor: resolved,
				color: resolved ? resolved.effectiveHex : null,
				schemeColor: scheme,
				noFill: !!firstChild(ln, 'a:noFill'),
			}
		}
		const borders: CellBorders = {
			left: decode('a:lnL'),
			right: decode('a:lnR'),
			top: decode('a:lnT'),
			bottom: decode('a:lnB'),
			tlToBr: decode('a:lnTlToBr'),
			blToTr: decode('a:lnBlToTr'),
		}
		return Object.values(borders).some((b) => b !== null) ? borders : null
	}

	/** The `schemeClr` token of a container's solid fill (`a:solidFill/a:schemeClr/@val`), or `null`. */
	#fillSchemeColorOf(container: Element): string | null {
		const fill = firstChild(container, 'a:solidFill')
		const scheme = fill && firstChild(fill, 'a:schemeClr')
		return scheme ? attr(scheme, 'val') : null
	}

	/**
	 * The cell's text direction (`a:tcPr/@vert`), e.g. `vert270` for a bottom-to-top
	 * vertical label, or `null` for default horizontal text.
	 */
	get verticalText(): string | null {
		const tcPr = this.#tcPr()
		return (tcPr && attr(tcPr, 'vert')) ?? null
	}

	/**
	 * The cell's vertical text anchor (`a:tcPr/@anchor`): `t`/`ctr`/`b` (top/middle/
	 * bottom), or `null` when unset (PowerPoint defaults to top).
	 */
	get anchor(): string | null {
		const tcPr = this.#tcPr()
		return (tcPr && attr(tcPr, 'anchor')) ?? null
	}

	/**
	 * Whether the cell's whole text **block** is centred horizontally within it
	 * (`a:tcPr/@anchorCtr`). `false` when unset — that is the schema default, and it is
	 * what PowerPoint writes nothing for.
	 *
	 * Distinct from a paragraph's `align`, which decides where each *line* sits inside the
	 * text block; this decides where the block sits inside the cell.
	 */
	get anchorCtr(): boolean {
		const tcPr = this.#tcPr()
		return tcPr ? attr(tcPr, 'anchorCtr') === '1' || attr(tcPr, 'anchorCtr') === 'true' : false
	}

	/**
	 * The cell's 3-D bevel (`a:tcPr/a:cell3D`), or `null` when it has none.
	 *
	 * `CT_Cell3D` requires an `a:bevel`, so a present `cell3D` always reports a bevel;
	 * `lightRig` is `null` when the cell leaves the scene lighting to the renderer. Sizes are
	 * reported in **points** (the attributes are EMU), matching how the write option takes them.
	 */
	get cell3D(): CellThreeD | null {
		const tcPr = this.#tcPr()
		const cell3D = tcPr && firstChild(tcPr, 'a:cell3D')
		if (!cell3D) return null
		const bevel = firstChild(cell3D, 'a:bevel')
		const rig = firstChild(cell3D, 'a:lightRig')
		const pts = (value: string | null): number | null => {
			const emu = intValue(value)
			return emu === null ? null : emu / EMU_PER_POINT
		}
		return {
			material: attr(cell3D, 'prstMaterial') ?? null,
			bevel: {
				preset: bevel ? (attr(bevel, 'prst') ?? null) : null,
				widthPt: bevel ? pts(attr(bevel, 'w')) : null,
				heightPt: bevel ? pts(attr(bevel, 'h')) : null,
			},
			lightRig: rig ? { rig: attr(rig, 'rig') ?? null, dir: attr(rig, 'dir') ?? null } : null,
		}
	}

	/**
	 * The cell's unique identifier (`a:tc/@id`), or `null` when it has none.
	 *
	 * This is what {@link headerIds} references: a data cell names the header cells that
	 * govern it by their `@id`, which is how a complex table tells a screen reader what a
	 * value means. Unrelated to the shape-level `p:cNvPr/@id` — it is an `xs:string` scoped
	 * to the table, not a slide-wide numeric id.
	 *
	 * **PowerPoint does not write this and strips it on save**, so it will read `null` on
	 * any deck PowerPoint has saved, however it was produced. It is surfaced for decks from
	 * other producers. There is deliberately no write-API counterpart — see
	 * `test/read/fixtures/authoring/probe-table-cell-a11y-and-3d.ps1` for the measurement.
	 */
	get id(): string | null {
		return attr(this.tc, 'id') ?? null
	}

	/**
	 * The `@id`s of the header cells that govern this cell
	 * (`a:tcPr/a:headers/a:header/@val`), in document order. Empty when the cell declares no
	 * header association — which is the common case, since only a complex table needs one.
	 *
	 * Carries the same caveat as {@link id}: PowerPoint strips `a:headers` on save, so this
	 * is empty on any PowerPoint-saved deck. `Table.firstRowHeader` (`a:tblPr/@firstRow`) is
	 * the header marker PowerPoint does keep.
	 */
	get headerIds(): string[] {
		const tcPr = this.#tcPr()
		const headers = tcPr && firstChild(tcPr, 'a:headers')
		if (!headers) return []
		return getElements(headers, 'a:header')
			.map((header) => attr(header, 'val'))
			.filter((val): val is string => val !== null && val !== '')
	}

	/**
	 * How the cell treats a glyph too wide for its text width (`a:tcPr/@horzOverflow`):
	 * `'clip'` cuts it at the cell edge, `'overflow'` lets it draw past. `null` when unset
	 * (PowerPoint clips). Not a wrap flag — cell text always wraps to the column width.
	 */
	get horzOverflow(): string | null {
		const tcPr = this.#tcPr()
		return (tcPr && attr(tcPr, 'horzOverflow')) ?? null
	}

	/**
	 * The cell's text insets in EMU (`a:tcPr/@marL`/`@marR`/`@marT`/`@marB`), or
	 * `null` when the cell sets none. Each side is `null` when only some are set.
	 */
	get marginsEmu(): { left: number | null; right: number | null; top: number | null; bottom: number | null } | null {
		const tcPr = this.#tcPr()
		if (!tcPr) return null
		const left = intValue(attr(tcPr, 'marL'))
		const right = intValue(attr(tcPr, 'marR'))
		const top = intValue(attr(tcPr, 'marT'))
		const bottom = intValue(attr(tcPr, 'marB'))
		if (left === null && right === null && top === null && bottom === null) return null
		return { left, right, top, bottom }
	}

	/** Number of grid columns this cell spans (`a:tc/@gridSpan`), default 1. */
	get gridSpan(): number {
		return intValue(attr(this.tc, 'gridSpan')) ?? 1
	}

	/** Number of rows this cell spans (`a:tc/@rowSpan`), default 1. */
	get rowSpan(): number {
		return intValue(attr(this.tc, 'rowSpan')) ?? 1
	}

	/** Whether this cell is a continuation of a merge (`@hMerge` or `@vMerge`), i.e. not the merge origin. */
	get isMergeContinuation(): boolean {
		return attr(this.tc, 'hMerge') === '1' || attr(this.tc, 'vMerge') === '1'
	}

	/** Escape hatch: the underlying `a:tc` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.tc
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}
}
