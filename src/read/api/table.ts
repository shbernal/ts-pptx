/**
 * Read/edit proxies for a table (`a:tbl`) hosted in a `p:graphicFrame`.
 *
 * `Table → TableRow[] → TableCell[]`, each wrapping a live DOM element and
 * holding the owning slide `Part` so edits mutate the node in place and call
 * `part.markDirty()`. Cell text reuses the `TextFrame`/`Paragraph`/`Run`
 * proxies, so per-run formatting edits work exactly as on a shape's text.
 */
import type { Part } from '../opc/part.js'
import { attr, firstChild, getElements, intValue, type Element } from '../oxml/dom.js'
import type { FlattenContext } from '../oxml/theme.js'
import { resolveSolidFillColor, type ResolvedColor } from './theme-context.js'
import { setTextBodyText, TextFrame } from './text.js'

/** A table: a grid of rows and cells inside a graphic frame. */
export class Table {
	constructor(
		private readonly tbl: Element,
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to each cell's text for `Run.resolvedColor`. */
		private readonly themeColors?: FlattenContext
	) {}

	/** The table's rows (`a:tr`) in document (top-to-bottom) order. */
	get rows(): TableRow[] {
		return getElements(this.tbl, 'a:tr').map((tr) => new TableRow(tr, this.part, this.themeColors))
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

	/** The underlying `a:tbl` element, for advanced reads and future mutation. */
	get element_(): Element {
		return this.tbl
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
		private readonly themeColors?: FlattenContext
	) {}

	/** The row's cells (`a:tc`) in left-to-right order. */
	get cells(): TableCell[] {
		return getElements(this.tr, 'a:tc').map((tc) => new TableCell(tc, this.part, this.themeColors))
	}

	/** Row height in EMU (`a:tr/@h`), or `null` if unset. */
	get heightEmu(): number | null {
		return intValue(attr(this.tr, 'h'))
	}

	/** The underlying `a:tr` element. */
	get element_(): Element {
		return this.tr
	}
}

/** One table cell (`a:tc`). */
export class TableCell {
	constructor(
		private readonly tc: Element,
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to the cell's text for `Run.resolvedColor`. */
		private readonly themeColors?: FlattenContext
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
		if (!txBody) throw new Error('Table cell has no a:txBody to set text on')
		setTextBodyText(txBody, value)
		this.part.markDirty()
	}

	/** The cell's properties element (`a:tcPr`), or `null` when the cell defines none. */
	#tcPr(): Element | null {
		return firstChild(this.tc, 'a:tcPr')
	}

	/**
	 * The cell's solid fill (`a:tcPr/a:solidFill`) resolved against the table's
	 * theme colour context to a literal hex — the table-cell counterpart of
	 * {@link import('./shapes.js').AutoShape.resolvedFill}. `null` when the cell has
	 * no solid fill, or the colour cannot be made literal (no theme context, an
	 * unmapped token, or a non-solid fill). The returned {@link ResolvedColor}
	 * carries `effectiveHex` (the base colour with its `lumMod`/`lumOff`/… transforms
	 * applied) — read that for the final rendered colour.
	 */
	get resolvedFill(): ResolvedColor | null {
		return this.themeColors ? resolveSolidFillColor(this.#tcPr(), this.themeColors) : null
	}

	/**
	 * The raw `schemeClr` token of the cell's solid fill (`a:tcPr/a:solidFill/a:schemeClr/@val`),
	 * e.g. `accent1`/`bg1`, or `null` when the fill is absent or an explicit `srgbClr`.
	 * The resolved literal is {@link resolvedFill}; this is the unresolved reference.
	 */
	get fillSchemeColor(): string | null {
		const fill = firstChild(this.#tcPr() ?? this.tc, 'a:solidFill')
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

	/** The underlying `a:tc` element. */
	get element_(): Element {
		return this.tc
	}
}
