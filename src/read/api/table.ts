/**
 * Read/edit proxies for a table (`a:tbl`) hosted in a `p:graphicFrame`.
 *
 * `Table → TableRow[] → TableCell[]`, each wrapping a live DOM element and
 * holding the owning slide `Part` so edits mutate the node in place and call
 * `part.markDirty()`. Cell text reuses the `TextFrame`/`Paragraph`/`Run`
 * proxies, so per-run formatting edits work exactly as on a shape's text.
 */
import type { Part } from '../opc/part.js'
import {
	ELEMENT_NODE,
	attr,
	createElement,
	firstChild,
	getElements,
	intValue,
	removeAttr,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import { TextFrame } from './text.js'

/** A table: a grid of rows and cells inside a graphic frame. */
export class Table {
	constructor(
		private readonly tbl: Element,
		private readonly part: Part
	) {}

	/** The table's rows (`a:tr`) in document (top-to-bottom) order. */
	get rows(): TableRow[] {
		return getElements(this.tbl, 'a:tr').map((tr) => new TableRow(tr, this.part))
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
		private readonly part: Part
	) {}

	/** The row's cells (`a:tc`) in left-to-right order. */
	get cells(): TableCell[] {
		return getElements(this.tr, 'a:tc').map((tc) => new TableCell(tc, this.part))
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
		private readonly part: Part
	) {}

	/** The cell's text frame (`a:txBody`); `null` only if the cell has none (non-conformant). */
	get textFrame(): TextFrame | null {
		const txBody = firstChild(this.tc, 'a:txBody')
		return txBody ? new TextFrame(txBody, this.part) : null
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
		const doc = txBody.ownerDocument
		if (!doc) throw new Error('Cannot edit cell text: a:txBody has no owner document')

		const paragraphs = getElements(txBody, 'a:p')
		// Capture the first run's character formatting before we discard runs.
		const firstRun = paragraphs[0] && firstChild(paragraphs[0], 'a:r')
		const rPrTemplate = firstRun && firstChild(firstRun, 'a:rPr')

		// Collapse to a single paragraph, dropping any extras.
		for (let i = paragraphs.length - 1; i >= 1; i--) txBody.removeChild(paragraphs[i])
		let p = paragraphs[0]
		if (!p) {
			p = createElement(doc, 'a:p')
			txBody.appendChild(p)
		}

		// Remove every run-level child (runs, breaks, fields); keep a:pPr / a:endParaRPr.
		for (const child of [...childElements(p)]) {
			if (child.localName === 'r' || child.localName === 'br' || child.localName === 'fld') p.removeChild(child)
		}

		// Build a single run, carrying over the captured formatting if present.
		const run = createElement(doc, 'a:r')
		if (rPrTemplate) run.appendChild(rPrTemplate.cloneNode(true))
		const t = createElement(doc, 'a:t')
		t.textContent = value
		if (value !== value.trim()) setAttr(t, 'xml:space', 'preserve')
		else removeAttr(t, 'xml:space')
		run.appendChild(t)

		// Insert before a:endParaRPr if present (it must stay last), else append.
		const endParaRPr = firstChild(p, 'a:endParaRPr')
		p.insertBefore(run, endParaRPr)

		this.part.markDirty()
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

/** Direct child elements of `parent` in document order. */
function childElements(parent: Element): Element[] {
	const out: Element[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) out.push(node as Element)
	}
	return out
}
