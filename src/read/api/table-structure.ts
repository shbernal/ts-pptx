/**
 * Structural edits to an existing `a:tbl`: adding and removing rows and columns, and
 * merging and unmerging cells.
 *
 * The one fact that makes this tractable is that a stored table's grid is already
 * **rectangular**. Unlike the write path — which receives lopsided row arrays and has to
 * build the merge grid itself (`gen/slide/objects/table.ts` STEP 3) — every `a:tr` in a
 * conformant `a:tbl` holds exactly one `a:tc` per grid column, with the covered half of each
 * span present and flagged. So "the cell at (r, c)" is simply the c-th `a:tc` of the r-th
 * `a:tr`, and an edit's whole job is keeping that rectangle true.
 *
 * Three things have to stay in step, and every function here is written around them:
 *
 * 1. `a:tblGrid/a:gridCol` count == each row's `a:tc` count.
 * 2. A cell with `@gridSpan="n"` is followed by exactly `n-1` cells carrying `@hMerge="1"`,
 *    and a cell with `@rowSpan="n"` is under-hung by `n-1` cells carrying `@vMerge="1"`, in
 *    the same columns.
 * 3. A covered cell carries no span attributes of its own — its extent is the origin's.
 *
 * Breaking any of them produces a table PowerPoint reports as a corrupt file rather than as
 * a bad edit, which is why the span bookkeeping here is explicit rather than incidental.
 */
import { InvalidOptionError } from '../../errors.js'
import { EMU_PER_INCH } from '../../units.js'
import {
	attr,
	createElement,
	firstChild,
	getElements,
	intValue,
	ownerDocumentOf,
	removeAttr,
	setAttr,
	type Document,
	type Element,
} from '../oxml/dom.js'

/** Every `a:tr` of a table, in document order. */
export function rowsOf(tbl: Element): Element[] {
	return getElements(tbl, 'a:tr')
}

/** Every `a:tc` of a row, in document (left-to-right) order. */
function cellsOf(tr: Element): Element[] {
	return getElements(tr, 'a:tc')
}

/** The table's `a:tblGrid`, or a thrown error — `CT_Table` requires one. */
function gridOf(tbl: Element): Element {
	const grid = firstChild(tbl, 'a:tblGrid')
	if (!grid)
		throw new InvalidOptionError(
			'table/column-index-out-of-range',
			'This table has no a:tblGrid, so its columns cannot be edited'
		)
	return grid
}

/** A cell's `@gridSpan`, defaulting to 1. */
function gridSpanOf(tc: Element): number {
	return intValue(attr(tc, 'gridSpan')) ?? 1
}

/** A cell's `@rowSpan`, defaulting to 1. */
function rowSpanOf(tc: Element): number {
	return intValue(attr(tc, 'rowSpan')) ?? 1
}

/** Whether a cell is the covered half of a horizontal merge. */
function isHMerge(tc: Element): boolean {
	return attr(tc, 'hMerge') === '1' || attr(tc, 'hMerge') === 'true'
}

/** Whether a cell is the covered half of a vertical merge. */
function isVMerge(tc: Element): boolean {
	return attr(tc, 'vMerge') === '1' || attr(tc, 'vMerge') === 'true'
}

/**
 * Set a span attribute, removing it when the span collapses to 1 — the schema default, and
 * what PowerPoint itself writes for an unmerged cell.
 */
function setSpan(tc: Element, name: 'gridSpan' | 'rowSpan', value: number): void {
	if (value <= 1) removeAttr(tc, name)
	else setAttr(tc, name, String(value))
}

/**
 * Walk left from `col` to the origin of the horizontal run containing it.
 * A run is an origin followed by its `hMerge` continuations, so the origin is the first
 * cell at or left of `col` that is not one.
 * @returns the origin's column index, or `col` itself when the cell stands alone
 */
function hRunOrigin(cells: Element[], col: number): number {
	let idx = col
	while (idx > 0) {
		const cell = cells[idx]
		if (!cell || !isHMerge(cell)) break
		idx -= 1
	}
	return idx
}

/**
 * Walk up from `row` to the origin of the vertical run containing `(row, col)`.
 * @returns the origin's row index, or `row` itself when the cell stands alone
 */
function vRunOrigin(rows: Element[][], row: number, col: number): number {
	let idx = row
	while (idx > 0) {
		const cell = rows[idx]?.[col]
		if (!cell || !isVMerge(cell)) break
		idx -= 1
	}
	return idx
}

/** A fresh, empty `a:tc`, shaped the way PowerPoint writes one. */
function makeCell(doc: Document): Element {
	const tc = createElement(doc, 'a:tc')
	const txBody = createElement(doc, 'a:txBody')
	txBody.appendChild(createElement(doc, 'a:bodyPr'))
	txBody.appendChild(createElement(doc, 'a:lstStyle'))
	txBody.appendChild(createElement(doc, 'a:p'))
	tc.appendChild(txBody)
	tc.appendChild(createElement(doc, 'a:tcPr'))
	return tc
}

/** The grid as a row-major array of `a:tc` elements. */
function gridOfCells(tbl: Element): Element[][] {
	return rowsOf(tbl).map((tr) => cellsOf(tr))
}

/** Guard an index against `[0, limit]` (inclusive upper bound, for insertion points). */
function checkIndex(index: number, limit: number, what: 'row' | 'column', inclusive: boolean): number {
	const max = inclusive ? limit : limit - 1
	if (!Number.isInteger(index) || index < 0 || index > max) {
		throw new InvalidOptionError(
			what === 'row' ? 'table/row-index-out-of-range' : 'table/column-index-out-of-range',
			`${what} index ${String(index)} is out of range (expected 0..${max})`
		)
	}
	return index
}

/**
 * Insert a row at `index` (default: append).
 *
 * The subtle case is inserting **through** a vertical merge. If the row currently at `index`
 * holds a `vMerge` continuation in some column, the span it belongs to straddles the
 * insertion point — so the new row must continue that span rather than interrupt it: the
 * origin's `@rowSpan` grows by one and the new cell is another continuation. Interrupting it
 * instead would leave an origin claiming more rows than it has continuations, which is the
 * corrupt-file case.
 */
export function insertRow(tbl: Element, index?: number): Element {
	const rows = rowsOf(tbl)
	const at = index === undefined ? rows.length : checkIndex(index, rows.length, 'row', true)
	const grid = gridOfCells(tbl)
	const colCount = getElements(gridOf(tbl), 'a:gridCol').length
	const doc = ownerDocumentOf(tbl)

	const tr = createElement(doc, 'a:tr')
	// `@h` is required on CT_TableRow. Zero means "auto — as tall as the content needs",
	// which is the right default for a row nobody has sized.
	setAttr(tr, 'h', '0')

	for (let col = 0; col < colCount; col++) {
		const tc = makeCell(doc)
		// A span crosses this insertion point only when the row being pushed down is itself a
		// continuation; if `at` is past the end, or the cell there starts its own span, nothing
		// is being split.
		const displaced = grid[at]?.[col]
		if (displaced && isVMerge(displaced)) {
			const originRow = vRunOrigin(grid, at, col)
			const origin = grid[originRow]?.[col]
			if (origin) setSpan(origin, 'rowSpan', rowSpanOf(origin) + 1)
			setAttr(tc, 'vMerge', '1')
			// The continuation must match the origin's horizontal extent too, or the row's cells
			// stop lining up with the grid columns.
			if (isHMerge(displaced)) setAttr(tc, 'hMerge', '1')
			const span = gridSpanOf(displaced)
			if (span > 1) setSpan(tc, 'gridSpan', span)
		}
		tr.appendChild(tc)
	}

	// `CT_Table` sequences tblPr, tblGrid, then the rows, so an insert is always relative to
	// an existing `a:tr` (or appended, which lands after tblGrid either way).
	const before = rows[at] ?? null
	tbl.insertBefore(tr, before)
	return tr
}

/**
 * Remove the row at `index`.
 *
 * Two span cases, and they pull in opposite directions. A cell in this row that *continues*
 * a span from above shortens that span by one. A cell in this row that *starts* one cannot
 * simply take its span away — the rows below still hold its continuations — so the first
 * continuation is promoted to origin and inherits the remaining extent. Its content is gone
 * with the row, which is inherent to removing a row rather than a choice made here.
 */
export function removeRow(tbl: Element, index: number): void {
	const rows = rowsOf(tbl)
	const at = checkIndex(index, rows.length, 'row', false)
	const grid = gridOfCells(tbl)
	const row = grid[at] ?? []

	for (let col = 0; col < row.length; col++) {
		const tc = row[col]
		if (!tc) continue
		if (isVMerge(tc)) {
			const origin = grid[vRunOrigin(grid, at, col)]?.[col]
			if (origin) setSpan(origin, 'rowSpan', rowSpanOf(origin) - 1)
			continue
		}
		const span = rowSpanOf(tc)
		if (span > 1) {
			const heir = grid[at + 1]?.[col]
			if (heir) {
				removeAttr(heir, 'vMerge')
				setSpan(heir, 'rowSpan', span - 1)
				// The promoted cell keeps whatever horizontal extent the origin had, so the row it
				// now leads still matches the grid.
				setSpan(heir, 'gridSpan', gridSpanOf(tc))
			}
		}
	}

	const tr = rows[at]
	if (tr) tbl.removeChild(tr)
}

/**
 * Insert a column at `index` (default: append), `widthEmu` wide.
 *
 * Mirrors {@link insertRow}'s split case on the other axis: when the cell currently at
 * `index` is an `hMerge` continuation, the insertion point falls inside a horizontal span,
 * so the span widens by one and the new cell joins it as another continuation.
 */
export function insertColumn(tbl: Element, index?: number, widthEmu = EMU_PER_INCH): Element {
	const grid = gridOf(tbl)
	const cols = getElements(grid, 'a:gridCol')
	const at = index === undefined ? cols.length : checkIndex(index, cols.length, 'column', true)
	const doc = ownerDocumentOf(tbl)

	const gridCol = createElement(doc, 'a:gridCol')
	setAttr(gridCol, 'w', String(Math.round(widthEmu)))
	grid.insertBefore(gridCol, cols[at] ?? null)

	for (const tr of rowsOf(tbl)) {
		const cells = cellsOf(tr)
		const tc = makeCell(doc)
		const displaced = cells[at]
		if (displaced && isHMerge(displaced)) {
			const origin = cells[hRunOrigin(cells, at)]
			if (origin) setSpan(origin, 'gridSpan', gridSpanOf(origin) + 1)
			setAttr(tc, 'hMerge', '1')
			// Same reasoning as the row case: match the vertical extent so the column still
			// lines up down the table.
			if (isVMerge(displaced)) setAttr(tc, 'vMerge', '1')
			const span = rowSpanOf(displaced)
			if (span > 1) setSpan(tc, 'rowSpan', span)
		}
		tr.insertBefore(tc, cells[at] ?? firstChild(tr, 'a:extLst'))
	}
	return gridCol
}

/**
 * Remove the column at `index`.
 *
 * When the column falls inside a horizontal span, the span shrinks by one and a
 * *continuation* is removed rather than the origin — so the merged region narrows but keeps
 * its content. Only a column whose cell stands alone loses that cell outright.
 */
export function removeColumn(tbl: Element, index: number): void {
	const grid = gridOf(tbl)
	const cols = getElements(grid, 'a:gridCol')
	const at = checkIndex(index, cols.length, 'column', false)
	const col = cols[at]
	if (col) grid.removeChild(col)

	for (const tr of rowsOf(tbl)) {
		const cells = cellsOf(tr)
		const originIdx = hRunOrigin(cells, at)
		const origin = cells[originIdx]
		if (!origin) continue
		const span = gridSpanOf(origin)
		if (span > 1) {
			setSpan(origin, 'gridSpan', span - 1)
			// Drop the run's last continuation: it carries no content, so removing it costs
			// nothing, while removing the origin would take the region's text with it.
			const victim = cells[originIdx + span - 1]
			if (victim) tr.removeChild(victim)
			continue
		}
		const tc = cells[at]
		if (tc) tr.removeChild(tc)
	}
}

/**
 * Merge the rectangle `(row1, col1)`–`(row2, col2)` into one cell.
 *
 * The top-left cell becomes the origin and keeps its content; every other cell in the
 * rectangle becomes a covered cell — flagged, stripped of its own span attributes, and
 * emptied, since a covered cell is never rendered.
 *
 * A rectangle whose boundary **cuts through** an existing merge is rejected rather than
 * silently widened. Widening would be the friendlier-looking choice and the wrong one: the
 * caller asked for a specific region, and quietly returning a different one is how a layout
 * ends up subtly wrong with nothing to point at.
 */
export function mergeCells(tbl: Element, row1: number, col1: number, row2: number, col2: number): void {
	const grid = gridOfCells(tbl)
	const rowCount = grid.length
	const colCount = getElements(gridOf(tbl), 'a:gridCol').length
	const r1 = Math.min(row1, row2)
	const r2 = Math.max(row1, row2)
	const c1 = Math.min(col1, col2)
	const c2 = Math.max(col1, col2)
	checkIndex(r1, rowCount, 'row', false)
	checkIndex(r2, rowCount, 'row', false)
	checkIndex(c1, colCount, 'column', false)
	checkIndex(c2, colCount, 'column', false)
	if (r1 === r2 && c1 === c2) {
		throw new InvalidOptionError(
			'table/merge-range-invalid',
			'A merge needs at least two cells; the given range covers one'
		)
	}

	// Every span touching the rectangle must lie entirely inside it.
	for (let r = r1; r <= r2; r++) {
		for (let c = c1; c <= c2; c++) {
			const tc = grid[r]?.[c]
			if (!tc) continue
			const originRow = vRunOrigin(grid, r, c)
			const originCol = hRunOrigin(grid[r] ?? [], c)
			const origin = grid[originRow]?.[originCol]
			if (!origin) continue
			if (originRow < r1 || originCol < c1) {
				throw new InvalidOptionError(
					'table/merge-range-invalid',
					`The range (${r1},${c1})-(${r2},${c2}) starts inside an existing merged cell at (${originRow},${originCol}); unmerge it first`
				)
			}
			if (originRow + rowSpanOf(origin) - 1 > r2 || originCol + gridSpanOf(origin) - 1 > c2) {
				throw new InvalidOptionError(
					'table/merge-range-invalid',
					`The range (${r1},${c1})-(${r2},${c2}) cuts through an existing merged cell at (${originRow},${originCol}); unmerge it first`
				)
			}
		}
	}

	const origin = grid[r1]?.[c1]
	if (!origin) return
	setSpan(origin, 'gridSpan', c2 - c1 + 1)
	setSpan(origin, 'rowSpan', r2 - r1 + 1)
	removeAttr(origin, 'hMerge')
	removeAttr(origin, 'vMerge')

	for (let r = r1; r <= r2; r++) {
		for (let c = c1; c <= c2; c++) {
			if (r === r1 && c === c1) continue
			const tc = grid[r]?.[c]
			if (!tc) continue
			removeAttr(tc, 'gridSpan')
			removeAttr(tc, 'rowSpan')
			if (c > c1) setAttr(tc, 'hMerge', '1')
			else removeAttr(tc, 'hMerge')
			if (r > r1) setAttr(tc, 'vMerge', '1')
			else removeAttr(tc, 'vMerge')
			emptyCellText(tc)
		}
	}
}

/**
 * Split the merged cell whose origin is `(row, col)` back into individual cells.
 * The origin keeps its content; the cells it covered come back empty, which is what they
 * already were.
 */
export function unmergeCell(tbl: Element, row: number, col: number): void {
	const grid = gridOfCells(tbl)
	checkIndex(row, grid.length, 'row', false)
	checkIndex(col, getElements(gridOf(tbl), 'a:gridCol').length, 'column', false)
	const origin = grid[row]?.[col]
	if (!origin) return
	if (isHMerge(origin) || isVMerge(origin)) {
		throw new InvalidOptionError(
			'table/merge-range-invalid',
			`The cell at (${row},${col}) is a covered cell, not a merge origin; unmerge the origin at (${vRunOrigin(grid, row, col)},${hRunOrigin(grid[row] ?? [], col)}) instead`
		)
	}
	const rowSpan = rowSpanOf(origin)
	const gridSpan = gridSpanOf(origin)
	if (rowSpan === 1 && gridSpan === 1) return

	removeAttr(origin, 'gridSpan')
	removeAttr(origin, 'rowSpan')
	for (let r = row; r < row + rowSpan; r++) {
		for (let c = col; c < col + gridSpan; c++) {
			if (r === row && c === col) continue
			const tc = grid[r]?.[c]
			if (!tc) continue
			removeAttr(tc, 'hMerge')
			removeAttr(tc, 'vMerge')
		}
	}
}

/** Reduce a cell's text body to a single empty paragraph. */
function emptyCellText(tc: Element): void {
	const txBody = firstChild(tc, 'a:txBody')
	if (!txBody) return
	for (const p of getElements(txBody, 'a:p')) txBody.removeChild(p)
	txBody.appendChild(createElement(ownerDocumentOf(txBody), 'a:p'))
}
