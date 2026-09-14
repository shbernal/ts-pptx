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
 * 3. A covered cell's own span attributes, where it has any, agree with its origin's.
 *
 * The third is looser than it could be because a merge reaches this module in two forms.
 * PowerPoint repeats the spans on covered cells: the covered cells of a region's first row carry
 * its `@rowSpan`, and those of its first column its `@gridSpan`, and inserting or deleting a row
 * or column through the region rewrites them (`test/read/fixtures/table-merge-encoding.pptx`).
 * This library's writer and {@link mergeCells} write that form too, but another producer may
 * leave covered cells without span attributes. A table in either form has to take every edit, so no edit reads
 * a span off a covered cell. {@link regionAt} resolves a cell to its region's origin and the
 * spans come from there, and an edit that changes a region's extent rewrites the spans its
 * covered cells already carry, so the region stays in the form it was written in.
 *
 * Breaking any of them produces a table PowerPoint reports as a corrupt file rather than as
 * a bad edit, which is why the span bookkeeping here is explicit rather than incidental.
 */
import { InvalidOptionError } from '../../errors.js'
import { EMU_PER_INCH } from '../../units.js'
import {
	attr,
	boolAttr,
	createElement,
	firstChild,
	getElements,
	numberValue,
	ownerDocumentOf,
	removeAttr,
	setAttr,
	type Document,
	type Element,
} from '../oxml/dom.js'
import { checkPositiveEmu } from './coords.js'

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
	return numberValue(attr(tc, 'gridSpan')) ?? 1
}

/** A cell's `@rowSpan`, defaulting to 1. */
function rowSpanOf(tc: Element): number {
	return numberValue(attr(tc, 'rowSpan')) ?? 1
}

/** Whether a cell is the covered half of a horizontal merge. */
export function isHMerge(tc: Element): boolean {
	return boolAttr(tc, 'hMerge') === true
}

/** Whether a cell is the covered half of a vertical merge. */
export function isVMerge(tc: Element): boolean {
	return boolAttr(tc, 'vMerge') === true
}

/**
 * Set a span attribute, removing it when the span collapses to 1 — the schema default, and
 * what PowerPoint itself writes for an unmerged cell.
 */
function setSpan(tc: Element, name: 'gridSpan' | 'rowSpan', value: number): void {
	if (value <= 1) removeAttr(tc, name)
	else setAttr(tc, name, String(value))
}

/** A merged region, or a lone cell as a region of one: where its origin is, and how far it reaches. */
interface Region {
	readonly originRow: number
	readonly originCol: number
	readonly rowSpan: number
	readonly gridSpan: number
}

/**
 * The region the cell at `(row, col)` belongs to.
 *
 * Walks up through `vMerge` cells to the origin's row, then left through `hMerge` cells to the
 * origin itself, and takes the extent from the origin alone, since a covered cell may or may not
 * repeat it.
 */
function regionAt(grid: Element[][], row: number, col: number): Region {
	let originRow = row
	while (originRow > 0) {
		const cell = grid[originRow]?.[col]
		if (!cell || !isVMerge(cell)) break
		originRow -= 1
	}
	let originCol = col
	while (originCol > 0) {
		const cell = grid[originRow]?.[originCol]
		if (!cell || !isHMerge(cell)) break
		originCol -= 1
	}
	const origin = grid[originRow]?.[originCol]
	return {
		originRow,
		originCol,
		rowSpan: origin ? rowSpanOf(origin) : 1,
		gridSpan: origin ? gridSpanOf(origin) : 1,
	}
}

/** The distinct regions the given cells belong to, each once, in the order first met. */
function regionsAt(grid: Element[][], positions: Iterable<readonly [number, number]>): Region[] {
	const regions = new Map<string, Region>()
	for (const [row, col] of positions) {
		const region = regionAt(grid, row, col)
		const key = `${region.originRow},${region.originCol}`
		if (!regions.has(key)) regions.set(key, region)
	}
	return [...regions.values()]
}

/** The regions crossing row `row`. */
function regionsInRow(grid: Element[][], row: number): Region[] {
	return regionsAt(
		grid,
		(grid[row] ?? []).map((_, col) => [row, col] as const)
	)
}

/** The regions crossing column `col`. */
function regionsInColumn(grid: Element[][], col: number): Region[] {
	return regionsAt(
		grid,
		grid.map((_, row) => [row, col] as const)
	)
}

/** Every cell of `region` except its origin. */
function coveredCells(grid: Element[][], region: Region): Element[] {
	const out: Element[] = []
	for (let r = region.originRow; r < region.originRow + region.rowSpan; r++) {
		for (let c = region.originCol; c < region.originCol + region.gridSpan; c++) {
			if (r === region.originRow && c === region.originCol) continue
			const tc = grid[r]?.[c]
			if (tc) out.push(tc)
		}
	}
	return out
}

/**
 * Give a region a new `rowSpan` or `gridSpan`: on its origin, and on each covered cell of it that
 * already carries that attribute, so the region keeps the form it was written in.
 */
function setRegionSpan(grid: Element[][], region: Region, name: 'gridSpan' | 'rowSpan', value: number): void {
	const origin = grid[region.originRow]?.[region.originCol]
	if (origin) setSpan(origin, name, value)
	for (const tc of coveredCells(grid, region)) {
		if (attr(tc, name) !== null) setSpan(tc, name, value)
	}
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

/**
 * A fresh covered cell for a region an insertion widens, flagged and spanned exactly as `covered`,
 * the cell of the same region it is inserted beside.
 */
function coveredCellLike(doc: Document, covered: Element): Element {
	const tc = makeCell(doc)
	for (const name of ['gridSpan', 'rowSpan', 'hMerge', 'vMerge']) {
		const value = attr(covered, name)
		if (value !== null) setAttr(tc, name, value)
	}
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
 * The subtle case is inserting **through** a vertical merge. A region that started above the row
 * being pushed down straddles the insertion point, so the new row must continue it rather than
 * interrupt it: the region grows by one row, once however many columns it spans, and the new
 * cells under it are covered cells. Interrupting it instead would leave an origin claiming more
 * rows than it has continuations, which is the corrupt-file case.
 */
export function insertRow(tbl: Element, index?: number): Element {
	const rows = rowsOf(tbl)
	const at = index === undefined ? rows.length : checkIndex(index, rows.length, 'row', true)
	const grid = gridOfCells(tbl)
	const colCount = getElements(gridOf(tbl), 'a:gridCol').length
	const doc = ownerDocumentOf(tbl)

	for (const region of regionsInRow(grid, at)) {
		if (region.originRow < at) setRegionSpan(grid, region, 'rowSpan', region.rowSpan + 1)
	}

	const tr = createElement(doc, 'a:tr')
	// `@h` is required on CT_TableRow. Zero means "auto — as tall as the content needs",
	// which is the right default for a row nobody has sized.
	setAttr(tr, 'h', '0')
	for (let col = 0; col < colCount; col++) {
		// A pushed-down cell flagged `vMerge` is inside a region that started above, so the new
		// cell joins that region in the same form.
		const displaced = grid[at]?.[col]
		tr.appendChild(displaced && isVMerge(displaced) ? coveredCellLike(doc, displaced) : makeCell(doc))
	}

	// `CT_Table` sequences tblPr, tblGrid, then the rows, so an insert is always relative to
	// an existing `a:tr` (or appended, which lands after tblGrid either way).
	tbl.insertBefore(tr, rows[at] ?? null)
	return tr
}

/**
 * Remove the row at `index`.
 *
 * Two span cases, and they pull in opposite directions. A region that started above this row
 * shortens by one. A region that *starts* in this row and continues below cannot simply lose its
 * origin, since the rows below still hold its covered cells, so the next row takes over as its
 * first row and inherits the remaining extent. The origin's content is gone with the row, which is
 * inherent to removing a row rather than a choice made here.
 */
export function removeRow(tbl: Element, index: number): void {
	const rows = rowsOf(tbl)
	const at = checkIndex(index, rows.length, 'row', false)
	const grid = gridOfCells(tbl)

	for (const region of regionsInRow(grid, at)) {
		if (region.originRow < at) {
			setRegionSpan(grid, region, 'rowSpan', region.rowSpan - 1)
		} else if (region.rowSpan > 1) {
			promoteNextRow(grid, region)
		}
	}

	const tr = rows[at]
	if (tr) tbl.removeChild(tr)
}

/**
 * Make the second row of `region` its first, for a removal of the row its origin is in. The cell
 * under the origin becomes the origin, and each covered cell of the new first row takes on the
 * `@rowSpan` the removed row's cell above it carried, if any.
 */
function promoteNextRow(grid: Element[][], region: Region): void {
	const heirRow = region.originRow + 1
	for (let col = region.originCol; col < region.originCol + region.gridSpan; col++) {
		const removed = grid[region.originRow]?.[col]
		const heir = grid[heirRow]?.[col]
		if (!removed || !heir) continue
		removeAttr(heir, 'vMerge')
		if (col === region.originCol) {
			setSpan(heir, 'rowSpan', region.rowSpan - 1)
			setSpan(heir, 'gridSpan', region.gridSpan)
		} else if (attr(removed, 'rowSpan') !== null) {
			setSpan(heir, 'rowSpan', region.rowSpan - 1)
		}
	}
}

/**
 * Insert a column at `index` (default: append), `widthEmu` wide.
 *
 * Mirrors {@link insertRow}'s split case on the other axis: a region that started left of the
 * insertion point widens by one, and the new cells inside it are covered cells.
 * @throws {InvalidOptionError} when `widthEmu` is not a positive number
 */
export function insertColumn(tbl: Element, index?: number, widthEmu = EMU_PER_INCH): Element {
	const width = checkPositiveEmu(widthEmu, 'widthEmu')
	const tblGrid = gridOf(tbl)
	const cols = getElements(tblGrid, 'a:gridCol')
	const at = index === undefined ? cols.length : checkIndex(index, cols.length, 'column', true)
	const grid = gridOfCells(tbl)
	const doc = ownerDocumentOf(tbl)

	for (const region of regionsInColumn(grid, at)) {
		if (region.originCol < at) setRegionSpan(grid, region, 'gridSpan', region.gridSpan + 1)
	}

	const gridCol = createElement(doc, 'a:gridCol')
	setAttr(gridCol, 'w', String(width))
	tblGrid.insertBefore(gridCol, cols[at] ?? null)

	rowsOf(tbl).forEach((tr, row) => {
		const displaced = grid[row]?.[at]
		const tc = displaced && isHMerge(displaced) ? coveredCellLike(doc, displaced) : makeCell(doc)
		tr.insertBefore(tc, displaced ?? firstChild(tr, 'a:extLst'))
	})
	return gridCol
}

/**
 * Remove the column at `index`.
 *
 * A region the column crosses narrows by one, from its right-hand end: every covered column of a
 * region holds the same kind of cell, so taking the last one keeps the origin, and its content,
 * where it is. Only a column whose cell stands alone loses that cell outright.
 */
export function removeColumn(tbl: Element, index: number): void {
	const tblGrid = gridOf(tbl)
	const cols = getElements(tblGrid, 'a:gridCol')
	const at = checkIndex(index, cols.length, 'column', false)
	const grid = gridOfCells(tbl)

	const removed = grid.map((cells) => cells[at])
	for (const region of regionsInColumn(grid, at)) {
		if (region.gridSpan <= 1) continue
		setRegionSpan(grid, region, 'gridSpan', region.gridSpan - 1)
		const lastCol = region.originCol + region.gridSpan - 1
		for (let row = region.originRow; row < region.originRow + region.rowSpan; row++) {
			removed[row] = grid[row]?.[lastCol]
		}
	}

	const col = cols[at]
	if (col) tblGrid.removeChild(col)
	rowsOf(tbl).forEach((tr, row) => {
		const tc = removed[row]
		if (tc) tr.removeChild(tc)
	})
}

/**
 * Merge the rectangle `(row1, col1)`–`(row2, col2)` into one cell.
 *
 * The top-left cell becomes the origin and keeps its content; every other cell in the
 * rectangle becomes a covered cell — flagged, given the spans PowerPoint repeats on a
 * region's first row and column, and emptied, since a covered cell is never rendered.
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

	// Every region touching the rectangle must lie entirely inside it.
	for (let r = r1; r <= r2; r++) {
		for (let c = c1; c <= c2; c++) {
			const region = regionAt(grid, r, c)
			if (region.originRow < r1 || region.originCol < c1) {
				throw new InvalidOptionError(
					'table/merge-range-invalid',
					`The range (${r1},${c1})-(${r2},${c2}) starts inside an existing merged cell at (${region.originRow},${region.originCol}); unmerge it first`
				)
			}
			if (region.originRow + region.rowSpan - 1 > r2 || region.originCol + region.gridSpan - 1 > c2) {
				throw new InvalidOptionError(
					'table/merge-range-invalid',
					`The range (${r1},${c1})-(${r2},${c2}) cuts through an existing merged cell at (${region.originRow},${region.originCol}); unmerge it first`
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
			// What PowerPoint writes: the region's row span on the covered cells of its first row,
			// its column span on those of its first column, and neither on the cells inside.
			if (r === r1) setSpan(tc, 'rowSpan', r2 - r1 + 1)
			if (c === c1) setSpan(tc, 'gridSpan', c2 - c1 + 1)
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
 * already were, and with no span attributes of their own.
 */
export function unmergeCell(tbl: Element, row: number, col: number): void {
	const grid = gridOfCells(tbl)
	checkIndex(row, grid.length, 'row', false)
	checkIndex(col, getElements(gridOf(tbl), 'a:gridCol').length, 'column', false)
	const origin = grid[row]?.[col]
	if (!origin) return
	const region = regionAt(grid, row, col)
	if (isHMerge(origin) || isVMerge(origin)) {
		throw new InvalidOptionError(
			'table/merge-range-invalid',
			`The cell at (${row},${col}) is a covered cell, not a merge origin; unmerge the origin at (${region.originRow},${region.originCol}) instead`
		)
	}
	if (region.rowSpan === 1 && region.gridSpan === 1) return

	for (const tc of [origin, ...coveredCells(grid, region)]) {
		for (const name of ['gridSpan', 'rowSpan', 'hMerge', 'vMerge']) removeAttr(tc, name)
	}
}

/** Reduce a cell's text body to a single empty paragraph. */
function emptyCellText(tc: Element): void {
	const txBody = firstChild(tc, 'a:txBody')
	if (!txBody) return
	for (const p of getElements(txBody, 'a:p')) txBody.removeChild(p)
	txBody.appendChild(createElement(ownerDocumentOf(txBody), 'a:p'))
}
