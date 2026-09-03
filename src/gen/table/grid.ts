/**
 * ts-pptx: a table's cell grid.
 *
 * Rows arrive lopsided -- `row1: [A, B, C]`, `row2: [D]` when the first two columns are held by
 * a rowspan from above -- so a cell's position in its row array is not its column. Four places
 * needed to know where a cell actually lands (the auto-pager, twice, in one function; the
 * measured-fit pass; the table emitter) and each walked the grid itself, tracking rowspan
 * occupancy in its own array with its own off-by-one convention.
 *
 * This is that traversal, once. It lives beside {@link resolveSpan} rather than in `measure/`
 * because the grid is a fact about the authored table, not a measurement of it -- the pager and
 * the emitter want it as much as the fitter does, and `measure/table-fit.ts` was already
 * reaching into `gen/table/` for the span reader it is built on.
 */

import type { TableCell } from '../../types/index.js'
import { resolveSpan } from './spans.js'

/**
 * How many grid columns a table has: the sum of the first row's colspans.
 *
 * The count cannot come from `rows[0].length` -- a colspan makes one cell occupy several
 * columns -- and it is read off the first row alone, which is the writer's standing rule: a
 * later row wider than the first is ragged input, and the grid it is measured, paged and
 * emitted against is the one `a:tblGrid` declares.
 */
export function tableColCount(rows: TableCell[][]): number {
	const first = rows[0]
	return first ? first.reduce((n, c) => n + resolveSpan(c?.options?.colspan, 'colspan'), 0) : 0
}

/** A placed (non-merged origin) cell yielded by {@link walkTableGrid}. */
export interface GridPlacement {
	cell: TableCell
	/** Zero-based grid row of the cell's top-left origin. */
	row: number
	/** Zero-based grid column of the cell's top-left origin. */
	col: number
	/** Rows spanned, clamped to the available rows. */
	rowSpan: number
	/** Columns spanned, clamped to the grid width. */
	colSpan: number
}

/**
 * Walk a table's cell grid in row-major order, resolving each authored cell to its
 * grid origin (`row`/`col`) and clamped colspan/rowspan. Tracks rowspan occupancy so
 * a cell never lands beneath one spanned from above (grid build). This is
 * the single traversal shared by the measured-fit shrink pass and
 * {@link computeTableLayout}, so cell placement cannot drift between them.
 */
export function* walkTableGrid(rows: TableCell[][], numCols: number): Generator<GridPlacement> {
	// occupied[c] = rows still covered by a rowspan started above (incl. current row).
	const occupied = new Array<number>(numCols).fill(0)
	for (let r = 0; r < rows.length; r++) {
		const row = rows[r]
		if (!row) continue
		let col = 0
		for (const cell of row) {
			while (col < numCols && (occupied[col] ?? 0) > 0) col++
			if (col >= numCols) break
			const colspan = resolveSpan(cell?.options?.colspan, 'colspan')
			const rowspan = Math.min(resolveSpan(cell?.options?.rowspan, 'rowspan'), rows.length - r)
			const colStart = col
			const colEnd = Math.min(colStart + colspan, numCols)
			for (let c = colStart; c < colEnd; c++) occupied[c] = rowspan
			col = colEnd
			yield { cell, row: r, col: colStart, rowSpan: rowspan, colSpan: colEnd - colStart }
		}
		for (let c = 0; c < numCols; c++) {
			const cur = occupied[c] ?? 0
			if (cur > 0) occupied[c] = cur - 1
		}
	}
}
