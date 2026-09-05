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

import type { TableCellInternal } from '../../types/internal.js'
import { resolveSpan } from './spans.js'

/**
 * How many grid columns a table has: the sum of the first row's colspans.
 *
 * The count cannot come from `rows[0].length` -- a colspan makes one cell occupy several
 * columns -- and it is read off the first row alone, which is the writer's standing rule: a
 * later row wider than the first is ragged input, and the grid it is measured, paged and
 * emitted against is the one `a:tblGrid` declares.
 */
export function tableColCount(rows: TableCellInternal[][]): number {
	const first = rows[0]
	return first ? first.reduce((n, c) => n + resolveSpan(c?.options?.colspan, 'colspan'), 0) : 0
}

/** The bookkeeping a row-major grid walk needs to route around rowspans; see {@link createRowSpanOccupancy}. */
interface RowSpanOccupancy {
	/** The first column at or after `col` that no rowspan from above is holding. */
	nextFree: (col: number) => number
	/** Record that `colSpan` columns from `col` are held for `rowSpan` rows, this one included. */
	hold: (col: number, colSpan: number, rowSpan: number) => void
	/** Close the current row: every held column has one fewer row left to hold. */
	endRow: () => void
}

/**
 * The rowspan-occupancy rule, stated once.
 *
 * A row states only the cells it *starts*: a column held by a `rowspan` from above is filled
 * without the row mentioning it, so a cell's position in its row array is not its column. Tracking
 * that is three operations -- skip what is held, mark what this cell holds, age the holds by a row
 * -- and getting the ageing off by one silently shifts every column below.
 *
 * It has two callers asking different questions on different input. {@link walkTableGrid} places
 * authored `TableCell`s into a grid whose width `a:tblGrid` already declares; `measureGridColumns`
 * (`gen/table/html-dom.ts`) measures how wide an imported HTML table's grid *is*, reading spans off
 * HTML attributes, with no width known up front. Neither can be expressed as the other without
 * giving one of them a mode it does not want -- but the rule underneath is the same rule, and it is
 * here rather than in both.
 *
 * The array is sparse and unbounded on purpose: a caller that knows its width clamps what it
 * `hold`s, and then nothing past that width is ever held, so an unbounded `nextFree` returns the
 * same answer a bounded one would.
 */
export function createRowSpanOccupancy(): RowSpanOccupancy {
	// held[c] = rows still covered by a rowspan started at or above the current row.
	const held: number[] = []
	return {
		nextFree(col) {
			let free = col
			while ((held[free] ?? 0) > 0) free++
			return free
		},
		hold(col, colSpan, rowSpan) {
			for (let idx = 0; idx < colSpan; idx++) held[col + idx] = rowSpan
		},
		endRow() {
			for (let idx = 0; idx < held.length; idx++) {
				const rows = held[idx] ?? 0
				if (rows > 0) held[idx] = rows - 1
			}
		},
	}
}

/** A placed (non-merged origin) cell yielded by {@link walkTableGrid}. */
export interface GridPlacement {
	cell: TableCellInternal
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
export function* walkTableGrid(rows: TableCellInternal[][], numCols: number): Generator<GridPlacement> {
	const occupancy = createRowSpanOccupancy()
	for (let r = 0; r < rows.length; r++) {
		const row = rows[r]
		if (!row) continue
		let col = 0
		for (const cell of row) {
			col = occupancy.nextFree(col)
			if (col >= numCols) break
			const colspan = resolveSpan(cell?.options?.colspan, 'colspan')
			const rowspan = Math.min(resolveSpan(cell?.options?.rowspan, 'rowspan'), rows.length - r)
			const colStart = col
			// Clamped to the declared width, which is also what keeps `nextFree` unbounded and still
			// correct: nothing past `numCols` is ever held, so it never has a reason to walk past it.
			const colEnd = Math.min(colStart + colspan, numCols)
			occupancy.hold(colStart, colEnd - colStart, rowspan)
			col = colEnd
			yield { cell, row: r, col: colStart, rowSpan: rowspan, colSpan: colEnd - colStart }
		}
		occupancy.endRow()
	}
}
