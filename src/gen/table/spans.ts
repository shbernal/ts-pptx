/**
 * ts-pptx: table `colspan`/`rowspan` range checking.
 *
 * The spans decide two allocations, on two different paths, and both used to trust the caller:
 * the merge grid's `new Array(colspan - 1)` in `gen/slide/objects/table.ts`, and the auto-pager's
 * per-column depth array, sized from a column count that is a sum of colspans
 * (`gen/table/autopage.ts`). `new Array(4294967294).fill(undefined)` is not a slow path — V8
 * aborts on the allocation and there is no exception to catch — so a table built with
 * `autoPage: true` and one without needed the same guard, which is why it lives here rather than
 * in either of them.
 *
 * This is the write side, so the bad number comes from the calling program rather than from a
 * hostile file, which makes it far less serious than the read-side twin in the chart point caches
 * (`read/oxml/point-cache.ts`, the model this copies). It is still the project's stated line: warn
 * on out-of-range input rather than emit a degenerate result.
 */

import type { TableCell } from '../../types/index.js'
import { warn } from '../../diagnostics.js'

/**
 * The largest `colspan`/`rowspan` this library will honour.
 *
 * Unlike the read side's point-cache ceiling there is no format number to borrow: `a:tc@gridSpan`
 * is a bare `xsd:int`. So this is a sanity bound rather than a schema one, and it is set an order
 * of magnitude past PowerPoint's own maximum table (its Insert Table dialog stops at 75 columns
 * and 75 rows), which is what makes it something no authored table can reach. What it stops is
 * the other end, where the allocation takes the host process with it.
 */
export const MAX_TABLE_SPAN = 1000

/**
 * Validate one cell's `colspan`/`rowspan`, warning and falling back to `1` when it is not a whole
 * number in `1..MAX_TABLE_SPAN`.
 *
 * `undefined` is the common case and means "no span" — it warns about nothing and returns `1`.
 * Everything else is caller error: `NaN`, a negative, a fraction and a span past the ceiling all
 * describe a table that cannot be built (a fractional span silently truncates, a negative one
 * shifts every column after it, an enormous one is the allocation above).
 * @param value - the raw option value
 * @param kind - `'colspan'` or `'rowspan'`, for the diagnostic
 * @returns the span to use, at least 1
 */
export function resolveSpan(value: number | undefined, kind: 'colspan' | 'rowspan'): number {
	if (value === undefined) return 1
	if (!Number.isInteger(value) || value < 1 || value > MAX_TABLE_SPAN) {
		warn(
			'table/span-out-of-range',
			`table cell: ${kind} must be a whole number between 1 and ${MAX_TABLE_SPAN}; got ${String(value)}. Using 1.`
		)
		return 1
	}
	return value
}

/**
 * Re-clone `rows` with every out-of-range `colspan`/`rowspan` corrected to `1`.
 *
 * Runs once, up front, so the column count, the `_hmerge`/`_vmerge` dummies, the auto-pager's
 * grid and the emitted `gridSpan`/`rowSpan` attributes all read the same number — checking at
 * each of those sites instead would warn several times per bad cell and could still disagree.
 *
 * A cell whose spans are fine is passed through by identity, not copied: `_spanOrigin` links the
 * merge dummies back to the origin *object*, and the all-valid path is every real table, which
 * must stay allocation-free and byte-identical.
 * @param rows - the rows to check; row arrays are never mutated
 * @returns the same rows with only the offending cells replaced
 */
export function withCheckedSpans(rows: TableCell[][]): TableCell[][] {
	return rows.map((cells) =>
		cells.map((cell) => {
			const opts = cell?.options
			if (!opts) return cell
			const colspan = resolveSpan(opts.colspan, 'colspan')
			const rowspan = resolveSpan(opts.rowspan, 'rowspan')
			const colspanOk = opts.colspan === undefined || opts.colspan === colspan
			const rowspanOk = opts.rowspan === undefined || opts.rowspan === rowspan
			if (colspanOk && rowspanOk) return cell
			return {
				...cell,
				options: {
					...opts,
					...(opts.colspan === undefined ? {} : { colspan }),
					...(opts.rowspan === undefined ? {} : { rowspan }),
				},
			}
		})
	)
}

/**
 * Whether any cell in the grid carries a hyperlink — on the cell itself, or on any run of a
 * cell whose `text` is a run array.
 *
 * `addTable` uses this to decide whether to stand its black text default down: the default is
 * direct formatting on the table, so it would paint the words *after* a link black rather than
 * letting them follow the link colour. The test used to be
 * `JSON.stringify({ arrRows }).includes('hyperlink')`, which matched the literal word anywhere
 * in the grid's *content* — a cell reading "see the hyperlink docs" suppressed the default for
 * the whole table — and serialized every cell on every `addTable` to ask.
 * @param rows - the table's rows
 */
export function tableHasHyperlink(rows: TableCell[][]): boolean {
	return rows.some((cells) =>
		cells.some((cell) => {
			if (!cell) return false
			if (cell.options?.hyperlink) return true
			if (!Array.isArray(cell.text)) return false
			return cell.text.some((run) => !!run?.options?.hyperlink)
		})
	)
}
