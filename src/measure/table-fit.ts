/**
 * ts-pptx: measuring a table's row heights and cell fits
 *
 * A table is measured cell by cell against its resolved grid: each cell inherits the table's text
 * options where it sets none, gets its own margins resolved to EMU insets, and is laid out inside
 * the column span it actually occupies. A merged region is measured once, at its origin, across
 * the columns it covers — which is what `walkTableGrid` exists to hand out.
 */

import { DEF_CELL_MARGIN_IN, DEF_FONT_SIZE, LINEH_MODIFIER } from '../constants-internal.js'
import { EMU_PER_POINT, emuToInches } from '../units.js'
import {
	getSmartParseNumber,
	inch2Emu,
	marginToEmu,
	resolveTableColWidthsEmu,
	resolveTableRowHeightEmu,
} from '../units-internal.js'
import { measureLayout, WIDTH_SAFETY_FACTOR, HEIGHT_SAFETY_FACTOR } from './text-fit.js'
import { makeRegistryResolver, type FontMetricsRegistry } from './font-metrics.js'
import { extractParagraphs, type RunOpts } from './paragraphs.js'
import type {
	Margin,
	PresLayout,
	TableCell,
	TableCellLayout,
	TableCellProps,
	TableLayoutResult,
	TableProps,
} from '../types/index.js'

const CELL_INHERIT_KEYS = [
	'fontFace',
	'fontSize',
	'bold',
	'italic',
	'charSpacing',
	'align',
	'lineSpacing',
	'lineSpacingMultiple',
	'valign',
	'margin',
] as const

/** Effective cell options: the cell's own values, with table-level values filled in where unset. */
export function effectiveCellOpts(cellOpts: TableCellProps, tableOpts: RunOpts): RunOpts {
	const merged = { ...cellOpts } as RunOpts
	for (const k of CELL_INHERIT_KEYS) {
		if (merged[k] === undefined && tableOpts[k] !== undefined) (merged as Record<string, unknown>)[k] = tableOpts[k]
	}
	return merged
}

interface CellInsetsEmu {
	marL: number
	marR: number
	marT: number
	marB: number
}

/** Resolve a cell's margins to EMU insets, mirroring `gen/slide/objects/table.ts` (array is `[T,R,B,L]`, inches; see `marginToEmu`). */
export function resolveCellInsetsEmu(margin: Margin | undefined): CellInsetsEmu {
	let m: Margin = margin === 0 || margin ? margin : DEF_CELL_MARGIN_IN
	if (typeof m === 'number') m = [m, m, m, m]
	if (!Array.isArray(m) || m.length !== 4 || m.some((v) => typeof v !== 'number' || !Number.isFinite(v)))
		m = DEF_CELL_MARGIN_IN
	const arr = m
	return {
		marT: marginToEmu(arr[0]),
		marR: marginToEmu(arr[1]),
		marB: marginToEmu(arr[2]),
		marL: marginToEmu(arr[3]),
	}
}

/**
 * Bake a reduced font size onto a cell's runs by factor `f` (< 1). Clones every
 * options object before mutating: a plain-string cell shares the table's `opt`
 * object (`gen/define/`), so in-place mutation would corrupt every other such cell.
 */
export function scaleCellFontSizes(cell: TableCell, eff: RunOpts, f: number): void {
	const shrink = (sizePt: number): number => Math.floor(sizePt * f * 10) / 10 // floor: stay on the conservative (smaller) side
	const baseSize = Number(eff.fontSize ?? DEF_FONT_SIZE)
	cell.options = { ...cell.options, fontSize: shrink(baseSize) }
	if (Array.isArray(cell.text)) {
		cell.text = cell.text.map((run) =>
			run && typeof run === 'object' && typeof run.options?.fontSize === 'number'
				? { ...run, options: { ...run.options, fontSize: shrink(run.options.fontSize) } }
				: run
		)
	}
}

/** Grid column count of a table (sums the first row's colspans), mirroring `gen/slide/objects/table.ts`. */
export function tableColCount(rows: TableCell[][]): number {
	const first = rows[0]
	return first ? first.reduce((n, c) => n + (Number(c?.options?.colspan) || 1), 0) : 0
}

/** A placed (non-merged origin) cell yielded by {@link walkTableGrid}. */
interface GridPlacement {
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
			const colspan = Math.max(1, Number(cell?.options?.colspan) || 1)
			const rowspan = Math.min(Math.max(1, Number(cell?.options?.rowspan) || 1), rows.length - r)
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

/**
 * Compute per-cell geometry (inches) for a table laid out at `opts.x`/`y`/`w` — the
 * engine behind `pptx.tableLayout()`. Column widths come from the same
 * {@link resolveTableColWidthsEmu} the writer uses, so cell x/width are exact. Row
 * heights are exact when pinned by `rowH` (array or scalar) or table `h`; otherwise
 * each auto-height row is estimated with the same conservative (tall) text model as
 * {@link measureText} and flagged `heightExact: false`. Single-slide only —
 * `autoPage` paging across slides is not modeled. Rowspan cells do not drive a row's
 * estimated height (mirrors `gen/table/autopage.ts`, which exempts them from line-height growth).
 */
export function computeTableLayout(
	rows: TableCell[][],
	opts: TableProps,
	presLayout: PresLayout,
	registry: FontMetricsRegistry
): TableLayoutResult {
	const empty: TableLayoutResult = { cells: [], widthIn: 0, heightIn: 0, heightExact: true }
	if (!rows || rows.length === 0 || !rows[0]) return empty
	const numRows = rows.length
	const numCols = tableColCount(rows)
	if (!(numCols > 0)) return empty
	// Reuse TableProps as a bag of inheritable text props; only CELL_INHERIT_KEYS are
	// read from it. Route through `unknown` because TableProps.columns (per-column cell
	// styling, TableCellProps[]) collides by name with TextPropsOptions.columns (text
	// column count, number) — the two never overlap here since `columns` isn't inherited.
	const o = opts as unknown as RunOpts

	const tableXEmu = opts.x != null ? getSmartParseNumber(opts.x, 'X', presLayout) : 0
	const tableYEmu = opts.y != null ? getSmartParseNumber(opts.y, 'Y', presLayout) : 0
	const cxEmu =
		opts.w != null ? getSmartParseNumber(opts.w, 'X', presLayout) : getSmartParseNumber('75%', 'X', presLayout)
	const colWidthsEmu = resolveTableColWidthsEmu(opts.colW, cxEmu, numCols)

	// Prefix-sum column x offsets (length numCols+1): colXEmu[c] = left edge of column c.
	const colXEmu = new Array<number>(numCols + 1).fill(0)
	for (let c = 0; c < numCols; c++) colXEmu[c + 1] = (colXEmu[c] ?? 0) + (colWidthsEmu[c] ?? 0)

	const tableHeightEmu = opts.h != null ? getSmartParseNumber(opts.h, 'Y', presLayout) : 0
	// Explicit row height (EMU) or null when the row is auto-height — `resolveTableRowHeightEmu`,
	// the same reading the writer bakes and the export-time fit pass measures against.
	const explicitRowHEmu = (r: number): number | null => resolveTableRowHeightEmu(opts.rowH, r, tableHeightEmu, numRows)
	const resolve = makeRegistryResolver(registry)
	const defLineEmu = inch2Emu((DEF_FONT_SIZE * LINEH_MODIFIER) / 100)

	// Estimate a cell's content height (EMU) at its authored size, conservative/tall,
	// with a one-line floor. Mirrors measureText's inflated wrap + height safety.
	const estimateContentHeightEmu = (cell: TableCell, eff: RunOpts, innerWidthPt: number): number => {
		const fontSizePt = Number(eff.fontSize ?? DEF_FONT_SIZE) || DEF_FONT_SIZE
		const oneLineEmu = inch2Emu((fontSizePt * LINEH_MODIFIER) / 100)
		if (!(innerWidthPt > 0)) return oneLineEmu
		const paragraphs = extractParagraphs({ text: cell.text, options: eff })
		if (!paragraphs) return oneLineEmu
		const layout = measureLayout(paragraphs, innerWidthPt, resolve, 100, 0, WIDTH_SAFETY_FACTOR)
		if (layout === null) return oneLineEmu
		return Math.max(Math.round(layout.heightPt * HEIGHT_SAFETY_FACTOR * EMU_PER_POINT), oneLineEmu)
	}

	// PASS 1: place every origin cell and resolve each row's height.
	interface Placed {
		p: GridPlacement
		colStart: number
		colEnd: number
	}
	const placed: Placed[] = []
	const rowHeightsEmu = new Array<number>(numRows).fill(0)
	const rowExact = new Array<boolean>(numRows).fill(false)
	for (let r = 0; r < numRows; r++) {
		const ex = explicitRowHEmu(r)
		if (ex != null) {
			rowHeightsEmu[r] = ex
			rowExact[r] = true
		}
	}

	for (const p of walkTableGrid(rows, numCols)) {
		const colStart = p.col
		const colEnd = p.col + p.colSpan
		placed.push({ p, colStart, colEnd })

		// Only single-row, auto-height cells drive a row's estimated height.
		if (p.rowSpan === 1 && !rowExact[p.row]) {
			let widthEmu = 0
			for (let c = colStart; c < colEnd; c++) widthEmu += colWidthsEmu[c] ?? 0
			const eff = effectiveCellOpts(p.cell?.options ?? {}, o)
			const ins = resolveCellInsetsEmu(eff.margin)
			const innerWidthPt = (widthEmu - ins.marL - ins.marR) / EMU_PER_POINT
			const contentEmu = estimateContentHeightEmu(p.cell, eff, innerWidthPt) + ins.marT + ins.marB
			if (contentEmu > (rowHeightsEmu[p.row] ?? 0)) rowHeightsEmu[p.row] = contentEmu
		}
	}

	// A row touched only by rowspans (no single-row cell, no explicit height) still
	// needs a non-zero height: give it one default line. That height is this function's own
	// invention, so the row stops being exact — a caller told it was pinned would place the
	// next shape against a number nothing in the file agrees with.
	for (let r = 0; r < numRows; r++)
		if ((rowHeightsEmu[r] ?? 0) <= 0) {
			rowHeightsEmu[r] = defLineEmu
			rowExact[r] = false
		}

	// Prefix-sum row y offsets.
	const rowYEmu = new Array<number>(numRows + 1).fill(0)
	for (let r = 0; r < numRows; r++) rowYEmu[r + 1] = (rowYEmu[r] ?? 0) + (rowHeightsEmu[r] ?? 0)

	// PASS 2: emit one rect per origin cell.
	const cells: TableCellLayout[] = placed.map(({ p, colStart, colEnd }) => {
		const rowEnd = p.row + p.rowSpan
		let heightExact = true
		for (let rr = p.row; rr < rowEnd; rr++) if (!rowExact[rr]) heightExact = false
		return {
			row: p.row,
			col: p.col,
			rowSpan: p.rowSpan,
			colSpan: p.colSpan,
			xIn: emuToInches(tableXEmu + (colXEmu[colStart] ?? 0)),
			yIn: emuToInches(tableYEmu + (rowYEmu[p.row] ?? 0)),
			wIn: emuToInches((colXEmu[colEnd] ?? 0) - (colXEmu[colStart] ?? 0)),
			hIn: emuToInches((rowYEmu[rowEnd] ?? 0) - (rowYEmu[p.row] ?? 0)),
			heightExact,
		}
	})

	return {
		cells,
		widthIn: emuToInches(colXEmu[numCols] ?? 0),
		heightIn: emuToInches(rowYEmu[numRows] ?? 0),
		heightExact: rowExact.every(Boolean),
	}
}

/**
 * Build the `MetricsResolver` both the export pass and `measureText` use, so they
 * agree run-for-run: exact registered metrics → conservative heuristic for any
 * **named** face without exact metrics → `undefined` only for an unnamed
 * (theme-default) face that cannot be guessed. `onHeuristic` is called with each
 * named face that fell back to the heuristic (for the export pass's warn-once).
 */
