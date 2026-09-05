/**
 * ts-pptx: measuring a table's row heights and cell fits
 *
 * A table is measured cell by cell against its resolved grid: each cell inherits the table's text
 * options where it sets none, gets its own margins resolved to EMU insets, and is laid out inside
 * the column span it actually occupies. A merged region is measured once, at its origin, across
 * the columns it covers — which is what `walkTableGrid` exists to hand out.
 */

import { DEF_FONT_SIZE } from '../constants-internal.js'
import { type GridPlacement, tableColCount, walkTableGrid } from '../gen/table/grid.js'
import { EMU_PER_POINT, emuToInches } from '../units.js'
import {
	autoPageLineHeightEmu,
	getSmartParseNumber,
	marginToEmu,
	resolveCellMarginsInches,
	resolveTableColWidthsEmu,
	resolveTableRowHeightEmu,
} from '../units-internal.js'
import { measureLayout, WIDTH_SAFETY_FACTOR, HEIGHT_SAFETY_FACTOR } from './text-fit.js'
import { makeRegistryResolver, type FontMetricsRegistry } from './font-metrics.js'
import { extractParagraphs, type RunOpts } from './paragraphs.js'
import { CELL_INHERITED_TEXT_KEYS } from '../gen/table/cell-inherit.js'
import type { TableCellInternal } from '../types/internal.js'
import type {
	Coord,
	Margin,
	PresLayout,
	TableCellLayout,
	TableCellProps,
	TableLayoutResult,
	TableProps,
} from '../types/index.js'

/**
 * Effective cell options: the cell's own values, with table-level values filled in where unset.
 *
 * The keys are `CELL_INHERITED_TEXT_KEYS` (`gen/table/cell-inherit.ts`), the half of the
 * emitter's list that decides how text lays out. This used to be its own list, claiming in its
 * docstring to mirror the emitter's while naming four keys the emitter did not.
 *
 * PowerPoint has no text-autofit for a table cell — `a:tcPr` carries no autofit, and the app
 * ignores a `normAutofit` inside one; rows auto-grow instead. So a cell's `fit: 'shrink'` is
 * honoured by baking a *reduced literal font size* onto its runs, which both PowerPoint and
 * LibreOffice render identically with no edit or resize. That is what makes these options worth
 * resolving here: the shrink is computed against the cell's effective text, not the table's.
 */
export function effectiveCellOpts(cellOpts: TableCellProps, tableOpts: RunOpts): RunOpts {
	const merged = { ...cellOpts } as RunOpts
	for (const k of CELL_INHERITED_TEXT_KEYS) {
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
	const arr = resolveCellMarginsInches(margin)
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
export function scaleCellFontSizes(cell: TableCellInternal, eff: RunOpts, f: number): void {
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

/**
 * The options {@link resolveTableGridEmu} reads.
 *
 * `cy` is on `ObjectOptions` rather than `TableProps` — it is the already-resolved EMU height
 * the auto-pager and the measured-fit pass stamp on — so the parameter is the intersection
 * rather than either type, and every caller's own bag satisfies it structurally.
 */
type TableGridOpts = Pick<TableProps, 'w' | 'h' | 'colW' | 'rowH'> & { cy?: Coord }

/** A table's resolved grid: how many columns, how wide each is, and how tall each row is. */
interface TableGridEmu {
	/** Grid columns, counting colspans. */
	numCols: number
	/** Per-column width in EMU, length {@link numCols}. */
	colWidthsEmu: number[]
	/** The table's own height in EMU, or `0` when it has none and every row is auto-height. */
	tableHeightEmu: number
	/** One row's pinned height in EMU, or `null` when the row grows to fit its content. */
	rowHeightEmu: (rowIndex: number) => number | null
}

/**
 * Resolve a table's grid the one way every consumer has to see it.
 *
 * Three sites derived these four things: `pptx.tableLayout()`, the export-time measured-fit
 * pass, and the table emitter (the height alone). They already disagreed — only the fit pass
 * read `cy`, the EMU height the auto-pager and the fit pass itself stamp onto the options, so
 * `addTable(rows, { cy })` with no `h` gave a file whose rows are pinned and a
 * `pptx.tableLayout()` that reported every row auto-height. That is the same drift the one
 * reading of `rowH` closed, arriving through a second option, which is why the answer is a
 * shared resolver rather than a third correction.
 *
 * `w` defaults to `75%` of the slide, matching what the emitter falls back to.
 * @param rows - the table's authored rows
 * @param opts - the table's options
 * @param presLayout - the presentation layout, for resolving percentages
 */
export function resolveTableGridEmu(
	rows: TableCellInternal[][],
	opts: TableGridOpts,
	presLayout: PresLayout
): TableGridEmu {
	const numRows = rows.length
	const numCols = tableColCount(rows)
	const cxEmu =
		opts.w != null ? getSmartParseNumber(opts.w, 'X', presLayout) : getSmartParseNumber('75%', 'X', presLayout)
	// `cy` is the already-resolved EMU height the auto-pager and `applyMeasuredFit` stamp on;
	// it is the table's height just as much as `h` is, and only one of the three readings had it.
	const tableHeightEmu =
		opts.h != null ? getSmartParseNumber(opts.h, 'Y', presLayout) : typeof opts.cy === 'number' ? opts.cy : 0
	return {
		numCols,
		colWidthsEmu: resolveTableColWidthsEmu(opts.colW, cxEmu, numCols),
		tableHeightEmu,
		rowHeightEmu: (rowIndex: number) => resolveTableRowHeightEmu(opts.rowH, rowIndex, tableHeightEmu, numRows),
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
	rows: TableCellInternal[][],
	opts: TableProps,
	presLayout: PresLayout,
	registry: FontMetricsRegistry
): TableLayoutResult {
	const empty: TableLayoutResult = { cells: [], widthIn: 0, heightIn: 0, heightExact: true }
	if (!rows || rows.length === 0 || !rows[0]) return empty
	const numRows = rows.length
	const grid = resolveTableGridEmu(rows, opts, presLayout)
	const numCols = grid.numCols
	if (!(numCols > 0)) return empty
	// Reuse TableProps as a bag of inheritable text props; only CELL_INHERITED_TEXT_KEYS are
	// read from it. Route through `unknown` because TableProps.columns (per-column cell
	// styling, TableCellProps[]) collides by name with TextPropsOptions.columns (text
	// column count, number) — the two never overlap here since `columns` isn't inherited.
	const o = opts as unknown as RunOpts

	const tableXEmu = opts.x != null ? getSmartParseNumber(opts.x, 'X', presLayout) : 0
	const tableYEmu = opts.y != null ? getSmartParseNumber(opts.y, 'Y', presLayout) : 0
	const colWidthsEmu = grid.colWidthsEmu

	// Prefix-sum column x offsets (length numCols+1): colXEmu[c] = left edge of column c.
	const colXEmu = new Array<number>(numCols + 1).fill(0)
	for (let c = 0; c < numCols; c++) colXEmu[c + 1] = (colXEmu[c] ?? 0) + (colWidthsEmu[c] ?? 0)

	// Explicit row height (EMU) or null when the row is auto-height — the same reading the writer
	// bakes and the export-time fit pass measures against.
	const explicitRowHEmu = grid.rowHeightEmu
	const resolve = makeRegistryResolver(registry)
	const defLineEmu = autoPageLineHeightEmu(DEF_FONT_SIZE)

	// Estimate a cell's content height (EMU) at its authored size, conservative/tall,
	// with a one-line floor. Mirrors measureText's inflated wrap + height safety.
	const estimateContentHeightEmu = (cell: TableCellInternal, eff: RunOpts, innerWidthPt: number): number => {
		const fontSizePt = Number(eff.fontSize ?? DEF_FONT_SIZE) || DEF_FONT_SIZE
		const oneLineEmu = autoPageLineHeightEmu(fontSizePt)
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
