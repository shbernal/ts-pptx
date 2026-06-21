/**
 * PptxGenJS: Measured-fit serialization pass
 *
 * Bridges slide text objects to the line-break simulator / shrink solver
 * (`text-fit.ts`). Runs during async export, BEFORE `gen-xml` builds the body:
 * for each text box with `fit: 'shrink'` and a registered font, it computes the
 * `fontScale` PowerPoint would have baked and rewrites `options.fit` to the
 * explicit object form (`{ type:'shrink', fontScale }`) so `genXmlNormAutofit`
 * emits `<a:normAutofit fontScale=…/>`. Without metrics it leaves the bare flag
 * untouched (current behavior) and warns once. See `docs/measured-text-fit.md`.
 */

import { SLIDE_OBJECT_TYPES } from './core-enums.js'
import { DEF_CELL_MARGIN_IN, DEF_FONT_SIZE } from './core-enums.js'
import { EMU_PER_POINT, POINTS_PER_INCH } from './units.js'
import { getSmartParseNumber, inch2Emu, resolveTableColWidthsEmu, valToPts } from './gen-utils.js'
import { getHeuristicFontMetrics, type FontMetricsRegistry } from './font-metrics.js'
import { solveShrink, solveResize, measureLayout, WIDTH_SAFETY_FACTOR, HEIGHT_SAFETY_FACTOR, type FitBox, type FitParagraph, type FitRun, type MetricsResolver } from './text-fit.js'
import type { ISlideObject, Margin, MeasureTextOptions, ObjectOptions, PresSlideInternal, TableCell, TableCellProps, TextMeasurement, TextProps, TextPropsOptions } from './core-interfaces.js'

// PowerPoint's default text-frame insets (EMU): l/r = 0.1in, t/b = 0.05in.
const DEF_INS_LR_EMU = 91440
const DEF_INS_TB_EMU = 45720

const CRLF_RE = /\r*\n/g

type RunOpts = TextPropsOptions & ObjectOptions

/** Normalize `slideObj.text` (string | TextProps | TextProps[]) to a run list. */
function normalizeRuns (obj: ISlideObject): TextProps[] {
	const opts = (obj.options ?? {}) as ObjectOptions
	const text = obj.text as unknown
	if (text == null) return []
	if (typeof text === 'string' || typeof text === 'number') return [{ text: String(text), options: opts }]
	if (!Array.isArray(text) && typeof text === 'object' && 'text' in (text as object)) {
		const t = text as TextProps
		return [{ text: t.text, options: t.options ?? opts }]
	}
	if (Array.isArray(text)) return (text as TextProps[]).map(t => ({ text: t.text, options: t.options ?? opts }))
	return []
}

/** Build a measurable `FitParagraph[]` from a text object, or null if not measurable. */
function extractParagraphs (obj: ISlideObject): FitParagraph[] | null {
	const opts = (obj.options ?? {}) as RunOpts
	const runs = normalizeRuns(obj)
	if (runs.length === 0) return null
	return buildFitParagraphs(runs, opts)
}

/**
 * Convert a run list (+ box-level default options) into a measurable
 * `FitParagraph[]`, or null if empty. The single converter shared by the
 * export-time pass ({@link extractParagraphs}) and the public layout-time
 * `measureText` API, so a layout-time prediction and the baked export never drift.
 */
export function buildFitParagraphs (runs: TextProps[], opts: RunOpts): FitParagraph[] | null {
	if (runs.length === 0) return null

	// Expand "\n" inside a run into separate pieces, flagging the paragraph break
	// after each (mirrors gen-xml STEP 4). `breakLine` ends a paragraph too.
	interface Piece { text: string, options: RunOpts, breakAfter: boolean }
	const pieces: Piece[] = []
	for (const run of runs) {
		const ro = (run.options ?? opts) as RunOpts
		const raw = String(run.text ?? '').replace(CRLF_RE, '\n')
		if (raw.includes('\n')) {
			const lines = raw.split('\n')
			lines.forEach((line, i) => {
				const isLast = i === lines.length - 1
				pieces.push({ text: line, options: ro, breakAfter: isLast ? !!ro.breakLine : true })
			})
		} else {
			pieces.push({ text: raw, options: ro, breakAfter: !!ro.breakLine })
		}
	}

	const toRun = (p: Piece): FitRun => {
		const ro = p.options
		const sizePt = Number(ro.fontSize ?? opts.fontSize ?? DEF_FONT_SIZE)
		return {
			text: p.text,
			sizePt: Number.isFinite(sizePt) && sizePt > 0 ? sizePt : DEF_FONT_SIZE,
			bold: !!(ro.bold ?? opts.bold),
			italic: !!(ro.italic ?? opts.italic),
			fontFace: ro.fontFace ?? opts.fontFace,
			charSpacingPt: (ro.charSpacing ?? opts.charSpacing) || undefined,
		}
	}
	const toPara = (runsForPara: FitRun[], paraOpts: RunOpts): FitParagraph => {
		const lineSpacing = paraOpts.lineSpacing ?? opts.lineSpacing
		const lineSpacingMultiple = paraOpts.lineSpacingMultiple ?? opts.lineSpacingMultiple
		return {
			runs: runsForPara,
			lineSpacingPts: typeof lineSpacing === 'number' && lineSpacing > 0 ? lineSpacing : undefined,
			lineSpacingPct: typeof lineSpacingMultiple === 'number' && lineSpacingMultiple > 0 ? lineSpacingMultiple * 100 : 100,
			spaceBeforePts: Number(paraOpts.paraSpaceBefore ?? opts.paraSpaceBefore ?? 0) || 0,
			spaceAfterPts: Number(paraOpts.paraSpaceAfter ?? opts.paraSpaceAfter ?? 0) || 0,
		}
	}

	const paras: FitParagraph[] = []
	let cur: FitRun[] = []
	let curParaOpts: RunOpts = opts
	for (const piece of pieces) {
		cur.push(toRun(piece))
		curParaOpts = piece.options
		if (piece.breakAfter) {
			paras.push(toPara(cur, curParaOpts))
			cur = []
		}
	}
	if (cur.length > 0) paras.push(toPara(cur, curParaOpts))
	return paras.length > 0 ? paras : null
}

interface InsetsEmu { lIns: number, rIns: number, tIns: number, bIns: number }

/** Resolve text-frame insets (EMU): explicit `_bodyProp` (from `inset`) → `margin` → PowerPoint defaults. */
function resolveInsetsEmu (opts: RunOpts): InsetsEmu {
	const bp = opts._bodyProp ?? {}
	const margin = opts.margin
	let lIns = bp.lIns
	let rIns = bp.rIns
	let tIns = bp.tIns
	let bIns = bp.bIns
	if (lIns == null && rIns == null && tIns == null && bIns == null && margin != null) {
		if (Array.isArray(margin)) {
			tIns = valToPts(margin[0] ?? 0)
			rIns = valToPts(margin[1] ?? 0)
			bIns = valToPts(margin[2] ?? 0)
			lIns = valToPts(margin[3] ?? 0)
		} else if (typeof margin === 'number') {
			lIns = rIns = tIns = bIns = valToPts(margin)
		}
	}
	return {
		lIns: lIns ?? DEF_INS_LR_EMU,
		rIns: rIns ?? DEF_INS_LR_EMU,
		tIns: tIns ?? DEF_INS_TB_EMU,
		bIns: bIns ?? DEF_INS_TB_EMU,
	}
}

/** Resolve the inner box (shape minus insets) in points; null if degenerate. */
function computeBox (obj: ISlideObject, presLayout: PresSlideInternal['_presLayout']): { innerWidthPt: number, innerHeightPt: number } | null {
	const opts = (obj.options ?? {}) as RunOpts
	const wEmu = getSmartParseNumber(opts.w, 'X', presLayout)
	const hEmu = getSmartParseNumber(opts.h, 'Y', presLayout)
	if (!(wEmu > 0) || !(hEmu > 0)) return null

	const { lIns, rIns, tIns, bIns } = resolveInsetsEmu(opts)

	// wrap=none means no width wrapping (only hard breaks): unbounded line width.
	const wrap = (opts._bodyProp ?? {}).wrap !== false
	const innerWidthPt = wrap ? (wEmu - lIns - rIns) / EMU_PER_POINT : Infinity
	const innerHeightPt = (hEmu - tIns - bIns) / EMU_PER_POINT
	if (!(innerHeightPt > 0)) return null
	if (wrap && !(innerWidthPt > 0)) return null
	return { innerWidthPt, innerHeightPt }
}

/** Vertical-anchor share of a height change that moves the box top up (`off.y` shift). */
function anchorTopShareOfDelta (opts: RunOpts): number {
	// `_bodyProp.anchor` is the resolved valign ('t' | 'ctr' | 'b'); default 'ctr'.
	const anchor = (opts._bodyProp ?? {}).anchor
	if (anchor === 't') return 0 // grow downward — top fixed
	if (anchor === 'b') return 1 // grow upward — bottom fixed
	return 0.5 // centered growth (default)
}

// --- Table cells -----------------------------------------------------------
// PowerPoint has no text-autofit for table cells (`a:tcPr` carries no autofit and
// the app ignores `normAutofit` inside a cell — rows auto-grow instead). So a cell's
// `fit:'shrink'` is honored by baking a *reduced literal font size* onto its runs,
// which both PowerPoint and LibreOffice render identically with no edit/resize.

/** Text/format options that a cell inherits from the table when it sets none itself (mirrors gen-xml). */
const CELL_INHERIT_KEYS = ['fontFace', 'fontSize', 'bold', 'italic', 'charSpacing', 'align', 'lineSpacing', 'lineSpacingMultiple', 'valign', 'margin'] as const

/** Effective cell options: the cell's own values, with table-level values filled in where unset. */
function effectiveCellOpts (cellOpts: TableCellProps, tableOpts: RunOpts): RunOpts {
	const merged = { ...cellOpts } as RunOpts
	for (const k of CELL_INHERIT_KEYS) {
		if (merged[k] === undefined && tableOpts[k] !== undefined) (merged as Record<string, unknown>)[k] = tableOpts[k]
	}
	return merged
}

interface CellInsetsEmu { marL: number, marR: number, marT: number, marB: number }

/** Resolve a cell's margins to EMU insets, mirroring gen-xml (array is `[T,R,B,L]`; ≥1 ⇒ points, else inches). */
function resolveCellInsetsEmu (margin: Margin | undefined): CellInsetsEmu {
	let m: Margin = margin === 0 || margin ? margin : DEF_CELL_MARGIN_IN
	if (typeof m === 'number') m = [m, m, m, m]
	if (!Array.isArray(m) || m.length !== 4 || m.some(v => typeof v !== 'number' || !isFinite(v))) m = DEF_CELL_MARGIN_IN
	const arr = m as [number, number, number, number]
	const toEmu = arr[0] >= 1 ? valToPts : inch2Emu
	return { marT: toEmu(arr[0]), marR: toEmu(arr[1]), marB: toEmu(arr[2]), marL: toEmu(arr[3]) }
}

/**
 * Bake a reduced font size onto a cell's runs by factor `f` (< 1). Clones every
 * options object before mutating: a plain-string cell shares the table's `opt`
 * object (gen-objects), so in-place mutation would corrupt every other such cell.
 */
function scaleCellFontSizes (cell: TableCell, eff: RunOpts, f: number): void {
	const shrink = (sizePt: number): number => Math.floor(sizePt * f * 10) / 10 // floor: stay on the conservative (smaller) side
	const baseSize = Number(eff.fontSize ?? DEF_FONT_SIZE)
	cell.options = { ...(cell.options ?? {}), fontSize: shrink(baseSize) }
	if (Array.isArray(cell.text)) {
		cell.text = (cell.text as TableCell[]).map(run =>
			run && typeof run === 'object' && typeof run.options?.fontSize === 'number'
				? { ...run, options: { ...run.options, fontSize: shrink(run.options.fontSize) } }
				: run
		)
	}
}

/**
 * Build the `MetricsResolver` both the export pass and `measureText` use, so they
 * agree run-for-run: exact registered metrics → conservative heuristic for any
 * **named** face without exact metrics → `undefined` only for an unnamed
 * (theme-default) face that cannot be guessed. `onHeuristic` is called with each
 * named face that fell back to the heuristic (for the export pass's warn-once).
 */
export function makeRegistryResolver (registry: FontMetricsRegistry, onHeuristic?: (face: string) => void): MetricsResolver {
	return run => {
		const exact = registry.get(run.fontFace, run.bold, run.italic)
		if (exact) return exact
		if (typeof run.fontFace === 'string' && run.fontFace.length > 0) {
			onHeuristic?.(run.fontFace)
			return getHeuristicFontMetrics()
		}
		return undefined
	}
}

/** Map the public {@link MeasureTextOptions} onto the internal run-option shape. */
function measureOptsToRunOpts (opts: MeasureTextOptions): RunOpts {
	return {
		fontSize: opts.fontSize,
		fontFace: opts.fontFace,
		bold: opts.bold,
		italic: opts.italic,
		charSpacing: opts.charSpacing,
		lineSpacing: opts.lineSpacing,
		lineSpacingMultiple: opts.lineSpacingMultiple,
		paraSpaceBefore: opts.paraSpaceBefore,
		paraSpaceAfter: opts.paraSpaceAfter,
	} as RunOpts
}

const UNMEASURABLE: TextMeasurement = Object.freeze({
	heightIn: 0,
	lineCount: 0,
	widestLineIn: 0,
	measurable: false,
	fitsBox: () => false,
	shrinkScaleFor: () => 100,
})

/**
 * Layout-time text measurement against registered metrics — the public engine
 * behind `pptx.measureText()`. Uses the **same** calibrated wrap model, resolver
 * semantics, and conservative safety factors as the export-time bake
 * ({@link applyMeasuredFit} / {@link solveResize} / {@link solveShrink}), so a
 * layout-time prediction matches the value the export would bake.
 *
 * Synchronous: assumes metrics are pre-registered (lookup is sync). A named face
 * with no exact metrics silently uses the conservative heuristic (same as export);
 * an unnamed theme-default face returns `measurable: false`.
 */
export function measureText (registry: FontMetricsRegistry, text: string | TextProps[], opts: MeasureTextOptions): TextMeasurement {
	const runs: TextProps[] = typeof text === 'string' || typeof text === 'number'
		? [{ text: String(text) }]
		: Array.isArray(text) ? text : []
	const paragraphs = buildFitParagraphs(runs, measureOptsToRunOpts(opts))
	if (!paragraphs) return UNMEASURABLE

	const inset = opts.insetIn ?? 0
	const innerWidthPt = (opts.wIn - 2 * inset) * POINTS_PER_INCH
	const resolve = makeRegistryResolver(registry)

	// Conservative (tall) layout at full size, mirroring solveResize: inflate width
	// (earlier wrap) by WIDTH_SAFETY and the height by HEIGHT_SAFETY.
	const layout = measureLayout(paragraphs, innerWidthPt, resolve, 100, 0, WIDTH_SAFETY_FACTOR)
	if (layout === null) return UNMEASURABLE
	const heightPt = layout.heightPt * HEIGHT_SAFETY_FACTOR
	const heightIn = heightPt / POINTS_PER_INCH

	return {
		heightIn,
		lineCount: layout.lineCount,
		widestLineIn: layout.widestLineWidthPt / POINTS_PER_INCH,
		measurable: true,
		// Mirrors solveShrink's fit check at scale 100 (height already inflated).
		fitsBox: (hIn: number) => heightPt <= hIn * POINTS_PER_INCH,
		shrinkScaleFor: (hIn: number) => {
			const box: FitBox = { innerWidthPt, innerHeightPt: hIn * POINTS_PER_INCH }
			const outcome = solveShrink(paragraphs, box, resolve)
			if (outcome.kind === 'shrink') return outcome.result.fontScalePct
			return 100 // 'fits' (or, defensively, 'unmeasurable') → no shrink
		},
	}
}

/**
 * Apply measured fit across every slide. For each text box that opts in via
 * `fit:'shrink'` or `fit:'resize'` and whose font has registered metrics, this
 * bakes the computed result before the sync XML pass reads it:
 * - `'shrink'` → rewrites `options.fit` to the object form so `<a:normAutofit
 *   fontScale=…/>` is emitted (text renders pre-shrunk).
 * - `'resize'` → rewrites `options.h` (and `options.y` per vertical anchor) so the
 *   shape's `a:ext/@cy` is the height the text needs; the `<a:spAutoFit/>` marker is
 *   left in place (the renderer draws the baked `cy`).
 *
 * Safe to call with an empty registry (no-op). Warns once if any opted-in box could
 * not be measured (missing metrics) so overflow is not silently ignored.
 */
export function applyMeasuredFit (slides: PresSlideInternal[], registry: FontMetricsRegistry): void {
	if (registry.size === 0) return

	// A deck that registered *some* metrics has opted into measured fit, so a named
	// face we have no exact metrics for falls back to the conservative heuristic rather
	// than degrading to the bare flag — overflow still self-corrects, just less precisely.
	// An unnamed (theme-default) face stays unmeasurable: we cannot guess which face it is.
	const heuristicFaces = new Set<string>()
	const resolve = makeRegistryResolver(registry, face => heuristicFaces.add(face))
	const unmeasuredShrink = new Set<string>()
	const unmeasuredResize = new Set<string>()

	const collectUnmeasured = (paragraphs: FitParagraph[], into: Set<string>): void => {
		for (const para of paragraphs) for (const run of para.runs) if (!resolve(run)) into.add(run.fontFace ?? '(theme default)')
	}

	/**
	 * Bake measured shrink into a table's cells. Walks the cell grid (accounting for
	 * colspan/rowspan) to size each cell's box from its column widths and row heights,
	 * then lowers the run font sizes of any `fit:'shrink'` cell that overflows. Cells
	 * in auto-height rows (no fixed `rowH`/table `h`) are skipped — the row grows instead.
	 */
	const measureTableCells = (tableObj: ISlideObject, layout: PresSlideInternal['_presLayout']): void => {
		const rows = tableObj.arrTabRows
		if (!rows || rows.length === 0 || !rows[0]) return
		const tableOpts = (tableObj.options ?? {}) as RunOpts
		const numRows = rows.length
		const numCols = rows[0].reduce((n, c) => n + (Number(c?.options?.colspan) || 1), 0)
		if (!(numCols > 0)) return

		const cxEmu = tableOpts.w != null ? getSmartParseNumber(tableOpts.w, 'X', layout) : getSmartParseNumber('75%', 'X', layout)
		const colWidthsEmu = resolveTableColWidthsEmu(tableOpts.colW, cxEmu, numCols)
		const tableHeightEmu = tableOpts.h != null ? getSmartParseNumber(tableOpts.h, 'Y', layout) : (typeof tableOpts.cy === 'number' ? tableOpts.cy : 0)
		const rowHeightEmu = (rIdx: number): number => {
			if (Array.isArray(tableOpts.rowH) && tableOpts.rowH[rIdx]) return inch2Emu(Number(tableOpts.rowH[rIdx]))
			if (tableOpts.rowH != null && !Array.isArray(tableOpts.rowH) && !isNaN(Number(tableOpts.rowH))) return inch2Emu(Number(tableOpts.rowH))
			if (tableHeightEmu > 0) return Math.round(tableHeightEmu / numRows)
			return 0 // auto-height row → grows to fit, no shrink
		}

		// occupied[c] = rows still covered by a rowspan started above (incl. current row).
		const occupied = new Array<number>(numCols).fill(0)
		for (let r = 0; r < numRows; r++) {
			let col = 0
			for (const cell of rows[r]) {
				while (col < numCols && occupied[col] > 0) col++
				if (col >= numCols) break
				const colspan = Math.max(1, Number(cell?.options?.colspan) || 1)
				const rowspan = Math.max(1, Number(cell?.options?.rowspan) || 1)
				const colStart = col
				const colEnd = Math.min(colStart + colspan, numCols)
				for (let c = colStart; c < colEnd; c++) occupied[c] = rowspan
				col = colEnd

				const cellOpts = (cell?.options ?? {}) as TableCellProps
				const fit = cellOpts.fit ?? (tableOpts.fit === 'shrink' ? 'shrink' : undefined)
				if (fit !== 'shrink') continue

				let widthEmu = 0
				for (let c = colStart; c < colEnd; c++) widthEmu += colWidthsEmu[c] ?? 0
				let heightEmu = 0
				let autoHeight = false
				for (let rr = r; rr < Math.min(r + rowspan, numRows); rr++) {
					const h = rowHeightEmu(rr)
					if (h <= 0) { autoHeight = true; break }
					heightEmu += h
				}
				if (autoHeight) continue

				const eff = effectiveCellOpts(cellOpts, tableOpts)
				const ins = resolveCellInsetsEmu(eff.margin)
				const innerWidthPt = (widthEmu - ins.marL - ins.marR) / EMU_PER_POINT
				const innerHeightPt = (heightEmu - ins.marT - ins.marB) / EMU_PER_POINT
				if (!(innerWidthPt > 0) || !(innerHeightPt > 0)) continue

				const paragraphs = extractParagraphs({ text: cell.text, options: eff } as unknown as ISlideObject)
				if (!paragraphs) continue
				const box: FitBox = { innerWidthPt, innerHeightPt }
				const outcome = solveShrink(paragraphs, box, resolve)
				if (outcome.kind === 'shrink') {
					const f = outcome.result.fontScalePct / 100
					if (f < 1) scaleCellFontSizes(cell, eff, f)
				} else if (outcome.kind === 'unmeasurable') {
					collectUnmeasured(paragraphs, unmeasuredShrink)
				}
				// 'fits' → leave the authored size; the text already fits.
			}
			for (let c = 0; c < numCols; c++) if (occupied[c] > 0) occupied[c]--
		}
	}

	for (const slide of slides) {
		for (const obj of slide._slideObjects ?? []) {
			if (obj._type === SLIDE_OBJECT_TYPES.table) {
				measureTableCells(obj, slide._presLayout)
				continue
			}
			if (obj._type !== SLIDE_OBJECT_TYPES.text) continue
			// Only the bare string forms opt into measurement; an explicit object form is
			// already baked by the caller, and 'none' is a no-op.
			const fit = obj.options?.fit
			if (fit !== 'shrink' && fit !== 'resize') continue

			const paragraphs = extractParagraphs(obj)
			if (!paragraphs) continue
			const box = computeBox(obj, slide._presLayout)
			if (!box) continue

			if (fit === 'shrink') {
				const outcome = solveShrink(paragraphs, box, resolve)
				if (outcome.kind === 'shrink') {
					const { fontScalePct, lnSpcReductionPct } = outcome.result
					obj.options!.fit = {
						type: 'shrink',
						fontScale: fontScalePct,
						lnSpcReduction: lnSpcReductionPct || undefined,
					}
				} else if (outcome.kind === 'unmeasurable') {
					collectUnmeasured(paragraphs, unmeasuredShrink)
				}
				// 'fits' → leave the bare flag; the text already fits, so no scale is needed.
			} else {
				const outcome = solveResize(paragraphs, box, resolve)
				if (outcome.kind === 'resize') {
					const opts = obj.options as RunOpts
					const { tIns, bIns } = resolveInsetsEmu(opts)
					const oldHeightEmu = getSmartParseNumber(opts.h, 'Y', slide._presLayout)
					const newHeightEmu = Math.round(outcome.neededInnerHeightPt * EMU_PER_POINT) + tIns + bIns
					// Shift the box origin so growth/shrink honors the vertical anchor; `off.y`
					// moves up by the anchor's share of the height delta (0 / half / full for t / ctr / b).
					const oldYEmu = getSmartParseNumber(opts.y, 'Y', slide._presLayout)
					const shiftEmu = Math.round((newHeightEmu - oldHeightEmu) * anchorTopShareOfDelta(opts))
					opts.h = `${newHeightEmu}emu`
					if (shiftEmu !== 0) opts.y = `${oldYEmu - shiftEmu}emu`
				} else {
					collectUnmeasured(paragraphs, unmeasuredResize)
				}
			}
		}
	}

	if (unmeasuredShrink.size > 0) {
		console.warn(
			`Warning: fit:'shrink' could not be measured for font(s) [${[...unmeasuredShrink].join(', ')}] — ` +
				'no registered metrics. Emitting bare <a:normAutofit/> (text will not pre-shrink in headless renders). ' +
				'Call pptx.registerFontMetrics(face, fontFilePathOrBytes) to enable measured fit.'
		)
	}
	if (unmeasuredResize.size > 0) {
		console.warn(
			`Warning: fit:'resize' could not be measured for font(s) [${[...unmeasuredResize].join(', ')}] — ` +
				'no registered metrics. Emitting bare <a:spAutoFit/> with the authored height (box will not auto-grow in headless renders). ' +
				'Call pptx.registerFontMetrics(face, fontFilePathOrBytes) to enable measured fit.'
		)
	}
	if (heuristicFaces.size > 0) {
		console.warn(
			`Note: measured fit used a conservative average-advance estimate for font(s) [${[...heuristicFaces].join(', ')}] — ` +
				'no exact metrics registered. Fit is approximate (may shrink/grow more than necessary). ' +
				'Call pptx.registerFontMetrics(face, fontFilePathOrBytes) for an exact fit.'
		)
	}
}
