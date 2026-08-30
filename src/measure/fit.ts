/**
 * ts-pptx: Measured-fit serialization pass
 *
 * Bridges slide text objects to the line-break simulator / shrink solver
 * (`text-fit.ts`). Runs during async export, BEFORE the `gen/` emitter builds the body:
 * for each text box with `fit: 'shrink'` and a registered font, it computes the
 * `fontScale` PowerPoint would have baked and rewrites `options.fit` to the
 * explicit object form (`{ type:'shrink', fontScale }`) so `genXmlNormAutofit`
 * emits `<a:normAutofit fontScale=…/>`. Without metrics it leaves the bare flag
 * untouched (current behavior) and warns once. See `docs/measured-text-fit.md`.
 */

import { SlideObjectType } from '../enums.js'
import { EMU_PER_POINT, POINTS_PER_INCH } from '../units.js'
import { getSmartParseNumber, resolveTableColWidthsEmu, resolveTableRowHeightEmu } from '../units-internal.js'
import { warn } from '../diagnostics.js'
import { collectUncoveredCodepoints, makeRegistryResolver, type FontMetricsRegistry } from './font-metrics.js'
import {
	solveShrink,
	solveResize,
	measureLayout,
	WIDTH_SAFETY_FACTOR,
	HEIGHT_SAFETY_FACTOR,
	type FitBox,
	type FitParagraph,
} from './text-fit.js'
import {
	anchorTopShareOfDelta,
	buildFitParagraphs,
	computeBox,
	extractParagraphs,
	resolveInsetsEmu,
	type RunOpts,
} from './paragraphs.js'
import {
	effectiveCellOpts,
	resolveCellInsetsEmu,
	scaleCellFontSizes,
	tableColCount,
	walkTableGrid,
} from './table-fit.js'
import type { MeasureTextOptions, TextMeasurement, TextProps } from '../types/index.js'
import type { SlideObject, PresSlideInternal } from '../types/internal.js'

/** Map the public {@link MeasureTextOptions} onto the internal run-option shape. */
function measureOptsToRunOpts(opts: MeasureTextOptions): RunOpts {
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
	}
}

/** A fresh result each call: the arrays are caller-owned, so they cannot be shared. */
const unmeasurable = (): TextMeasurement => ({
	heightIn: 0,
	lineCount: 0,
	widestLineIn: 0,
	measurable: false,
	approximatedFaces: [],
	uncoveredCodepoints: [],
	fitsBox: () => false,
	shrinkScaleFor: () => 100,
})

/** Flatten a face → uncovered-code-point map into one sorted, deduplicated list. */
function flattenCodepoints(byFace: Map<string, Set<number>>): number[] {
	const all = new Set<number>()
	for (const cps of byFace.values()) for (const cp of cps) all.add(cp)
	return [...all].sort((a, b) => a - b)
}

/**
 * Render uncovered code points for a warning as `U+XXXX`, capped so one bad run cannot
 * flood the console with a whole unsupported script.
 */
function formatCodepoints(cps: Set<number>): string {
	const all = [...cps].sort((a, b) => a - b)
	const shown = all.slice(0, 8).map((cp) => `U+${cp.toString(16).toUpperCase().padStart(4, '0')}`)
	return all.length > shown.length ? `${shown.join(' ')} +${all.length - shown.length} more` : shown.join(' ')
}

/**
 * Layout-time text measurement against registered metrics — the public engine
 * behind `pptx.measureText()`. Uses the **same** calibrated wrap model, resolver
 * semantics, and conservative safety factors as the export-time bake
 * ({@link applyMeasuredFit} / {@link solveResize} / {@link solveShrink}), so a
 * layout-time prediction matches the value the export would bake **for any deck
 * that opted into measured fit (i.e. registered at least one face)**.
 *
 * With an **empty** registry the two intentionally diverge: {@link applyMeasuredFit}
 * treats "no metrics" as "not opted in" and bakes nothing, while this returns
 * heuristic numbers so the API is useful with zero setup. Check
 * `approximatedFaces` if that distinction matters.
 *
 * Synchronous: assumes metrics are pre-registered (lookup is sync). A named face
 * with no exact metrics silently uses the conservative heuristic (same as export)
 * and is reported in `approximatedFaces`; an unnamed theme-default face returns
 * `measurable: false`. Code points a *registered* face has no glyph for are reported
 * in `uncoveredCodepoints`, and are the one case where the height may come back short
 * rather than tall — there is no font fallback here, only the font's `.notdef` advance.
 */
export function measureText(
	registry: FontMetricsRegistry,
	text: string | TextProps[],
	opts: MeasureTextOptions
): TextMeasurement {
	const runs: TextProps[] =
		typeof text === 'string' || typeof text === 'number' ? [{ text: String(text) }] : Array.isArray(text) ? text : []
	const paragraphs = buildFitParagraphs(runs, measureOptsToRunOpts(opts))
	if (!paragraphs) return unmeasurable()

	const inset = opts.insetIn ?? 0
	const innerWidthPt = (opts.wIn - 2 * inset) * POINTS_PER_INCH
	const heuristicFaces = new Set<string>()
	const resolve = makeRegistryResolver(registry, (face) => heuristicFaces.add(face))

	// Coverage is a property of the runs, not of the layout, so it is audited up front —
	// which also keeps it out of the resolver that `shrinkScaleFor()` re-enters.
	const uncoveredByFace = new Map<string, Set<number>>()
	collectUncoveredCodepoints(paragraphs, registry, uncoveredByFace)

	// Conservative (tall) layout at full size, mirroring solveResize: inflate width
	// (earlier wrap) by WIDTH_SAFETY and the height by HEIGHT_SAFETY. The resolver runs
	// during layout, so `heuristicFaces` is only populated once this returns.
	const layout = measureLayout(paragraphs, innerWidthPt, resolve, 100, 0, WIDTH_SAFETY_FACTOR)
	if (layout === null) return unmeasurable()
	const heightPt = layout.heightPt * HEIGHT_SAFETY_FACTOR
	const heightIn = heightPt / POINTS_PER_INCH

	return {
		heightIn,
		lineCount: layout.lineCount,
		widestLineIn: layout.widestLineWidthPt / POINTS_PER_INCH,
		measurable: true,
		// Snapshot: shrinkScaleFor() re-enters the same resolver, so the live set must not leak.
		approximatedFaces: [...heuristicFaces],
		uncoveredCodepoints: flattenCodepoints(uncoveredByFace),
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
 * not be measured (missing metrics), and once if one held code points its registered
 * face has no glyph for, so neither kind of overflow is silently ignored.
 */
export function applyMeasuredFit(slides: PresSlideInternal[], registry: FontMetricsRegistry): void {
	if (registry.size === 0) return

	// A deck that registered *some* metrics has opted into measured fit, so a named
	// face we have no exact metrics for falls back to the conservative heuristic rather
	// than degrading to the bare flag — overflow still self-corrects, just less precisely.
	// An unnamed (theme-default) face stays unmeasurable: we cannot guess which face it is.
	const heuristicFaces = new Set<string>()
	const resolve = makeRegistryResolver(registry, (face) => heuristicFaces.add(face))
	const unmeasuredShrink = new Set<string>()
	const unmeasuredResize = new Set<string>()
	const uncoveredByFace = new Map<string, Set<number>>()

	const collectUnmeasured = (paragraphs: FitParagraph[], into: Set<string>): void => {
		for (const para of paragraphs)
			for (const run of para.runs) if (!resolve(run)) into.add(run.fontFace ?? '(theme default)')
	}

	/** Audit every set of runs the pass is about to measure — see the warning below. */
	const auditCoverage = (paragraphs: FitParagraph[]): void =>
		collectUncoveredCodepoints(paragraphs, registry, uncoveredByFace)

	/**
	 * Bake measured shrink into a table's cells. Walks the cell grid (accounting for
	 * colspan/rowspan) to size each cell's box from its column widths and row heights,
	 * then lowers the run font sizes of any `fit:'shrink'` cell that overflows. Cells
	 * in auto-height rows (no fixed `rowH`/table `h`) are skipped — the row grows instead.
	 */
	const measureTableCells = (tableObj: SlideObject, layout: PresSlideInternal['_presLayout']): void => {
		const rows = tableObj.arrTabRows
		if (!rows || rows.length === 0 || !rows[0]) return
		const tableOpts = (tableObj.options ?? {}) as RunOpts
		const numRows = rows.length
		const numCols = tableColCount(rows)
		if (!(numCols > 0)) return

		const cxEmu =
			tableOpts.w != null ? getSmartParseNumber(tableOpts.w, 'X', layout) : getSmartParseNumber('75%', 'X', layout)
		const colWidthsEmu = resolveTableColWidthsEmu(tableOpts.colW, cxEmu, numCols)
		const tableHeightEmu =
			tableOpts.h != null
				? getSmartParseNumber(tableOpts.h, 'Y', layout)
				: typeof tableOpts.cy === 'number'
					? tableOpts.cy
					: 0
		// `null` is an auto-height row — it grows to fit, so there is nothing to shrink into.
		const rowHeightEmu = (rIdx: number): number =>
			resolveTableRowHeightEmu(tableOpts.rowH, rIdx, tableHeightEmu, numRows) ?? 0

		for (const { cell, row: r, col: colStart, colSpan, rowSpan } of walkTableGrid(rows, numCols)) {
			const colEnd = colStart + colSpan
			const cellOpts = cell?.options ?? {}
			const fit = cellOpts.fit ?? (tableOpts.fit === 'shrink' ? 'shrink' : undefined)
			if (fit !== 'shrink') continue

			let widthEmu = 0
			for (let c = colStart; c < colEnd; c++) widthEmu += colWidthsEmu[c] ?? 0
			let heightEmu = 0
			let autoHeight = false
			for (let rr = r; rr < r + rowSpan; rr++) {
				const h = rowHeightEmu(rr)
				if (h <= 0) {
					autoHeight = true
					break
				}
				heightEmu += h
			}
			if (autoHeight) continue

			const eff = effectiveCellOpts(cellOpts, tableOpts)
			const ins = resolveCellInsetsEmu(eff.margin)
			const innerWidthPt = (widthEmu - ins.marL - ins.marR) / EMU_PER_POINT
			const innerHeightPt = (heightEmu - ins.marT - ins.marB) / EMU_PER_POINT
			if (!(innerWidthPt > 0) || !(innerHeightPt > 0)) continue

			const paragraphs = extractParagraphs({ text: cell.text, options: eff } as unknown as SlideObject)
			if (!paragraphs) continue
			auditCoverage(paragraphs)
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
	}

	const measureObject = (obj: SlideObject, layout: PresSlideInternal['_presLayout']): void => {
		// Groups keep their children outside slide._slideObjects. Descend through every
		// nesting level so grouping remains a visual/editability operation and does not
		// disable the export-time fit pass for contained text.
		if (obj._type === SlideObjectType.group) {
			for (const child of obj._groupObjects ?? []) measureObject(child, layout)
			return
		}
		if (obj._type === SlideObjectType.table) {
			measureTableCells(obj, layout)
			return
		}
		if (obj._type !== SlideObjectType.text) return
		// Only the bare string forms opt into measurement; an explicit object form is
		// already baked by the caller, and 'none' is a no-op.
		const options = obj.options
		if (!options) return
		const fit = options.fit
		if (fit !== 'shrink' && fit !== 'resize') return

		const paragraphs = extractParagraphs(obj)
		if (!paragraphs) return
		auditCoverage(paragraphs)
		const box = computeBox(obj, layout)
		if (!box) return

		if (fit === 'shrink') {
			const outcome = solveShrink(paragraphs, box, resolve)
			if (outcome.kind === 'shrink') {
				const { fontScalePct, lnSpcReductionPct } = outcome.result
				options.fit = {
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
				const oldHeightEmu = getSmartParseNumber(opts.h, 'Y', layout)
				const newHeightEmu = Math.round(outcome.neededInnerHeightPt * EMU_PER_POINT) + tIns + bIns
				// Shift the box origin so growth/shrink honors the vertical anchor; `off.y`
				// moves up by the anchor's share of the height delta (0 / half / full for t / ctr / b).
				const oldYEmu = getSmartParseNumber(opts.y, 'Y', layout)
				const shiftEmu = Math.round((newHeightEmu - oldHeightEmu) * anchorTopShareOfDelta(opts))
				opts.h = `${newHeightEmu}emu`
				if (shiftEmu !== 0) opts.y = `${oldYEmu - shiftEmu}emu`
			} else {
				collectUnmeasured(paragraphs, unmeasuredResize)
			}
		}
	}

	for (const slide of slides) {
		for (const obj of slide._slideObjects ?? []) measureObject(obj, slide._presLayout)
	}

	if (unmeasuredShrink.size > 0) {
		warn(
			'measure/shrink-unmeasured',
			`fit:'shrink' could not be measured for font(s) [${[...unmeasuredShrink].join(', ')}] — ` +
				'no registered metrics. Emitting bare <a:normAutofit/> (text will not pre-shrink in headless renders). ' +
				'Call pptx.registerFontMetrics(face, fontFilePathOrBytes) to enable measured fit.'
		)
	}
	if (unmeasuredResize.size > 0) {
		warn(
			'measure/resize-unmeasured',
			`fit:'resize' could not be measured for font(s) [${[...unmeasuredResize].join(', ')}] — ` +
				'no registered metrics. Emitting bare <a:spAutoFit/> with the authored height (box will not auto-grow in headless renders). ' +
				'Call pptx.registerFontMetrics(face, fontFilePathOrBytes) to enable measured fit.'
		)
	}
	if (uncoveredByFace.size > 0) {
		const perFace = [...uncoveredByFace].map(([face, cps]) => `${face} → ${formatCodepoints(cps)}`).join('; ')
		warn(
			'measure/uncovered-codepoints',
			`Measured fit charged the .notdef advance for code point(s) the registered face does not cover (${perFace}) — ` +
				'PowerPoint renders those from a substituted font, so the measured width can be wrong in EITHER direction. ' +
				'Too narrow drops a line and the text overflows the baked size, the one case measured fit does not err safe. ' +
				'Register a face that covers them, or set an explicit fontFace on those runs.'
		)
	}
	if (heuristicFaces.size > 0) {
		warn(
			'measure/heuristic-metrics',
			`Note: measured fit used a conservative average-advance estimate for font(s) [${[...heuristicFaces].join(', ')}] — ` +
				'no exact metrics registered. Fit is approximate (may shrink/grow more than necessary). ' +
				'Call pptx.registerFontMetrics(face, fontFilePathOrBytes) for an exact fit.'
		)
	}
}
