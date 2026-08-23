/**
 * ts-pptx: Measured text fit (line-break simulator + shrink solver)
 *
 * Computes a baked `fontScale` for `fit: 'shrink'` so overflowing text renders
 * pre-shrunk in headless renderers (and on plain file-open) without a manual
 * edit/resize. See `docs/measured-text-fit.md`.
 *
 * Calibration: every constant here is pinned against PowerPoint-authored fixtures
 * (`test/read/fixtures/autofit-calibration.json`). The model errs **conservative**
 * — it over-estimates width (raw advances, no kerning) and uses the calibrated
 * line pitch — so the computed `fontScale` is ≤ PowerPoint's and the text never
 * overflows in either PowerPoint or LibreOffice.
 */

import type { FontMetrics } from './font-metrics.js'

/**
 * PowerPoint single-spacing line pitch as a multiple of font size.
 *
 * Calibrated from `autofit-calibration.json` (Deck 1, `autofit-line-metrics`):
 * the per-line `cy` delta is a **font-independent** 1.2117× at 12/18/32 pt across
 * {Aptos, Aptos SemiBold, Calibri, Tahoma, Arial}, and PowerPoint vs LibreOffice
 * agree to ≤ 0.041 pt. See `test/read/fixtures/README.md` → "Findings" #1–#2.
 */
export const SINGLE_LINE_PITCH = 1.2117

/** PowerPoint's `fontScale` search grid: discrete 2.5% steps (Findings #3). */
export const FONT_SCALE_STEP_PCT = 2.5

/** Floor for the shrink search (see `docs/measured-text-fit.md` → Solvers). */
export const MIN_FONT_SCALE_PCT = 25

/**
 * Conservative safety factors applied by the shrink solver (not by the pure
 * `measureHeightPt`). Raw `hmtx` advances summed without shaping still pack a line
 * a little tighter than PowerPoint, which lays text out at a device DPI and rounds
 * each glyph advance up — so PowerPoint wraps marginally earlier and its fitted
 * box is marginally taller. Inflating measured width and height by these factors
 * makes the computed `fontScale` ≤ PowerPoint's across the calibration oracle
 * (`test/read/autofit-calibration-oracle.test.js`: 0 violations over the Aptos
 * cases). They err on the side of shrinking slightly too much, never overflowing.
 */
export const WIDTH_SAFETY_FACTOR = 1.03
export const HEIGHT_SAFETY_FACTOR = 1.04

/** One text run with the inputs measurement needs. */
export interface FitRun {
	text: string
	sizePt: number
	bold?: boolean
	italic?: boolean
	fontFace?: string
	/** Character spacing in points (PowerPoint `spc`/1200). */
	charSpacingPt?: number
}

/** One paragraph: its runs plus vertical-rhythm options. */
export interface FitParagraph {
	runs: FitRun[]
	/** Line spacing as a percent of single (100 = single). Ignored if `lineSpacingPts` set. */
	lineSpacingPct?: number
	/** Exact line spacing in points (overrides `lineSpacingPct`). */
	lineSpacingPts?: number
	spaceBeforePts?: number
	spaceAfterPts?: number
}

/** Inner box (shape minus insets) the text must fit, in points. */
export interface FitBox {
	innerWidthPt: number
	innerHeightPt: number
	/**
	 * `false` when the frame does not wrap (`wrap:false` / `wrap="none"`): text lays
	 * out as one line per paragraph (no width-wrapping), and the shrink solver must
	 * additionally keep the widest line within `innerWidthPt`. A too-wide line
	 * overflows HORIZONTALLY — `normAutofit`'s vertical shrink never catches it, so
	 * without this check a non-wrapping line is never shrunk and spills out of the box
	 * in PowerPoint. Defaults to wrapping (`true`), so table cells and `measureText`
	 * are unaffected.
	 */
	wrap?: boolean
}

/** Resolve the metrics for a run, or `undefined` if the face is not registered. */
export type MetricsResolver = (run: FitRun) => FontMetrics | undefined

export interface ShrinkResult {
	fontScalePct: number
	lnSpcReductionPct: number
}

export type ShrinkOutcome =
	| { kind: 'fits' } // already fits at 100% — no shrink needed
	| { kind: 'unmeasurable' } // a run's face has no registered metrics — caller degrades
	| { kind: 'shrink'; result: ShrinkResult }

interface WordToken {
	kind: 'word'
	w: number
	charWidths: number[]
}
type Token = WordToken | { kind: 'space'; w: number } | { kind: 'newline'; w: 0 }

const isWhitespace = (ch: string): boolean => ch === ' ' || ch === '\t' || ch === ' ' || /\s/.test(ch)

/**
 * A code point after which a line break is allowed without a space — the CJK and
 * full-width ranges (Hangul Jamo through the CJK Compatibility Ideographs and the
 * CJK Extension B/C blocks). Latin words break only at whitespace, but Chinese,
 * Japanese and Korean text breaks between any two characters, so each such code
 * point is its own wrap opportunity: PowerPoint lays these scripts out per
 * character (UAX #14 class ID for Han/Hangul/Kana), and treating one long CJK run
 * as a single unbreakable "word" would over-estimate the laid-out line count and
 * shrink text that actually fits.
 */
export function isCjkBreakCharacter(ch: string): boolean {
	const cp = ch.codePointAt(0) ?? 0
	return (
		(cp >= 0x1100 && cp <= 0x11ff) || // Hangul Jamo
		(cp >= 0x2e80 && cp <= 0x2fff) || // CJK Radicals / Kangxi / CJK Symbols
		(cp >= 0x3000 && cp <= 0x303f) || // CJK Symbols and Punctuation
		(cp >= 0x3040 && cp <= 0x30ff) || // Hiragana / Katakana
		(cp >= 0x3130 && cp <= 0x318f) || // Hangul Compatibility Jamo
		(cp >= 0x31a0 && cp <= 0x31bf) || // Hangul Jamo Extended-A
		(cp >= 0x31c0 && cp <= 0x31ef) || // CJK Strokes
		(cp >= 0x3200 && cp <= 0x32ff) || // Enclosed CJK Letters and Months
		(cp >= 0x3400 && cp <= 0x4dbf) || // CJK Unified Ideographs Extension A
		(cp >= 0x4e00 && cp <= 0x9fff) || // CJK Unified Ideographs
		(cp >= 0xac00 && cp <= 0xd7af) || // Hangul Syllables
		(cp >= 0xf900 && cp <= 0xfaff) || // CJK Compatibility Ideographs
		(cp >= 0xff00 && cp <= 0xffef) || // Halfwidth and Fullwidth Forms
		(cp >= 0x20000 && cp <= 0x2ffff) || // CJK Extensions B, C… (Plane 2)
		(cp >= 0x30000 && cp <= 0x323af) // CJK Extension G (Plane 3)
	)
}

/**
 * Tokenize one paragraph at a given `fontScalePct`. Words can span run boundaries
 * (a break is only allowed at whitespace, or after a CJK character — see
 * `isCjkBreakCharacter`). Returns `null` if any run's face has no registered
 * metrics (the paragraph cannot be measured).
 */
function tokenizeParagraph(
	para: FitParagraph,
	resolve: MetricsResolver,
	fontScalePct: number,
	widthSafety: number
): Token[] | null {
	const tokens: Token[] = []
	let curWord: WordToken | null = null
	const flushWord = (): void => {
		if (curWord) {
			tokens.push(curWord)
			curWord = null
		}
	}
	for (const run of para.runs) {
		const metrics = resolve(run)
		if (!metrics) return null
		const scaledSize = run.sizePt * (fontScalePct / 100)
		const charSpacingPt = run.charSpacingPt ?? 0
		for (const ch of run.text ?? '') {
			if (ch === '\n') {
				flushWord()
				tokens.push({ kind: 'newline', w: 0 })
			} else if (isWhitespace(ch)) {
				flushWord()
				// Tabs are approximated as a single space-width gap (documented P1 gap).
				const w = metrics.advanceWidthPt(ch === '\t' ? ' ' : ch, scaledSize, charSpacingPt) * widthSafety
				tokens.push({ kind: 'space', w })
			} else {
				const w = metrics.advanceWidthPt(ch, scaledSize, charSpacingPt) * widthSafety
				if (isCjkBreakCharacter(ch)) {
					flushWord()
					tokens.push({ kind: 'word', w, charWidths: [w] })
					continue
				}
				if (!curWord) curWord = { kind: 'word', w: 0, charWidths: [] }
				curWord.w += w
				curWord.charWidths.push(w)
			}
		}
	}
	flushWord()
	return tokens
}

/** Greedy `wrap=square` line layout for a tokenized paragraph. */
interface LineLayout {
	lines: number
	/** Width (points) of the widest laid-out line — for natural/widest-line measurement. */
	maxLineWidthPt: number
}

/**
 * Greedy `wrap=square` line layout: the wrapped line count plus the widest
 * laid-out line width. `maxLineWidthPt` tracks the running line width and records
 * it whenever a line completes (a wrap or the final line), so callers can measure
 * the actual text extent, not just the box width. Trailing whitespace counts
 * toward width (Findings #7), which slightly over-reports — the conservative
 * direction, consistent with the wrap model erring wide.
 */
function countLines(tokens: Token[], innerWidthPt: number): LineLayout {
	let lines = 1
	let lineW = 0
	let maxLineWidthPt = 0
	const finishLine = (): void => {
		if (lineW > maxLineWidthPt) maxLineWidthPt = lineW
	}
	for (const tok of tokens) {
		if (tok.kind === 'newline') {
			finishLine()
			lines++
			lineW = 0
		} else if (tok.kind === 'space') {
			// Trailing whitespace counts toward line width (Findings #7); if it
			// overflows, the line wraps and the space is consumed at the break.
			if (lineW + tok.w > innerWidthPt && lineW > 0) {
				finishLine()
				lines++
				lineW = 0
			} else {
				lineW += tok.w
			}
		} else if (tok.w <= innerWidthPt) {
			if (lineW + tok.w > innerWidthPt && lineW > 0) {
				finishLine()
				lines++
				lineW = tok.w
			} else {
				lineW += tok.w
			}
		} else {
			// Over-long unbreakable word → character-wrap (Findings #7).
			if (lineW > 0) {
				finishLine()
				lines++
				lineW = 0
			}
			for (const cw of tok.charWidths) {
				if (lineW + cw > innerWidthPt && lineW > 0) {
					finishLine()
					lines++
					lineW = cw
				} else {
					lineW += cw
				}
			}
		}
	}
	finishLine()
	return { lines, maxLineWidthPt }
}

/** Max run size in a paragraph (line height follows the tallest run — Findings #7). */
function maxRunSizePt(para: FitParagraph): number {
	let max = 0
	for (const run of para.runs) if (run.sizePt > max) max = run.sizePt
	return max
}

/** Laid-out result of {@link measureLayout}: total height (points) + wrapped line count. */
export interface LayoutResult {
	heightPt: number
	lineCount: number
	/**
	 * Width (points) of the widest laid-out line across all paragraphs. Useful for
	 * tightening a box to the actual text extent or measuring a natural (single-line)
	 * width. Carries the same `widthSafety` inflation as the wrap decision, so a box
	 * set to this width will not re-wrap. Conservative (slightly wide).
	 */
	widestLineWidthPt: number
}

/**
 * Lay out `paragraphs` at a given scale/reduction and return the total height
 * (points) and wrapped line count, or `null` if any paragraph is unmeasurable.
 * Errs tall (conservative): includes space-before/after and the calibrated line
 * pitch. The single source of truth for both {@link measureHeightPt} (the solvers)
 * and the public layout-time measure API.
 */
export function measureLayout(
	paragraphs: FitParagraph[],
	innerWidthPt: number,
	resolve: MetricsResolver,
	fontScalePct: number,
	lnSpcReductionPct: number,
	widthSafety = 1
): LayoutResult | null {
	if (innerWidthPt <= 0) return null
	let total = 0
	let lineCount = 0
	let widestLineWidthPt = 0
	for (const para of paragraphs) {
		const tokens = tokenizeParagraph(para, resolve, fontScalePct, widthSafety)
		if (tokens === null) return null
		const { lines, maxLineWidthPt } = countLines(tokens, innerWidthPt)
		lineCount += lines
		if (maxLineWidthPt > widestLineWidthPt) widestLineWidthPt = maxLineWidthPt
		const scaledMax = maxRunSizePt(para) * (fontScalePct / 100)
		let lineHeight =
			para.lineSpacingPts != null && para.lineSpacingPts > 0
				? para.lineSpacingPts
				: SINGLE_LINE_PITCH * scaledMax * ((para.lineSpacingPct ?? 100) / 100)
		lineHeight *= 1 - lnSpcReductionPct / 100
		total += lines * lineHeight + (para.spaceBeforePts ?? 0) + (para.spaceAfterPts ?? 0)
	}
	return { heightPt: total, lineCount, widestLineWidthPt }
}

/**
 * Total laid-out height (points) at a given scale/reduction, or `null` if any
 * paragraph is unmeasurable. Thin wrapper over {@link measureLayout}.
 */
export function measureHeightPt(
	paragraphs: FitParagraph[],
	innerWidthPt: number,
	resolve: MetricsResolver,
	fontScalePct: number,
	lnSpcReductionPct: number,
	widthSafety = 1
): number | null {
	const result = measureLayout(paragraphs, innerWidthPt, resolve, fontScalePct, lnSpcReductionPct, widthSafety)
	return result === null ? null : result.heightPt
}

/**
 * Largest `fontScale` (on PowerPoint's 2.5% grid) at which the text fits the box.
 * Width and height both re-measured per scale. Conservative: because the model
 * over-estimates, the returned scale is ≤ PowerPoint's and the text never overflows.
 *
 * `lnSpcReduction` is left at 0 in P1 — dropping `fontScale` alone is provably
 * conservative (PowerPoint trades line-spacing reduction to keep the font *larger*,
 * so a reduction-free fit is always ≤ PowerPoint's scale). The calibration data for
 * the reduction ramp (Findings #4) is recorded for a future refinement.
 */
export type ResizeOutcome =
	| { kind: 'unmeasurable' } // a run's face has no registered metrics — caller degrades
	| { kind: 'resize'; neededInnerHeightPt: number }

/**
 * Inner height (points) the text needs at full size — the value to bake into the
 * shape's `a:ext/@cy` for `fit: 'resize'` (`spAutoFit`). Unlike shrink, resize has
 * **no safety net**: an under-estimate overflows (there is no text scaling
 * fallback), so this errs **tall** — width is inflated (earlier wrap ⇒ more lines)
 * and the laid-out height by the calibrated height factor, so the computed `cy` is
 * ≥ PowerPoint's and ≥ the LibreOffice-rendered height across the resize oracle.
 */
/**
 * Width used for wrap layout: the inner box width when wrapping, or unbounded for a
 * non-wrapping frame (`wrap:false`) so each paragraph lays out as a single line. A
 * non-wrapping frame's width constraint is enforced separately, against the widest
 * laid-out line — see {@link solveShrink}.
 */
function layoutWidthPt(box: FitBox): number {
	return box.wrap === false ? Infinity : box.innerWidthPt
}

export function solveResize(paragraphs: FitParagraph[], box: FitBox, resolve: MetricsResolver): ResizeOutcome {
	const h = measureHeightPt(paragraphs, layoutWidthPt(box), resolve, 100, 0, WIDTH_SAFETY_FACTOR)
	if (h === null) return { kind: 'unmeasurable' }
	return { kind: 'resize', neededInnerHeightPt: h * HEIGHT_SAFETY_FACTOR }
}

export function solveShrink(paragraphs: FitParagraph[], box: FitBox, resolve: MetricsResolver): ShrinkOutcome {
	const width = layoutWidthPt(box)
	const noWrap = box.wrap === false
	// Inflate measured width (earlier wrap) and height by the calibrated safety
	// factors so the fit threshold is conservative against PowerPoint. For a
	// non-wrapping frame the single line always fits the box HEIGHT, so height alone
	// never triggers a shrink: also require the widest (unwrapped) line to fit the box
	// WIDTH. `widestLineWidthPt` already carries the WIDTH_SAFETY_FACTOR inflation, so
	// comparing it straight to `innerWidthPt` stays on the conservative (shrink-more) side.
	const fits = (scale: number): boolean | null => {
		const layout = measureLayout(paragraphs, width, resolve, scale, 0, WIDTH_SAFETY_FACTOR)
		if (layout === null) return null
		if (layout.heightPt * HEIGHT_SAFETY_FACTOR > box.innerHeightPt) return false
		if (noWrap && layout.widestLineWidthPt > box.innerWidthPt) return false
		return true
	}

	const at100 = fits(100)
	if (at100 === null) return { kind: 'unmeasurable' }
	if (at100) return { kind: 'fits' }

	for (let scale = 100 - FONT_SCALE_STEP_PCT; scale >= MIN_FONT_SCALE_PCT; scale -= FONT_SCALE_STEP_PCT) {
		if (fits(scale)) return { kind: 'shrink', result: { fontScalePct: scale, lnSpcReductionPct: 0 } }
	}
	// Even at the floor it overflows our conservative model — bake the floor as best effort.
	return { kind: 'shrink', result: { fontScalePct: MIN_FONT_SCALE_PCT, lnSpcReductionPct: 0 } }
}
