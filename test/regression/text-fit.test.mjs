// Deterministic unit tests for the measured-fit line-break simulator + shrink
// solver (src/text-fit.ts). Uses SYNTHETIC font metrics so the suite is fully
// reproducible and needs no real font files — CI runs Node-only on Linux. The
// conservative-against-PowerPoint assertions against real fonts live in
// test/read/autofit-calibration-oracle.test.mjs (skipped when fonts are absent).
import { describe, test, expect } from 'vitest'
import {
	solveShrink,
	solveResize,
	measureHeightPt,
	SINGLE_LINE_PITCH,
	FONT_SCALE_STEP_PCT,
	MIN_FONT_SCALE_PCT,
	WIDTH_SAFETY_FACTOR,
	HEIGHT_SAFETY_FACTOR,
} from '../../src/text-fit.ts'

// Monospace-ish synthetic metrics: every code point advances `emPerChar` ems.
// advanceWidthPt(text, size) === count(text) * emPerChar * size (+ charSpacing).
const mono = (emPerChar = 0.5) => ({
	unitsPerEm: 1000,
	advanceWidthPt(text, sizePt, charSpacingPt = 0) {
		const n = [...text].length
		return n * emPerChar * sizePt + n * charSpacingPt
	},
})

const para = (text, extra = {}) => ({ runs: [{ text, sizePt: 18, fontFace: 'Mono' }], ...extra })
const resolveMono = () => mono()
const linesOf = (height, sizePt = 18) => Math.round(height / (SINGLE_LINE_PITCH * sizePt))

describe('text-fit: measureHeightPt', () => {
	test('single short line height = pitch * size', () => {
		const h = measureHeightPt([para('hi')], 1000, resolveMono, 100, 0)
		expect(h).toBeCloseTo(SINGLE_LINE_PITCH * 18, 6)
	})

	test('greedy wrap counts lines (0.5em mono, 18pt → 9pt/char)', () => {
		// 10 chars/word * 9pt = 90pt per word; innerWidth 100pt fits one word, two wrap.
		const h = measureHeightPt([para('aaaaaaaaaa bbbbbbbbbb cccccccccc')], 100, resolveMono, 100, 0)
		expect(linesOf(h)).toBe(3)
	})

	test('hard newline forces a line', () => {
		const h = measureHeightPt([para('a\nb\nc')], 10000, resolveMono, 100, 0)
		expect(linesOf(h)).toBe(3)
	})

	test('over-long unbreakable token character-wraps', () => {
		// 20 chars * 9pt = 180pt in a 50pt box → ceil(180/50) = 4 lines via char-wrap.
		const h = measureHeightPt([para('aaaaaaaaaaaaaaaaaaaa')], 50, resolveMono, 100, 0)
		expect(linesOf(h)).toBe(4)
	})

	test('fontScale narrows lines (re-wrap at smaller size)', () => {
		const big = measureHeightPt([para('aaaaaaaaaa bbbbbbbbbb')], 100, resolveMono, 100, 0)
		const small = measureHeightPt([para('aaaaaaaaaa bbbbbbbbbb')], 100, resolveMono, 50, 0)
		expect(linesOf(big)).toBe(2)
		expect(linesOf(small, 9)).toBe(1) // at 50% each word is 45pt → both fit one line
	})

	test('line height follows tallest run in a mixed-size paragraph', () => {
		const mixed = {
			runs: [
				{ text: 'a', sizePt: 12, fontFace: 'Mono' },
				{ text: 'b', sizePt: 40, fontFace: 'Mono' },
			],
		}
		const h = measureHeightPt([mixed], 10000, resolveMono, 100, 0)
		expect(h).toBeCloseTo(SINGLE_LINE_PITCH * 40, 6)
	})

	test('space-before/after add to height', () => {
		const h = measureHeightPt([para('hi', { spaceBeforePts: 5, spaceAfterPts: 7 })], 1000, resolveMono, 100, 0)
		expect(h).toBeCloseTo(SINGLE_LINE_PITCH * 18 + 12, 6)
	})

	test('exact line spacing overrides the calibrated pitch', () => {
		const h = measureHeightPt([para('a\nb', { lineSpacingPts: 30 })], 10000, resolveMono, 100, 0)
		expect(h).toBeCloseTo(60, 6)
	})

	test('unmeasurable when a run has no metrics', () => {
		expect(measureHeightPt([para('hi')], 1000, () => undefined, 100, 0)).toBeNull()
		expect(measureHeightPt([para('hi')], 0, resolveMono, 100, 0)).toBeNull()
	})
})

describe('text-fit: solveShrink', () => {
	const box = (innerWidthPt, innerHeightPt) => ({ innerWidthPt, innerHeightPt })

	test('fits at 100% → no shrink', () => {
		expect(solveShrink([para('hi')], box(1000, 100), resolveMono)).toEqual({ kind: 'fits' })
	})

	test('unmeasurable propagates', () => {
		expect(solveShrink([para('hi')], box(1000, 100), () => undefined)).toEqual({ kind: 'unmeasurable' })
	})

	test('overflow → largest grid scale that fits, and it actually fits', () => {
		// 5 words of 10 chars; tall enough to need shrinking in a short box.
		const paras = [para('aaaaaaaaaa bbbbbbbbbb cccccccccc dddddddddd eeeeeeeeee')]
		const b = box(120, 30)
		const out = solveShrink(paras, b, resolveMono)
		expect(out.kind).toBe('shrink')
		const { fontScalePct } = out.result
		// On the 2.5% grid, ≤ 100, ≥ floor.
		expect(Number.isInteger(fontScalePct / FONT_SCALE_STEP_PCT)).toBe(true)
		expect(fontScalePct).toBeLessThan(100)
		expect(fontScalePct).toBeGreaterThanOrEqual(MIN_FONT_SCALE_PCT)
		// Mirror the solver's safety-inflated fit criterion.
		const fitsInflated = (scale) =>
			measureHeightPt(paras, b.innerWidthPt, resolveMono, scale, 0, WIDTH_SAFETY_FACTOR) * HEIGHT_SAFETY_FACTOR <=
			b.innerHeightPt
		// The chosen scale fits (and a fortiori the pure height fits — conservative)…
		expect(fitsInflated(fontScalePct)).toBe(true)
		expect(measureHeightPt(paras, b.innerWidthPt, resolveMono, fontScalePct, 0)).toBeLessThanOrEqual(b.innerHeightPt)
		// …and the next step up does NOT (it is the largest that fits).
		if (fontScalePct + FONT_SCALE_STEP_PCT <= 100) {
			expect(fitsInflated(fontScalePct + FONT_SCALE_STEP_PCT)).toBe(false)
		}
	})

	test('extreme overflow floors at MIN_FONT_SCALE_PCT', () => {
		const long = 'x '.repeat(400).trim()
		const out = solveShrink([para(long)], box(50, 20), resolveMono)
		expect(out).toEqual({ kind: 'shrink', result: { fontScalePct: MIN_FONT_SCALE_PCT, lnSpcReductionPct: 0 } })
	})
})

describe('text-fit: solveResize', () => {
	const box = (innerWidthPt, innerHeightPt) => ({ innerWidthPt, innerHeightPt })

	test('unmeasurable propagates', () => {
		expect(solveResize([para('hi')], box(1000, 100), () => undefined)).toEqual({ kind: 'unmeasurable' })
	})

	test('needed height = pitch * size, inflated by the height safety factor', () => {
		const out = solveResize([para('hi')], box(1000, 5), resolveMono)
		expect(out.kind).toBe('resize')
		expect(out.neededInnerHeightPt).toBeCloseTo(SINGLE_LINE_PITCH * 18 * HEIGHT_SAFETY_FACTOR, 6)
	})

	test('never under-estimates: needed height ≥ the true laid-out height (no overflow)', () => {
		const paras = [para('aaaaaaaaaa bbbbbbbbbb cccccccccc dddddddddd')]
		const out = solveResize(paras, box(120, 30), resolveMono)
		expect(out.kind).toBe('resize')
		const trueHeight = measureHeightPt(paras, 120, resolveMono, 100, 0)
		expect(out.neededInnerHeightPt).toBeGreaterThanOrEqual(trueHeight)
	})

	test('wraps to more lines as width shrinks → taller needed height', () => {
		const paras = [para('aaaaaaaaaa bbbbbbbbbb cccccccccc')]
		const wide = solveResize(paras, box(10000, 1), resolveMono).neededInnerHeightPt
		const narrow = solveResize(paras, box(100, 1), resolveMono).neededInnerHeightPt
		expect(linesOf(wide / HEIGHT_SAFETY_FACTOR)).toBe(1)
		expect(narrow).toBeGreaterThan(wide)
	})
})
