// The wrap model + solvers through the public `ts-pptx/measure` subpath (dist/measure.js).
//
// `measureLayout` / `measureHeightPt` / `solveShrink` / `solveResize` are published so a
// consumer can lay out its own geometry from `FitParagraph[]` it builds itself
// (docs/measured-text-fit.md → "reach for these primitives"). That entry reaches parts of
// the model the deck path provably cannot:
//
//   - `buildFitParagraphs` splits every "\n" into its own paragraph before the tokenizer
//     runs, so a newline never survives into a run — yet the tokenizer and the line counter
//     both handle one, because a hand-built run is free to contain it.
//   - `buildFitParagraphs` always fills `lineSpacingPct` / `spaceBeforePts` / `spaceAfterPts`,
//     so the defaults for an omitted field are only exercised by a hand-built paragraph.
//
// The sibling `text-fit.test.mjs` covers the same functions from `src/`, which proves the
// behavior but leaves the shipped bundle unmeasured; this drives the built artifact.
// `measured-fit-dist.test.mjs` covers everything reachable from a deck.
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, test, expect, beforeAll } from 'vitest'
import {
	measureLayout,
	measureHeightPt,
	solveShrink,
	solveResize,
	parseFontMetrics,
	SINGLE_LINE_PITCH,
	MIN_FONT_SCALE_PCT,
} from '../../../dist/measure.js'

const REG_PATH = fileURLToPath(new URL('../../read/fixtures/fonts/Silkscreen-Regular.ttf', import.meta.url))

/** Resolver backed by real Silkscreen metrics for every run. */
let resolve
/** Resolver that knows no face — the "unmeasurable" input every solver must detect. */
const resolveNone = () => undefined

beforeAll(async () => {
	const metrics = await parseFontMetrics(new Uint8Array(readFileSync(REG_PATH)))
	resolve = () => metrics
})

const para = (text, extra = {}) => ({ runs: [{ text, sizePt: 12 }], ...extra })

describe('measureLayout: a newline inside a hand-built run', () => {
	test('starts a new line without splitting the paragraph', () => {
		const single = measureLayout([para('alpha beta')], 400, resolve, 100, 0)
		const broken = measureLayout([para('alpha\nbeta')], 400, resolve, 100, 0)
		// 400pt is wide enough that neither wraps, so the extra line is the newline's doing.
		expect(single.lineCount).toBe(1)
		expect(broken.lineCount).toBe(2)
		expect(broken.heightPt).toBeCloseTo(single.heightPt * 2, 6)
	})

	test('flushes the word in progress, so the text either side is not merged', () => {
		// If the tokenizer did not flush at the newline, "alpha" and "beta" would fuse into
		// one unbreakable token and the widest line would be the width of both together.
		const broken = measureLayout([para('alpha\nbeta')], 400, resolve, 100, 0)
		const fused = measureLayout([para('alphabeta')], 400, resolve, 100, 0)
		expect(broken.widestLineWidthPt).toBeLessThan(fused.widestLineWidthPt)
	})

	test('a newline resets the line width, so it does not carry into the next line', () => {
		const trailing = measureLayout([para('alphabeta\nx')], 400, resolve, 100, 0)
		const alone = measureLayout([para('x')], 400, resolve, 100, 0)
		expect(trailing.widestLineWidthPt).toBeGreaterThan(alone.widestLineWidthPt)
		// The second line is just "x": its width is the short one, so the widest line is line 1.
		const firstOnly = measureLayout([para('alphabeta')], 400, resolve, 100, 0)
		expect(trailing.widestLineWidthPt).toBeCloseTo(firstOnly.widestLineWidthPt, 6)
	})

	test('counts consecutive newlines as blank lines', () => {
		expect(measureLayout([para('a\n\n\nb')], 400, resolve, 100, 0).lineCount).toBe(4)
	})

	test('a run with no text at all measures as an empty line rather than throwing', () => {
		// @ts-expect-error `FitRun.text` is required, but this entry is reachable from plain JS
		const layout = measureLayout([{ runs: [{ sizePt: 12 }] }], 400, resolve, 100, 0)
		expect(layout.lineCount).toBe(1)
		expect(layout.widestLineWidthPt).toBe(0)
	})
})

describe('measureLayout: paragraph fields the deck path always fills', () => {
	test('an omitted `lineSpacingPct` defaults to single spacing', () => {
		const implicit = measureLayout([{ runs: [{ text: 'x', sizePt: 12 }] }], 400, resolve, 100, 0)
		const explicit = measureLayout([para('x', { lineSpacingPct: 100 })], 400, resolve, 100, 0)
		expect(implicit.heightPt).toBeCloseTo(explicit.heightPt, 6)
		expect(implicit.heightPt).toBeCloseTo(SINGLE_LINE_PITCH * 12, 6)
	})

	test('omitted `spaceBeforePts`/`spaceAfterPts` contribute nothing', () => {
		const bare = measureLayout([{ runs: [{ text: 'x', sizePt: 12 }] }], 400, resolve, 100, 0)
		const zeroed = measureLayout([para('x', { spaceBeforePts: 0, spaceAfterPts: 0 })], 400, resolve, 100, 0)
		expect(bare.heightPt).toBe(zeroed.heightPt)
	})

	test('an exact `lineSpacingPts` overrides the calibrated pitch outright', () => {
		const layout = measureLayout([para('x', { lineSpacingPts: 30 })], 400, resolve, 100, 0)
		expect(layout.heightPt).toBeCloseTo(30, 6)
	})

	test('a non-positive `lineSpacingPts` falls back to the pitch instead of collapsing', () => {
		const layout = measureLayout([para('x', { lineSpacingPts: 0 })], 400, resolve, 100, 0)
		expect(layout.heightPt).toBeCloseTo(SINGLE_LINE_PITCH * 12, 6)
	})
})

describe('measureLayout / measureHeightPt: unmeasurable and degenerate inputs', () => {
	test('an unresolvable face returns null rather than a zero height', () => {
		expect(measureLayout([para('x')], 400, resolveNone, 100, 0)).toBeNull()
		expect(measureHeightPt([para('x')], 400, resolveNone, 100, 0)).toBeNull()
	})

	test('a non-positive width returns null', () => {
		expect(measureLayout([para('x')], 0, resolve, 100, 0)).toBeNull()
		expect(measureHeightPt([para('x')], -5, resolve, 100, 0)).toBeNull()
	})

	test('measureHeightPt is exactly measureLayout().heightPt', () => {
		const paras = [para('alpha beta gamma')]
		expect(measureHeightPt(paras, 60, resolve, 100, 0)).toBe(measureLayout(paras, 60, resolve, 100, 0).heightPt)
	})
})

describe('solvers through dist', () => {
	const box = (w, h, wrap) => ({ innerWidthPt: w, innerHeightPt: h, ...(wrap === undefined ? {} : { wrap }) })

	/**
	 * Assert a discriminated outcome's `kind` and narrow to it, so the assertions that
	 * follow can read the payload without a cast.
	 * @template {{ kind: string }} T
	 * @template {T['kind']} K
	 * @param {T} outcome
	 * @param {K} kind
	 * @returns {Extract<T, { kind: K }>}
	 */
	function expectKind(outcome, kind) {
		expect(outcome.kind).toBe(kind)
		return /** @type {Extract<T, { kind: K }>} */ (outcome)
	}

	test('solveShrink reports `fits` when the text already fits', () => {
		expect(solveShrink([para('x')], box(400, 400), resolve)).toEqual({ kind: 'fits' })
	})

	test('solveShrink reports `unmeasurable` for an unresolvable face', () => {
		expect(solveShrink([para('x')], box(400, 400), resolveNone)).toEqual({ kind: 'unmeasurable' })
	})

	test('solveShrink lands on the 2.5% grid below 100', () => {
		const outcome = expectKind(solveShrink([para('alpha beta gamma delta epsilon')], box(80, 20), resolve), 'shrink')
		expect(outcome.result.fontScalePct).toBeLessThan(100)
		expect(outcome.result.fontScalePct % 2.5).toBeCloseTo(0, 10)
	})

	test('solveShrink bakes the floor when even the floor overflows', () => {
		const paras = [para('alpha beta gamma delta epsilon zeta eta theta')]
		const outcome = expectKind(solveShrink(paras, box(20, 6), resolve), 'shrink')
		expect(outcome.result.fontScalePct).toBe(MIN_FONT_SCALE_PCT)
	})

	test('a `wrap:false` box lays out one line but is still held to the box width', () => {
		// Tall enough that the height test can never fire: any shrink is the width check.
		expectKind(solveShrink([para('alpha beta gamma delta')], box(40, 5000, false), resolve), 'shrink')
	})

	test('the same box wrapping normally needs no shrink', () => {
		expect(solveShrink([para('alpha beta gamma delta')], box(40, 5000), resolve).kind).toBe('fits')
	})

	test('solveResize returns the needed inner height, and `unmeasurable` without metrics', () => {
		const outcome = expectKind(solveResize([para('alpha beta gamma delta')], box(40, 10), resolve), 'resize')
		expect(outcome.neededInnerHeightPt).toBeGreaterThan(10)
		expect(solveResize([para('x')], box(40, 10), resolveNone)).toEqual({ kind: 'unmeasurable' })
	})
})
