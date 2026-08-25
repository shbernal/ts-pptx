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
	getHeuristicFontMetrics,
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

describe('getHeuristicFontMetrics through dist: the per-character width buckets', () => {
	// The unregistered-font fallback is the width model behind every `fit:` on a face with
	// no registered metrics, so its buckets decide whether a headless render pre-shrinks
	// enough. `font-heuristic.test.js` proves them from `src/`; this drives the shipped
	// bundle, where an entry dropped by the emitter would otherwise go unnoticed.
	const h = getHeuristicFontMetrics()
	/** Advance of `text` in em, at the 1000pt size that makes an em 1000pt. */
	const em = (text) => h.advanceWidthPt(text, 1000) / 1000

	test('every bucket is a distinct width, widest to narrowest', () => {
		// One character per branch, in the order the model ranks them. Asserting the whole
		// ordering at once is what makes a bucket that silently merges into its neighbour
		// fail here rather than just shifting a measurement somewhere downstream.
		expect(em('日')).toBeCloseTo(1.0, 6) // CJK / full-width and beyond
		expect(em('@')).toBeCloseTo(1.0, 6) // and the two full-width ASCII symbols
		expect(em('%')).toBeCloseTo(1.0, 6)
		expect(em('W')).toBeCloseTo(0.98, 6)
		expect(em('m')).toBeCloseTo(0.9, 6)
		expect(em('A')).toBeCloseTo(0.72, 6)
		expect(em('ā')).toBeCloseTo(0.62, 6) // Latin-Extended / accented / misc symbols
		expect(em('a')).toBeCloseTo(0.58, 6) // lowercase, digits, default punctuation
		expect(em('é')).toBeCloseTo(0.58, 6) // Latin-1 is below the cut, so it is a default
		expect(em('f')).toBeCloseTo(0.4, 6)
		expect(em('i')).toBeCloseTo(0.3, 6)
		expect(em(' ')).toBeCloseTo(0.3, 6)
	})

	test("'@' and '%' are charged full width, unlike every other ASCII symbol", () => {
		// They sit above 'W' in a proportional face, and the bias here is deliberate:
		// over-estimating width is the direction that cannot overflow the box.
		expect(em('@')).toBeGreaterThan(em('W'))
		expect(em('%')).toBeGreaterThan(em('W'))
		expect(em('&')).toBeLessThan(em('@'))
	})

	test('the CJK bucket starts above the U+2E7F boundary, not at it', () => {
		// U+2E80 is the first code point charged a full em; the one below it is not, or
		// every Latin-Extended and symbol run would measure a third too wide.
		expect(em('⺀')).toBeCloseTo(1.0, 6)
		expect(em('⹿')).toBeCloseTo(0.62, 6)
	})
})
