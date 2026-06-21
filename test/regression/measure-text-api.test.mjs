// Layout-time public measurement API (docs/measured-text-fit.md). Three layers:
//  A) the core measureText(registry, …) + shared buildFitParagraphs against src,
//     with SYNTHETIC metrics so the suite is reproducible and needs no font files;
//  B) the same through the built `pptxgenjs/measure` subpath (P1 re-exports);
//  C) the pptx.measureText()/overflowsBox() instance methods through dist (the
//     heuristic path, so no real font is required).
// The KEY correctness assertion is the no-drift test: measureText's height equals
// the height the export-time resize bake (solveResize) uses for the same input.
import { describe, test, expect } from 'vitest'
import { measureText, buildFitParagraphs, makeRegistryResolver } from '../../src/measure-fit.ts'
import { solveResize, solveShrink, HEIGHT_SAFETY_FACTOR } from '../../src/text-fit.ts'
import { FontMetricsRegistry } from '../../src/font-metrics.ts'
import PptxGenJS from '../../dist/node.js'

// Monospace synthetic metrics: every code point advances `emPerChar` ems.
const mono = (emPerChar = 0.5) => ({
	unitsPerEm: 1000,
	advanceWidthPt(text, sizePt, charSpacingPt = 0) {
		const n = [...text].length
		return n * emPerChar * sizePt + n * charSpacingPt
	},
})

const regWith = (face = 'Mono') => {
	const r = new FontMetricsRegistry()
	r.set(face, mono())
	return r
}

const SENTENCE = 'aaaaaaaaaa bbbbbbbbbb cccccccccc dddddddddd eeeeeeeeee'

describe('measureText core (synthetic metrics)', () => {
	test('no drift: measureText height == solveResize bake height for the same input', () => {
		const reg = regWith()
		const opts = { wIn: 2, fontSize: 18, fontFace: 'Mono' }
		const m = measureText(reg, SENTENCE, opts)

		// Reproduce what the export pass would do, independently.
		const paras = buildFitParagraphs([{ text: SENTENCE }], { fontSize: 18, fontFace: 'Mono' })
		const innerWidthPt = opts.wIn * 72
		const outcome = solveResize(paras, { innerWidthPt, innerHeightPt: 9999 }, makeRegistryResolver(reg))
		expect(outcome.kind).toBe('resize')
		expect(m.heightIn * 72).toBeCloseTo(outcome.neededInnerHeightPt, 6)
	})

	test('height errs tall by the HEIGHT_SAFETY_FACTOR vs the raw laid-out height', () => {
		const reg = regWith()
		const m = measureText(reg, 'one short line', { wIn: 10, fontSize: 12, fontFace: 'Mono' })
		// Single line at 12pt: pitch 1.2117 * 12, inflated by the height safety factor.
		expect(m.lineCount).toBe(1)
		expect(m.heightIn * 72).toBeCloseTo(1.2117 * 12 * HEIGHT_SAFETY_FACTOR, 4)
	})

	test('narrower width wraps to more lines', () => {
		const reg = regWith()
		const wide = measureText(reg, SENTENCE, { wIn: 10, fontSize: 18, fontFace: 'Mono' })
		const narrow = measureText(reg, SENTENCE, { wIn: 2, fontSize: 18, fontFace: 'Mono' })
		expect(narrow.lineCount).toBeGreaterThan(wide.lineCount)
		expect(narrow.heightIn).toBeGreaterThan(wide.heightIn)
	})

	test('insetIn shrinks the usable width on both sides', () => {
		const reg = regWith()
		const noInset = measureText(reg, SENTENCE, { wIn: 4, fontSize: 18, fontFace: 'Mono' })
		const inset = measureText(reg, SENTENCE, { wIn: 4, fontSize: 18, fontFace: 'Mono', insetIn: 0.5 })
		expect(inset.lineCount).toBeGreaterThanOrEqual(noInset.lineCount)
	})

	test('fitsBox mirrors the conservative height', () => {
		const reg = regWith()
		const m = measureText(reg, SENTENCE, { wIn: 2, fontSize: 18, fontFace: 'Mono' })
		expect(m.fitsBox(m.heightIn)).toBe(true)
		expect(m.fitsBox(m.heightIn - 0.01)).toBe(false)
		expect(m.fitsBox(m.heightIn + 1)).toBe(true)
	})

	test('shrinkScaleFor: 100 when it fits, and matches solveShrink when it overflows', () => {
		const reg = regWith()
		const m = measureText(reg, SENTENCE, { wIn: 2, fontSize: 18, fontFace: 'Mono' })
		expect(m.shrinkScaleFor(m.heightIn + 1)).toBe(100)

		const tightIn = m.heightIn / 3
		const paras = buildFitParagraphs([{ text: SENTENCE }], { fontSize: 18, fontFace: 'Mono' })
		const outcome = solveShrink(paras, { innerWidthPt: 2 * 72, innerHeightPt: tightIn * 72 }, makeRegistryResolver(reg))
		const expected = outcome.kind === 'shrink' ? outcome.result.fontScalePct : 100
		expect(m.shrinkScaleFor(tightIn)).toBe(expected)
		expect(m.shrinkScaleFor(tightIn)).toBeLessThan(100)
	})

	test('named face without exact metrics is still measurable (heuristic)', () => {
		const reg = regWith() // only 'Mono' registered
		const m = measureText(reg, 'hello world', { wIn: 5, fontSize: 12, fontFace: 'SomeOtherFace' })
		expect(m.measurable).toBe(true)
		expect(m.heightIn).toBeGreaterThan(0)
	})

	test('unnamed (theme-default) face is unmeasurable', () => {
		const reg = regWith()
		const m = measureText(reg, 'hello world', { wIn: 5, fontSize: 12 })
		expect(m.measurable).toBe(false)
		expect(m.heightIn).toBe(0)
		expect(m.fitsBox(100)).toBe(false)
		expect(m.shrinkScaleFor(100)).toBe(100)
	})

	test('TextProps[] runs honor per-run overrides', () => {
		const reg = regWith()
		const m = measureText(
			reg,
			[
				{ text: 'small ', options: { fontSize: 10 } },
				{ text: 'BIG', options: { fontSize: 40 } },
			],
			{ wIn: 20, fontSize: 12, fontFace: 'Mono' }
		)
		// One line; line height follows the tallest run (40pt).
		expect(m.lineCount).toBe(1)
		expect(m.heightIn * 72).toBeCloseTo(1.2117 * 40 * HEIGHT_SAFETY_FACTOR, 4)
	})
})

describe('pptxgenjs/measure subpath (P1 re-exports, built)', () => {
	test('re-exports resolve and measureText works through dist with a synthetic registry', async () => {
		const mod = await import('../../dist/measure.js')
		for (const name of [
			'measureLayout',
			'measureHeightPt',
			'solveShrink',
			'solveResize',
			'SINGLE_LINE_PITCH',
			'FONT_SCALE_STEP_PCT',
			'MIN_FONT_SCALE_PCT',
			'WIDTH_SAFETY_FACTOR',
			'HEIGHT_SAFETY_FACTOR',
			'parseFontMetrics',
			'getHeuristicFontMetrics',
			'FontMetricsRegistry',
			'buildFitParagraphs',
			'makeRegistryResolver',
			'measureText',
		]) {
			expect(mod[name], `missing export: ${name}`).toBeDefined()
		}
		const reg = new mod.FontMetricsRegistry()
		reg.set('Mono', mono())
		const m = mod.measureText(reg, SENTENCE, { wIn: 2, fontSize: 18, fontFace: 'Mono' })
		expect(m.measurable).toBe(true)
		expect(m.lineCount).toBeGreaterThan(1)
	})
})

describe('pptx.measureText() / overflowsBox() instance methods (heuristic path)', () => {
	test('named face is measurable via the heuristic even with no metrics registered', () => {
		const pptx = new PptxGenJS()
		const m = pptx.measureText('A heading that is reasonably long', { wIn: 2, fontSize: 18, fontFace: 'Arial' })
		expect(m.measurable).toBe(true)
		expect(m.heightIn).toBeGreaterThan(0)
		expect(m.lineCount).toBeGreaterThanOrEqual(1)
	})

	test('unnamed face returns measurable:false', () => {
		const pptx = new PptxGenJS()
		expect(pptx.measureText('text', { wIn: 5, fontSize: 12 }).measurable).toBe(false)
	})

	test('overflowsBox: tall text in a tiny box overflows; short text does not', () => {
		const pptx = new PptxGenJS()
		const long = 'The quick brown fox jumps over the lazy dog. '.repeat(6)
		expect(pptx.overflowsBox(long, { wIn: 2, hIn: 0.5, fontSize: 18, fontFace: 'Arial' })).toBe(true)
		expect(pptx.overflowsBox('hi', { wIn: 5, hIn: 3, fontSize: 12, fontFace: 'Arial' })).toBe(false)
	})

	test('overflowsBox reports false for an unmeasurable (unnamed) face', () => {
		const pptx = new PptxGenJS()
		expect(pptx.overflowsBox('x'.repeat(500), { wIn: 1, hIn: 0.2, fontSize: 40 })).toBe(false)
	})
})
