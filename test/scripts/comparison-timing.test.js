/**
 * The generation-time measurement, in the parts of it that can be checked without a clock.
 *
 * A timing harness cannot be regression-tested on its numbers — that is the whole problem
 * with numbers a clock produced. What it can be tested on is everything around them: that
 * the statistics do what their names say, that the two matched modes really are matched,
 * and that the two arms of the scale deck have not drifted apart into two different decks
 * whose difference would be reported as a speed.
 *
 * That last one is the reason this file exists. `scripts/comparison/workloads.mjs` writes
 * the deck out twice, once per library, because the syntax page prints those bodies and a
 * body that branched on which library was running would show a reader a line no consumer
 * writes. The cost of writing it twice is that the two copies can drift, and a drift here
 * is invisible: both arms still run, both still produce a deck, and the timing table
 * quietly starts comparing forty slides against forty-one.
 */

import { describe, expect, test } from 'vitest'
import { SUBJECTS } from '../../scripts/comparison/probes.mjs'
import { bytesAgreement, MODES, quantile, summarise } from '../../scripts/comparison/timing.mjs'
import { scaleSource, scaleWorkload, SCALE_ARMS, SCALE_SIZES, WORKLOADS } from '../../scripts/comparison/workloads.mjs'

describe('the timing statistics', () => {
	test('quantile interpolates between neighbours and hits the ends exactly', () => {
		const values = [1, 2, 3, 4]
		expect(quantile(values, 0)).toBe(1)
		expect(quantile(values, 1)).toBe(4)
		expect(quantile(values, 0.5)).toBe(2.5)
	})

	test('quantile does not care what order it is handed', () => {
		expect(quantile([9, 1, 5], 0.5)).toBe(5)
	})

	// An empty sample has no median, and a harness that returned 0 for one would publish a
	// library that generates decks instantaneously.
	test('quantile refuses an empty sample', () => {
		expect(() => quantile([], 0.5)).toThrow(/empty sample/)
	})

	test('summarise reports the median, the floor and the spread around it', () => {
		const summary = summarise([10, 10, 10, 10])
		expect(summary.median).toBe(10)
		expect(summary.min).toBe(10)
		expect(summary.spread).toBe(0)
		expect(summary.rounds).toBe(4)
	})

	test('spread grows with the scatter, as a percentage of the median', () => {
		expect(summarise([8, 9, 11, 12]).spread).toBeGreaterThan(summarise([9, 10, 10, 11]).spread)
	})
})

describe('output sizes beside the times', () => {
	test('agreement is a percentage of the larger of the two', () => {
		expect(bytesAgreement([100, 100])).toBe(0)
		expect(bytesAgreement([50, 100])).toBe(50)
	})

	// `null` rather than a number wherever there is nothing to divide, so a missing arm
	// cannot come out of the arithmetic as two libraries agreeing perfectly.
	test('nothing to compare reads as nothing, not as zero', () => {
		expect(bytesAgreement([100])).toBeNull()
		expect(bytesAgreement([0, 100])).toBeNull()
		expect(bytesAgreement([/** @type {any} */ (undefined), 100])).toBeNull()
	})
})

describe('the matched modes', () => {
	test('every mode carries write options for every subject', () => {
		for (const mode of MODES)
			for (const subject of SUBJECTS) expect(mode.props[subject], `${mode.id}/${subject}`).toBeTypeOf('object')
	})

	// The point of the pair. If both modes asked for the same thing there would be no
	// control, and the page's claim that the two tables isolate the compressor would be
	// measuring one thing twice.
	test('the two modes differ, and upstream stays on one code path across them', () => {
		const [deflate, store] = MODES
		expect(deflate.id).toBe('deflate')
		expect(store.id).toBe('store')
		expect(deflate.props.pptxgenjs).not.toEqual(store.props.pptxgenjs)
		expect(deflate.props['ts-pptx']).not.toEqual(store.props['ts-pptx'])
		expect(/** @type {any} */ (deflate.props.pptxgenjs).outputType).toBe(
			/** @type {any} */ (store.props.pptxgenjs).outputType
		)
	})
})

describe('the scale corpus', () => {
	test('every size becomes a workload with both arms', () => {
		expect(WORKLOADS.map((workload) => workload.slides)).toEqual(SCALE_SIZES)
		for (const workload of WORKLOADS)
			for (const subject of SUBJECTS) expect(workload.build[subject], workload.id).toBeTypeOf('function')
	})

	test('workload ids are unique', () => {
		const ids = WORKLOADS.map((workload) => workload.id)
		expect(new Set(ids).size).toBe(ids.length)
	})

	// The two arms are written out separately so the page can print consumer-shaped code.
	// This is what keeps them the same deck: identical line for line, except for the one
	// call whose signature the two libraries genuinely spell differently.
	test('the two arms differ only in the addChart call', () => {
		const ours = scaleSource('ts-pptx').split('\n')
		const upstream = scaleSource('pptxgenjs').split('\n')
		expect(ours.length).toBe(upstream.length)
		const differing = ours.map((line, index) => [line, upstream[index]]).filter(([a, b]) => a !== b)
		expect(differing.length).toBe(1)
		for (const line of differing[0]) expect(line).toContain('addChart')
	})

	// The harness's own vocabulary must not reach the page. A body that mentions the subject
	// is a body that was written for the measurement rather than for a reader.
	test('neither printed arm mentions the harness', () => {
		for (const subject of SUBJECTS) {
			const source = scaleSource(subject)
			expect(source, subject).not.toContain('subject')
			expect(source, subject).toContain('pres.addSlide()')
		}
	})

	test('an unregistered subject is a throw, not an empty snippet', () => {
		expect(() => scaleSource('nope')).toThrow(/no timing workload arm/)
	})

	// The arms close over the slide count, so a workload built at one size must not be able
	// to hand back the deck for another.
	test('a workload builds exactly the slides it is named for', () => {
		const built = []
		const pres = {
			addSlide() {
				const slide = slideStub()
				built.push(slide)
				return slide
			},
		}
		scaleWorkload(3).build['ts-pptx'](pres)
		expect(built.length).toBe(3)
	})
})

/** The surface one slide of the scale deck touches, and nothing else. */
function slideStub() {
	return { addText() {}, addTable() {}, addChart() {}, addNotes() {} }
}

describe('the arms agree with what the timing loop runs', () => {
	// `scaleSource` prints the arm with the slide count bound away; the loop runs the arm
	// bound to a real size. Both come from `SCALE_ARMS`, and this is what says so — a page
	// printing one function while the clock timed another is the failure this whole file is
	// arranged to prevent.
	test('the printed body is the body a workload was built from', () => {
		for (const subject of SUBJECTS) {
			const printed = scaleSource(subject)
			const armBody = SCALE_ARMS[subject](7).toString()
			for (const line of printed.split('\n').filter((line) => line.trim() !== ''))
				expect(armBody, subject).toContain(line.trim())
		}
	})
})
