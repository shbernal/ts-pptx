/**
 * The timing corpus: one deck shape, built at three sizes.
 *
 * `./programs.mjs` supplies the small end of the generation-time measurement for free — the
 * five bundle programs are already whole decks both libraries build, and timing the decks
 * whose bundles the page above already weighs is worth more than timing five new ones. What
 * they cannot supply is scale. The largest of them is three slides, which on a warm process
 * is a few milliseconds: small enough that the clock is measuring the harness as much as
 * the library, and far too small to answer the question a reader actually has, which is
 * what happens to a deck with a hundred slides in it.
 *
 * So this corpus is the other end. One shape — a title, a bulleted body, a table or a chart
 * on alternating slides, and speaker notes — repeated {@link SCALE_SIZES} times. The shape
 * is deliberately dull. A timing corpus is not trying to find the construct that is slowest
 * to emit; it is trying to make the per-slide cost visible, and a mixed but repetitive deck
 * is what lets three sizes of the same thing be compared with each other as well as across
 * the two libraries.
 *
 * ## Why the data is built inside the build function
 *
 * The other two corpora declare their shared values once at module scope, because the page
 * prints those declarations beside the code that used them, and `./corpus-data.mjs` guards
 * them from the libraries that edit what they are handed. Neither reason applies here, and
 * the second one would bite hard if it were ignored: a timing run calls one build function
 * dozens of times in a row, so a series pptxgenjs wrapped in place on round one would be
 * wrapped again on round two and every round after it. That is not a corpus the page has to
 * be protected from printing — it is a workload that grows while it is being timed.
 *
 * Building the rows and the series inside the function makes every round start from the
 * same place by construction, with nothing to reset and nothing to get wrong. It also costs
 * both arms the same handful of microseconds per round, which is the other half of why it
 * is safe: the measurement is a comparison, and an overhead both columns pay cancels.
 */
import { functionBody } from './source.mjs'

/**
 * The three deck sizes, in slides.
 *
 * Three rather than one because a single size cannot show a slope. Two libraries that are
 * both linear in slide count and two that diverge look identical at one point and obviously
 * different across three, and "does the gap widen with the deck" is the question a reader
 * with a big deck is really asking. The top size is the largest that keeps a full timing
 * pass inside its budget; the bottom overlaps the range a real deck lives in.
 */
export const SCALE_SIZES = [50, 200, 500]

/**
 * One timing workload, in the shape {@link import('./programs.mjs').Program} has.
 *
 * The same `id` / `label` / `what` / `build` shape as the bundle corpus on purpose: the
 * timing pass runs both corpora through one loop, and a second shape would mean a second
 * loop that could drift from the first in how it warms up or how it counts.
 * @typedef {object} Workload
 * @property {string} id - kebab-case, stable; the timing table's row key
 * @property {string} label - how the row reads on the page
 * @property {string} what - one line saying what the deck contains
 * @property {number} slides - how many slides, for the per-slide column
 * @property {Record<string, (pres: any) => unknown>} build - per subject; every workload has both arms
 */

/**
 * The deck, once per library, as a function of how many slides it holds.
 *
 * Two literal arms rather than one taking the subject as a parameter, for the reason the
 * other two corpora have two: the page prints these bodies, and a body that branches on
 * which library is running would put `subject === 'ts-pptx'` in front of a reader as though
 * it were a call they should make. The cost of writing the deck twice is that the two can
 * drift; `test/scripts/comparison-timing.test.js` is what stops them, by holding the two
 * bodies against each other and allowing exactly the line that is allowed to differ.
 *
 * That line is `addChart`, which took a divergent signature at the detach: ts-pptx puts
 * `type` in the options object where upstream takes it as the first argument. Every other
 * call is spelled the same on both sides.
 * @type {Record<string, (slides: number) => (pres: any) => void>}
 */
export const SCALE_ARMS = {
	'ts-pptx': (slides) => (pres) => {
		const rows = [
			[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }],
			[{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }],
			[{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }],
			[{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }],
		]
		const bars = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]
		for (let index = 0; index < slides; index++) {
			const slide = pres.addSlide()
			slide.addText('Slide ' + (index + 1), { x: 0.6, y: 0.4, w: 8.8, h: 0.8, fontSize: 24, bold: true })
			slide.addText(
				[
					{ text: 'What we saw', options: { bullet: true, bold: true } },
					{ text: 'What we changed', options: { bullet: true } },
					{ text: 'What it cost', options: { bullet: true } },
				],
				{ x: 0.6, y: 1.4, w: 4, h: 2.4, fontSize: 14 }
			)
			if (index % 2 === 0) slide.addTable(rows, { x: 5, y: 1.4, w: 4.4, colW: [2, 1.2, 1.2], fontSize: 10 })
			else slide.addChart(bars, { type: 'bar', barDir: 'col', x: 5, y: 1.4, w: 4.4, h: 3, showValue: true })
			slide.addNotes('Speaker notes for slide ' + (index + 1) + '.')
		}
	},
	pptxgenjs: (slides) => (pres) => {
		const rows = [
			[{ text: 'Region' }, { text: 'Revenue' }, { text: 'Growth' }],
			[{ text: 'North America' }, { text: '24.9' }, { text: '14.2%' }],
			[{ text: 'EMEA' }, { text: '12.6' }, { text: '16.8%' }],
			[{ text: 'APAC' }, { text: '6.8' }, { text: '9.4%' }],
		]
		const bars = [{ name: 'Revenue', labels: ['Q1', 'Q2', 'Q3', 'Q4'], values: [12, 19, 7, 24] }]
		for (let index = 0; index < slides; index++) {
			const slide = pres.addSlide()
			slide.addText('Slide ' + (index + 1), { x: 0.6, y: 0.4, w: 8.8, h: 0.8, fontSize: 24, bold: true })
			slide.addText(
				[
					{ text: 'What we saw', options: { bullet: true, bold: true } },
					{ text: 'What we changed', options: { bullet: true } },
					{ text: 'What it cost', options: { bullet: true } },
				],
				{ x: 0.6, y: 1.4, w: 4, h: 2.4, fontSize: 14 }
			)
			if (index % 2 === 0) slide.addTable(rows, { x: 5, y: 1.4, w: 4.4, colW: [2, 1.2, 1.2], fontSize: 10 })
			else slide.addChart('bar', bars, { barDir: 'col', x: 5, y: 1.4, w: 4.4, h: 3, showValue: true })
			slide.addNotes('Speaker notes for slide ' + (index + 1) + '.')
		}
	},
}

/**
 * One library's arm, as the page prints it.
 *
 * The slide count is bound away before the body is recovered, so what comes back is the
 * shape rather than one of the three sizes: the sizes differ in nothing else, and printing
 * the same forty lines three times would bury that.
 * @param {string} subject
 * @returns {string}
 */
export function scaleSource(subject) {
	const arm = SCALE_ARMS[subject]
	if (!arm) throw new Error('no timing workload arm registered for subject "' + subject + '"')
	return functionBody(arm(0))
}

/**
 * The deck both libraries build, at one size.
 * @param {number} slides
 * @returns {Workload}
 */
export function scaleWorkload(slides) {
	/** @type {Record<string, (pres: any) => unknown>} */
	const build = {}
	for (const [subject, arm] of Object.entries(SCALE_ARMS)) build[subject] = arm(slides)
	return {
		id: 'scale-' + slides,
		label: slides + ' slides',
		what:
			slides +
			' slides of the same shape: a title, three bullets, a table or a column chart on ' +
			'alternating slides, and speaker notes.',
		slides,
		build,
	}
}

/** @type {Workload[]} */
export const WORKLOADS = SCALE_SIZES.map(scaleWorkload)
