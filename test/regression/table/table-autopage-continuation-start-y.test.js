/**
 * Where `addTable({ autoPage })` puts a continuation table: where the pager budgeted it.
 *
 * The pager decides each page's usable height from where that page's table starts, and the
 * definer then placed the continuation tables with a rule of its own, `autoPageSlideStartY ||
 * margin`. The two disagreed:
 *
 * - `autoPageSlideStartY: 0` read as unset, so continuations landed at the 0.5in margin while each
 *   was paged for a table starting at 0, and the table overran the slide by 0.5in.
 * - A `y` above the top margin was pushed down to the margin on every page after the first, while
 *   the pager's `Math.min(y, margin)` budgeted it from `y`.
 *
 * The same three cases as `html-table-continuation-start-y.test.js`, which covers `tableToSlides`.
 */
import { assert, assertEqual, build, defineRegressionSuite, listEntries, readEntry } from '../../helpers.js'

const EMU_PER_INCH = 914400
const ROWS = Array.from({ length: 90 }, (_unused, i) => [`Row ${i} column A`, `Row ${i} column B`])

/** The `y` (EMU) of the table frame and its row count on every emitted slide, in slide order. */
async function pages(opts) {
	const { zip } = await build((pptx) => {
		pptx.addSlide().addTable(ROWS, { x: 0.5, w: 9, autoPage: true, ...opts })
	})
	const names = listEntries(zip)
		.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
		.sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]))
	const result = []
	for (const name of names) {
		const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(await readEntry(zip, name))
		assert(frame, `${name} has no table frame`)
		const match = /<a:off x="-?\d+" y="(-?\d+)"\/>/.exec(frame[0])
		assert(match, `${name}'s table frame has no offset`)
		result.push({ y: Number(match[1]), rows: (frame[0].match(/<a:tr[ />]/g) ?? []).length })
	}
	return result
}

/** The `y` (EMU) of the table frame on every emitted slide, in slide order. */
async function frameYs(opts) {
	return (await pages(opts)).map((page) => page.y)
}

defineRegressionSuite('addTable continuation start-Y', [
	{
		// The pager budgeted the first page from `y || margin`, so a stated `y: 0` was paged as if it
		// started at the 0.5in margin while the table was placed at 0, leaving the bottom half inch
		// of the first page unused.
		name: 'y: 0 is budgeted from the top of the slide, on the first page as on the rest',
		fn: async () => {
			const top = await pages({ y: 0 })
			const margin = await pages({ y: 0.5 })
			assert(top.length > 1, `the fixture must page; got ${top.length} slide(s)`)
			for (const [idx, { y }] of top.entries()) assertEqual(y, 0, `slide ${idx + 1} must start at 0; got ${y}`)
			assert(
				top[0].rows > margin[0].rows,
				`a first page starting at 0 fits more rows than one at the margin; got ${top[0].rows} and ${margin[0].rows}`
			)
		},
	},
	{
		name: 'autoPageSlideStartY: 0 puts continuations at the top of the slide',
		fn: async () => {
			const ys = await frameYs({ y: 1.5, autoPageSlideStartY: 0 })
			assert(ys.length > 1, `the fixture must page; got ${ys.length} slide(s)`)
			for (const [idx, y] of ys.slice(1).entries()) assertEqual(y, 0, `slide ${idx + 2} must start at 0; got ${y}`)
		},
	},
	{
		name: 'a `y` above the top margin is kept, not pushed down to the margin',
		fn: async () => {
			const ys = await frameYs({ y: 0.1 })
			assert(ys.length > 1, `the fixture must page; got ${ys.length} slide(s)`)
			for (const [idx, y] of ys.slice(1).entries())
				assertEqual(y, Math.round(0.1 * EMU_PER_INCH), `slide ${idx + 2} must keep y=0.1in; got ${y}`)
		},
	},
	{
		name: 'a `y` below the top margin still hands continuations back to the margin',
		fn: async () => {
			const ys = await frameYs({ y: 1.5 })
			assert(ys.length > 1, `the fixture must page; got ${ys.length} slide(s)`)
			for (const [idx, y] of ys.slice(1).entries())
				assertEqual(y, Math.round(0.5 * EMU_PER_INCH), `slide ${idx + 2} must start at the 0.5in margin; got ${y}`)
		},
	},
])
