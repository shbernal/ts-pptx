/**
 * Where `tableToSlides` puts a continuation table, and why that is not a separate question from
 * how tall the pager thought it was.
 *
 * Two places state the same rule: `startYEmu` in `gen/table/autopage.ts` decides how much height
 * each page has, and `tableToSlides` decides where on the slide the table is placed. They must
 * agree -- a table placed where it was not measured runs off the slide it fits on paper.
 *
 * They had drifted the same two ways. `autoPageSlideStartY: 0` is a caller asking for the top of
 * the slide, and `opts.autoPageSlideStartY || margin` read it as unset. And a `y` already ABOVE
 * the top margin was pushed back down to the margin on every page after the first, throwing away
 * the space the pager's own `Math.min(y, topMargin)` exists to protect.
 */
import { Window } from 'happy-dom'
import { tableToSlides } from '../../../dist/html.js'
import { assert, assertEqual, build, defineRegressionSuite, listEntries, readEntry } from '../../helpers.js'

const EMU_PER_INCH = 914400

/** A table with enough rows to page several times over. */
const TALL_TABLE = `<table id="t"><tbody>${Array.from(
	{ length: 90 },
	(_, i) => `<tr><td>Row ${i} column A</td><td>Row ${i} column B</td></tr>`
).join('')}</tbody></table>`

/** The `y` (EMU) of the table frame on every emitted slide, in slide order. */
async function frameYs(opts) {
	const win = new Window()
	win.document.body.innerHTML = TALL_TABLE
	const { zip } = await build((pptx) => {
		tableToSlides(pptx, win.document.getElementById('t'), opts)
	})
	const names = listEntries(zip)
		.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
		.sort((a, b) => Number(a.match(/\d+/)[0]) - Number(b.match(/\d+/)[0]))
	const ys = []
	for (const name of names) {
		// The table's own frame, not the first `<a:off>` on the slide -- a slide carries other
		// shapes whose offsets come first in document order.
		const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(await readEntry(zip, name))
		assert(frame, `${name} has no table frame`)
		const match = /<a:off x="-?\d+" y="(-?\d+)"\/>/.exec(frame[0])
		assert(match, `${name}'s table frame has no offset`)
		ys.push(Number(match[1]))
	}
	return ys
}

defineRegressionSuite('tableToSlides continuation start-Y', [
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
			// The other side of that `Math.min`: paging reclaims the space above a low `y`.
			const ys = await frameYs({ y: 1.5 })
			assert(ys.length > 1, `the fixture must page; got ${ys.length} slide(s)`)
			for (const [idx, y] of ys.slice(1).entries())
				assert(y < ys[0], `slide ${idx + 2} must move up from the first page's ${ys[0]}; got ${y}`)
		},
	},
])
