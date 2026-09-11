/**
 * Where `tableToSlides` puts a table across the slide, and how wide it makes it.
 *
 * The width was the slide width (or `w`) less both slide margins. That ignored `x`, the bug
 * `usableTableWidthEmu` exists to fix and that `addTable` already reads correctly, and it shrank a
 * width the caller stated. On a 10in slide `{ x: 3 }` gave a 9in table running to 12in, `{ w: 8 }`
 * gave 7in, and `x: 0` was read as unset and placed the table at the 0.5in margin.
 *
 * This is arithmetic on stated options, not DOM measurement, so it holds under any DOM.
 */
import { Window } from 'happy-dom'
import { tableToSlides } from '../../../dist/html.js'
import { assert, assertEqual, build, defineRegressionSuite, readEntry } from '../../helpers.js'

const EMU_PER_INCH = 914400
const inches = (emu) => Math.round((emu / EMU_PER_INCH) * 1000) / 1000

/** The table frame's `x` and `cx` in inches, and the slide width, for one set of options. */
async function frameOf(opts) {
	const win = new Window()
	win.document.body.innerHTML = '<table id="t"><tbody><tr><td>a</td><td>b</td></tr></tbody></table>'
	const { zip, pres } = await build((pptx) => {
		tableToSlides(pptx, win.document.getElementById('t'), opts)
	})
	const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/.exec(await readEntry(zip, 'ppt/slides/slide1.xml'))
	assert(frame, 'the slide has a table frame')
	const off = /<a:off x="(-?\d+)" y="-?\d+"\/>/.exec(frame[0])
	const ext = /<a:ext cx="(\d+)" cy="\d+"\/>/.exec(frame[0])
	assert(off && ext, 'the frame has an offset and an extent')
	return { x: inches(Number(off[1])), cx: inches(Number(ext[1])), slideW: inches(pres.presLayout.width) }
}

defineRegressionSuite('tableToSlides table width and x', [
	{
		name: 'with neither x nor w, the table fills the slide between the margins',
		fn: async () => {
			const f = await frameOf({})
			assertEqual(f.x, 0.5, 'x at the left margin')
			assertEqual(f.cx, f.slideW - 1, 'the slide less both margins')
		},
	},
	{
		name: 'a stated x moves the table and narrows it to the right margin',
		fn: async () => {
			const f = await frameOf({ x: 3 })
			assertEqual(f.x, 3, 'x as stated')
			assertEqual(f.cx, f.slideW - 3 - 0.5, 'from x to the right margin')
		},
	},
	{
		name: 'a stated w is the width, not the width less the margins',
		fn: async () => {
			const f = await frameOf({ w: 8 })
			assertEqual(f.cx, 8, 'w as stated')
			const both = await frameOf({ x: 1, w: 4 })
			assertEqual(both.x, 1, 'x as stated alongside w')
			assertEqual(both.cx, 4, 'w as stated alongside x')
		},
	},
	{
		name: 'x: 0 is the left edge of the slide, and the table stays on the slide',
		fn: async () => {
			const f = await frameOf({ x: 0 })
			assertEqual(f.x, 0, 'x at the slide edge')
			assert(f.x + f.cx <= f.slideW - 0.5, `the right edge stays inside the right margin; got ${f.x + f.cx}`)
		},
	},
])
