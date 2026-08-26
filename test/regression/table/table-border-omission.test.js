import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'

const SLIDE_XML = 'ppt/slides/slide1.xml'

// Regression for upstream #23: a null side in a TRBL border tuple is *omitted* — it
// inherits from the built-in table style / theme — while `{ type: 'none' }` is an
// explicit "erase this edge". `normalizeTableRows` used to replace every falsy side
// with an explicit `<a:lnX><a:noFill/></a:lnX>`, so there was no spelling left for
// "draw only these edges", and the holes silently overrode style inheritance.
//
// This mirrors the distinction `normalizeOuterBorder` already documents for the
// table perimeter ("a sparse side is NOT `{ type: 'none' }`"); the cell path now
// agrees with it.
//
// The separate question of whether a cell with *nothing authored* should keep
// receiving the forced four-side no-fill default is tracked on #23 as well — the
// force-fill is deliberate today (it keeps unstyled tables free of grid lines), so
// this suite pins that behavior rather than changing it.
defineRegressionSuite('Table border tuple null sides [upstream-23]', [
	{
		name: 'null tuple sides stay absent while authored sides draw',
		fn: async () => {
			const solid = { type: 'solid', color: '4472C4', width: 1 }
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'sparse', options: { border: [solid, null, solid, null] } }]], {
					x: 1,
					y: 1,
					w: 4,
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(
				(xml.match(/<a:ln[LRTB]/g) ?? []).length === 2,
				`expected exactly two drawn edges; got: ${xml.slice(0, 900)}`
			)
			assert(xml.includes('<a:lnT'), 'top edge drawn')
			assert(xml.includes('<a:lnB'), 'bottom edge drawn')
			assert(!xml.includes('<a:lnL'), 'left side absent (inherits)')
			assert(!xml.includes('<a:lnR'), 'right side absent (inherits)')
		},
	},
	{
		name: 'explicit type:none still emits <a:noFill/> on all four sides',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'cleared', options: { border: { type: 'none' } } }]], { x: 1, y: 1, w: 4 })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(
				(xml.match(/<a:noFill\/>/g) ?? []).length >= 4,
				`expected explicit no-fill sides; got: ${xml.slice(0, 900)}`
			)
		},
	},
])
