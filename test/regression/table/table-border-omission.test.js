import { defineRegressionSuite, build, readEntry, assert } from '../../helpers.js'
import { TableStyle } from '../../../dist/node.js'

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
// The same distinction one level up settles the other half of #23: a cell with
// *nothing* authored used to receive a forced four-side no-fill, which is direct
// formatting and so erased the grid of the very `tableStyle` the caller selected.
// The default now applies only to a table that named no style, where it is what keeps
// an unstyled table free of PowerPoint's no-style black hairline grid.
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
	{
		name: 'a styled table with no border authored emits no cell edges at all',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([['A', 'B']], {
					x: 1,
					y: 1,
					w: 6,
					tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1,
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(!/<a:ln[LRTB]/.test(xml), `expected the style to paint the grid; got: ${xml.slice(0, 900)}`)
		},
	},
	{
		name: 'an unstyled table still takes the four-side no-fill default',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([['A', 'B']], { x: 1, y: 1, w: 6 })
			})
			const xml = await readEntry(zip, SLIDE_XML)
			assert(
				(xml.match(/<a:ln[LRTB]/g) ?? []).length === 8,
				`expected four edges on each of the two cells; got: ${xml.slice(0, 900)}`
			)
		},
	},
	{
		name: 'a styled table with an outerBorder draws the perimeter and leaves the interior to the style',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([['A', 'B']], {
					x: 1,
					y: 1,
					w: 6,
					tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1,
					outerBorder: { type: 'solid', color: '1A2B3C', width: 1 },
				})
			})
			const xml = await readEntry(zip, SLIDE_XML)
			// Two cells side by side: each is on the top, bottom and one vertical edge, and the
			// two interior verticals stay absent rather than being spelled out as no-fill.
			assert((xml.match(/<a:ln[LRTB]/g) ?? []).length === 6, `expected six perimeter edges; got: ${xml.slice(0, 900)}`)
			assert(!/<a:noFill\/>/.test(xml), `no edge should be an explicit no-fill; got: ${xml.slice(0, 900)}`)
		},
	},
])
