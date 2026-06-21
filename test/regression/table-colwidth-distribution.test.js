import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Acceptance: when a table is sized with `w` (or nothing) but no explicit `colW`,
// the emitted <a:gridCol w=…> must be the table width split evenly in EMU — not the
// raw inches value used as EMU. The historical bug divided inches (`w=9`) and emitted
// `gridCol w="3"` (≈0 EMU), collapsing every auto-width table to a sliver.

const ONE_IN_EMU = 914400

function gridColWidths(xml) {
	return [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]))
}

defineRegressionSuite('Table column-width distribution', [
	{
		name: '`w` without `colW` splits the EMU width evenly across columns',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B', 'C']], { x: 0.5, y: 0.5, w: 9, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColWidths(xml)
			assert(cols.length === 3, `expected 3 gridCols; got ${cols.length}`)
			const expected = Math.round((9 * ONE_IN_EMU) / 3) // 3 inches per column
			cols.forEach((w) => assert(w === expected, `expected gridCol w=${expected} EMU; got ${w}`))
			// Regression guard: the old bug emitted w="3" (raw inches as EMU).
			assert(!cols.includes(3), 'gridCol must not be the raw inches value treated as EMU (w="3")')
		},
	},
	{
		name: 'neither `w` nor `colW` (default full-slide width) still yields inch-scale columns',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B', 'C', 'D']], { x: 0.5, y: 0.5, h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColWidths(xml)
			assert(cols.length === 4, `expected 4 gridCols; got ${cols.length}`)
			// Each column should be a sane fraction of the slide width, not ~0 EMU.
			cols.forEach((w) => assert(w > ONE_IN_EMU, `expected each gridCol > 1in EMU; got ${w}`))
		},
	},
	{
		name: 'explicit `colW` array is still honored per column (inches → EMU)',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([['A', 'B', 'C']], { x: 0.5, y: 0.5, colW: [2, 3, 4], h: 1 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColWidths(xml)
			assert(
				cols.length === 3 && cols[0] === 2 * ONE_IN_EMU && cols[1] === 3 * ONE_IN_EMU && cols[2] === 4 * ONE_IN_EMU,
				`expected [2in,3in,4in] EMU; got ${JSON.stringify(cols)}`
			)
		},
	},
])
