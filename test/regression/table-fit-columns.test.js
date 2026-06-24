import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Acceptance for `fitColumns: 'shrink'` (#1451): an explicit `colW` array (or a `w`)
// wider than the space between the table's `x` and the right slide margin is scaled
// down proportionally so the whole table fits the slide. Opt-in: without the flag the
// over-wide widths are emitted as-is. Default layout is 10in wide; default margin 0.5in,
// so `x: 0.5` leaves 9in of usable width.

const EMU = 914400
const gridColsEmu = (xml) => [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]))
const sumEmu = (cols) => cols.reduce((a, b) => a + b, 0)

defineRegressionSuite('Table fitColumns shrink-to-fit', 'upstream-issue-1451', [
	{
		name: 'without fitColumns, an over-wide colW array overflows the slide (opt-in guard)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'a' }, { text: 'b' }, { text: 'c' }, { text: 'd' }]], {
					x: 0.5,
					y: 1,
					colW: [4, 4, 4, 4],
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColsEmu(xml)
			assert(cols.length === 4, 'expected 4 gridCol; got ' + cols.length)
			// Untouched: each column stays 4in, summing to 16in (> 9in usable) — runs off the slide.
			assert(
				cols.every((w) => w === 4 * EMU),
				'expected unscaled 4in columns; got ' + cols.join(',')
			)
		},
	},
	{
		name: 'fitColumns shrinks an over-wide equal colW array to the usable width',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'a' }, { text: 'b' }, { text: 'c' }, { text: 'd' }]], {
					x: 0.5,
					y: 1,
					colW: [4, 4, 4, 4],
					fitColumns: 'shrink',
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColsEmu(xml)
			// 16in scaled to the 9in usable width (10 - 0.5 x - 0.5 right margin).
			assert(
				Math.abs(sumEmu(cols) - 9 * EMU) <= cols.length,
				'expected columns to sum to ~9in; got ' + sumEmu(cols) / EMU
			)
			assert(
				cols.every((w) => Math.abs(w - cols[0]) <= 1),
				'expected equal columns to stay equal; got ' + cols.join(',')
			)
		},
	},
	{
		name: 'fitColumns preserves the ratio between unequal columns',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'a' }, { text: 'b' }, { text: 'c' }]], {
					x: 0.5,
					y: 1,
					colW: [6, 2, 2], // sum 10in > 9in usable -> factor 0.9 -> [5.4, 1.8, 1.8]
					fitColumns: 'shrink',
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColsEmu(xml)
			assert(Math.abs(sumEmu(cols) - 9 * EMU) <= cols.length, 'expected sum ~9in; got ' + sumEmu(cols) / EMU)
			// First column stays 3x each of the others (6:2:2 ratio preserved).
			assert(Math.abs(cols[0] - 3 * cols[1]) <= 3, 'expected 3:1 ratio preserved; got ' + cols.join(','))
			assert(Math.abs(cols[1] - cols[2]) <= 1, 'expected last two columns equal; got ' + cols.join(','))
		},
	},
	{
		name: 'fitColumns leaves an already-fitting colW array untouched (shrink only, never grows)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'a' }, { text: 'b' }]], {
					x: 0.5,
					y: 1,
					colW: [2, 3], // sum 5in < 9in usable
					fitColumns: 'shrink',
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColsEmu(xml)
			assert(cols[0] === 2 * EMU && cols[1] === 3 * EMU, 'expected fitting columns unchanged; got ' + cols.join(','))
		},
	},
	{
		name: 'fitColumns clamps an over-wide `w` (no colW) to the usable width',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addTable([[{ text: 'a' }, { text: 'b' }, { text: 'c' }]], {
					x: 0.5,
					y: 1,
					w: 20, // far wider than the 9in usable width
					fitColumns: 'shrink',
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColsEmu(xml)
			// Even distribution of the clamped 9in width across 3 columns.
			assert(Math.abs(sumEmu(cols) - 9 * EMU) <= cols.length, 'expected clamped width ~9in; got ' + sumEmu(cols) / EMU)
		},
	},
])
