import { defineRegressionSuite, build, readEntry, assert, assertEqual } from '../../helpers.js'

/**
 * Run `fn` with `console.log` captured, returning the lines it emitted. Restoring in a
 * `finally` matters: a throwing build must not leave the rest of the suite stubbed.
 */
async function captureLog(fn) {
	const orig = console.log
	const lines = []
	console.log = (...args) => lines.push(args.map((a) => (typeof a === 'string' ? a : String(a))).join(' '))
	try {
		await fn()
	} finally {
		console.log = orig
	}
	return lines
}

// The auto-pager's width arithmetic — the three ways `getSlidesForTableRows`
// (src/gen/table/autopage.ts) used to price a column wrongly.
//
// All three produce wrong output rather than an error, and the emitted grid is the
// only place the first one shows: the pager writes its distribution back onto
// `tableProps.colW`, and `addTableDefinition` carries that into every paged table.

const ONE_IN_EMU = 914400

function gridColWidths(xml) {
	return [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]))
}

/** Enough rows to make the pager page. */
function bodyRows(n, cols) {
	return Array.from({ length: n }, (_, i) => Array.from({ length: cols }, (_, c) => ({ text: `r${i}c${c}` })))
}

defineRegressionSuite('Table autoPage width arithmetic', [
	{
		// The fallback arm used to ADD the left and right slide margins, giving ~1in of
		// usable width however wide the slide was, then split THAT across the columns.
		name: 'usable width is the slide minus its margins, not the sum of them',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				// `colW.length !== column count` makes addTableDefinition warn, drop `colW`,
				// and set no `w` — the easiest route to the pager's fallback width.
				s.addTable(bodyRows(3, 3), { x: 0.5, y: 0.5, colW: [1, 2], autoPage: true, fontSize: 12 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const cols = gridColWidths(xml)
			assertEqual(cols.length, 3, 'three grid columns')
			// 10in slide − 0.5in x − 0.5in right margin = 9in over 3 columns.
			cols.forEach((w) => assert(w > 2 * ONE_IN_EMU, `expected each gridCol > 2in EMU; got ${w}`))
			// The old arithmetic emitted a third of an inch per column.
			assert(!cols.includes(304800), `gridCol must not be the sliver width; got ${JSON.stringify(cols)}`)
		},
	},
	{
		// `idx < idx + cellColspan` is true for every positive span, so the filter was just
		// `idx >= iCell` and a spanning cell was measured against every column from its own
		// cell index onward. Indexing by cell position is the second half of it: every cell
		// after a span read the wrong column's width.
		//
		// The pager's own `[0/4] colWidth=…` dump is the direct oracle here — the wrapped
		// line count is downstream of it, and the emitted `<a:tr h>` is not (the writer
		// splits an explicit `h` evenly and never sees the measurement).
		name: 'each cell is measured against the grid columns it actually spans',
		fn: async () => {
			const rows = [
				[{ text: 'span', options: { colspan: 2 } }, { text: 'c' }, { text: 'd' }, { text: 'e' }],
				[{ text: 'a' }, { text: 'b' }, { text: 'c' }, { text: 'd' }, { text: 'e' }],
			]
			const lines = await captureLog(async () => {
				await build((p) => {
					const s = p.addSlide()
					s.addTable(rows, { x: 0.5, y: 0.5, w: 10, colW: [1, 2, 3, 4, 5], autoPage: true, verbose: true })
				})
			})
			const widths = lines
				.filter((line) => line.startsWith('[0/4] colWidth='))
				.map((line) => Number(/colWidth=([\d.]+)in/.exec(line)[1]))
			// Row 0: the colspan-2 cell covers columns 0+1 (1+2=3in), then 3, 4, 5.
			// Row 1: one cell per column, 1 through 5.
			assertEqual(
				JSON.stringify(widths),
				JSON.stringify([3, 3, 4, 5, 1, 2, 3, 4, 5]),
				`each cell measures its own columns; got ${JSON.stringify(widths)}`
			)
		},
	},
	{
		// `.reduce` with no seed throws "Reduce of empty array with no initial value"
		// the moment a row is longer than the grid the first row defined.
		name: 'a row longer than the first row does not throw',
		fn: async () => {
			const rows = [
				[{ text: 'a' }, { text: 'b' }],
				[{ text: 'c' }, { text: 'd' }, { text: 'e', options: { colspan: 2 } }],
			]
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(rows, { x: 0.5, y: 0.5, w: 9, h: 5, autoPage: true, fontSize: 12 })
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('<a:tbl>'), 'the table was emitted')
		},
	},
])
