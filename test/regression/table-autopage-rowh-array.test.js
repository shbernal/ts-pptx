import { defineRegressionSuite, build, readEntry, listEntries, assert } from '../helpers.js'

// Regression (upstream gitbrent/PptxGenJS#1145): a `rowH` *array* is keyed by the ORIGINAL row
// index. Auto-paging splits rows across slides (and can repeat the header row), so applying the
// array by physical row index on each generated slide is wrong: the height configured for original
// row 0 would land on the *first row of every overflow slide* instead of following its source row.
// The auto-pager now carries each output row's resolved height so a configured height follows its
// row across pages, and overflow rows fall back to auto height where none was configured.

function rowHeightsEmu(xml) {
	return (xml.match(/<a:tr h="(\d+)"/g) || []).map((m) => Number(/h="(\d+)"/.exec(m)[1]))
}

function slideXmls(zip) {
	return listEntries(zip)
		.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
		.sort((a, b) => Number(/slide(\d+)/.exec(a)[1]) - Number(/slide(\d+)/.exec(b)[1]))
}

// A tall first row (2") then 39 short rows. EMU: 1" = 914400.
const TALL_EMU = 2 * 914400
const SHORT_EMU = 0.3 * 914400

function makeRows(n) {
	return Array.from({ length: n }, (_, i) => [
		{ text: `row ${i} lorem ipsum dolor sit amet consectetur adipiscing elit` },
	])
}

defineRegressionSuite('Table autoPage rowH array follows original rows (#1145)', [
	{
		name: 'tall first-row height does not repeat on every overflow slide',
		fn: async () => {
			const rows = makeRows(40)
			const rowH = rows.map((_, i) => (i === 0 ? 2 : 0.3))
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 5, autoPage: true, rowH })
			})
			const slides = slideXmls(zip)
			assert(slides.length >= 2, `expected pagination across multiple slides, got ${slides.length}`)

			// Slide 1 owns original row 0 → its first row must be the configured 2".
			const first = rowHeightsEmu(await readEntry(zip, slides[0]))
			assert(first[0] === TALL_EMU, `slide1 first row should be 2" (${TALL_EMU}); got ${first[0]}`)

			// Every later slide holds short rows only → no row may carry the tall height.
			for (let i = 1; i < slides.length; i++) {
				const heights = rowHeightsEmu(await readEntry(zip, slides[i]))
				assert(
					!heights.includes(TALL_EMU),
					`${slides[i]} must not repeat the tall first-row height; got ${JSON.stringify(heights)}`
				)
				assert(
					heights.every((h) => h === SHORT_EMU),
					`${slides[i]} rows should all be the configured 0.3"; got ${JSON.stringify(heights)}`
				)
			}
		},
	},
	{
		name: 'repeated header keeps its configured height; body rows keep theirs',
		fn: async () => {
			const rows = [[{ text: 'HEADER', options: { bold: true } }], ...makeRows(40)]
			// header (row 0) = 1.5", all body rows = 0.3"
			const HDR_EMU = 1.5 * 914400
			const rowH = rows.map((_, i) => (i === 0 ? 1.5 : 0.3))
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 5, autoPage: true, autoPageRepeatHeader: true, rowH })
			})
			const slides = slideXmls(zip)
			assert(slides.length >= 2, `expected pagination across multiple slides, got ${slides.length}`)

			for (const s of slides) {
				const heights = rowHeightsEmu(await readEntry(zip, s))
				// First row of each slide is the repeated header → its configured 1.5".
				assert(heights[0] === HDR_EMU, `${s} header row should be 1.5" (${HDR_EMU}); got ${heights[0]}`)
				// Remaining rows are body rows → configured 0.3".
				assert(
					heights.slice(1).every((h) => h === SHORT_EMU),
					`${s} body rows should be 0.3"; got ${JSON.stringify(heights)}`
				)
			}
		},
	},
])
