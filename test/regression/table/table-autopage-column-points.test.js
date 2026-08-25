import { defineRegressionSuite, build, readEntry, listEntries, assert, assertEqual } from '../../helpers.js'

// The chars-per-line figure the auto-pager wraps on comes from the column's width in POINTS, and
// that conversion used to take a detour: `(colWidth / EMU_PER_POINT) * EMU_PER_INCH` is the same
// 72 as `colWidth * POINTS_PER_INCH`, with two EMU constants that cancel — except the detour is
// not exact in binary floating point. `6.625 * 72` is exactly 477, but `6.625 / 12700 * 914400`
// evaluates to 476.9999…, and the `Math.floor` around it took the hit.
//
// Over the 200,000 widths from 0.001in to 200in the two forms disagree on 51, always by one point
// low. One point of chars-per-line usually changes nothing, since the wrap test compares an
// integer character count against it. The case below is one where it does, and the consequence is
// not cosmetic: at 6.625in and 15pt the figure is 73.14 corrected against 72.99 as it was, so a
// 73-character run wraps to a second line it does not need, every row of the table is measured as
// twice as tall as it is, and a sixteen-row table that fits on one slide is paged onto two.
//
// The measurement is read through the pager's own output — how many `<a:tr>` rows land on each
// slide — because that is what the line count is *for*. A cell's own `<a:p>` count does not track
// it: the text is emitted as one paragraph and PowerPoint re-wraps it at render time.

/** Every `ppt/slides/slideN.xml`, in slide order. */
function slideFiles(zip) {
	return listEntries(zip)
		.filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
		.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}

/**
 * A run of `n` characters as space-separated four-letter words, so the wrapper has break
 * opportunities to take.
 */
function words(n) {
	let out = ''
	while (out.length < n) out += (out ? ' ' : '') + 'abcd'
	return out.slice(0, n)
}

/**
 * How the auto-pager distributes `rows` identical single-cell rows across slides.
 *
 * `fontSize` goes on the CELL, not on the table: `parseTextToLines` reads
 * `cell.options.fontSize` and falls back to the default, and the table-level option is not one of
 * the two the pager copies down.
 */
async function rowsPerSlide(text, colWidth, fontSize, rows) {
	const { zip } = await build((p) => {
		const data = Array.from({ length: rows }, () => [{ text, options: { fontSize } }])
		p.addSlide().addTable(data, { x: 0.25, y: 0.25, colW: [colWidth], autoPage: true, margin: 0, slideMargin: 0 })
	})
	const out = []
	for (const file of slideFiles(zip)) {
		const xml = await readEntry(zip, file)
		out.push((xml.match(/<a:tr /g) || []).length)
	}
	assert(out.length > 0, 'expected at least one slide')
	return out
}

defineRegressionSuite('Auto-page column width in points', [
	{
		name: 'a 6.625in column at 15pt fits 73 characters on one line, so sixteen rows fit on one slide',
		fn: async () => {
			assertEqual(
				JSON.stringify(await rowsPerSlide(words(73), 6.625, 15, 16)),
				'[16]',
				'all sixteen rows on one slide; the old arithmetic split them 11 and 6'
			)
			// The other side of the threshold, so the assertion above cannot pass merely because
			// the pager stopped paging.
			assertEqual(
				JSON.stringify(await rowsPerSlide(words(93), 6.625, 15, 16)),
				'[11,6]',
				'93 characters genuinely need two lines, and then sixteen rows do not fit'
			)
		},
	},
	{
		// A width both forms agree on, so this case is unaffected by the fix and stays a control:
		// if it moves, the conversion changed for everything rather than for the 51.
		name: 'a width both forms agree on is unchanged',
		fn: async () => {
			assertEqual(
				JSON.stringify(await rowsPerSlide(words(73), 7, 15, 16)),
				'[16]',
				'7in is 504 points under either form'
			)
		},
	},
])
