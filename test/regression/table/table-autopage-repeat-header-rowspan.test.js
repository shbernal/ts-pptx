/**
 * A repeated header costs a continuation page what the same rows cost the first page, including
 * a header row that sits inside a rowspan.
 *
 * The repeated header rows were priced by a second copy of the pager's row measure, with its own
 * column cursor. That cursor ignored the columns a rowspan holds, so the second row of a header
 * whose first cell spans both rows was measured one column to the left: here, its long label was
 * wrapped against the 4in column instead of its own 0.8in one, priced at one line instead of
 * several, and every continuation page took more body rows than it had room for.
 *
 * The oracle is the one `table-autopage-continuation-budget.test.js` uses. Every body row is
 * identical and `autoPageSlideStartY` equals `y`, so every page has the same usable height and the
 * same header; a full continuation page therefore holds exactly as many body rows as the first.
 */
import { assert, assertEqual, build, defineRegressionSuite, listEntries, readEntry } from '../../helpers.js'

const LONG = 'a header label long enough to wrap onto several lines in a narrow column'

async function rowsPerSlide() {
	const header = [
		[{ text: 'Group', options: { rowspan: 2 } }, { text: 'A' }, { text: 'B' }],
		[{ text: LONG }, { text: 'b' }],
	]
	const body = Array.from({ length: 80 }, (_unused, i) => [{ text: `r${i}` }, { text: 'v' }, { text: 'w' }])
	const { zip } = await build((p) => {
		p.addSlide().addTable([...header, ...body], {
			x: 0.5,
			y: 0.5,
			colW: [4, 0.8, 0.8],
			autoPage: true,
			autoPageRepeatHeader: true,
			autoPageHeaderRows: 2,
			autoPageSlideStartY: 0.5,
		})
	})
	const slides = listEntries(zip)
		.filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
		.sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
	const counts = []
	for (const name of slides) counts.push(((await readEntry(zip, name)).match(/<a:tr /g) || []).length)
	return counts
}

defineRegressionSuite('Auto-page repeated header inside a rowspan', [
	{
		name: 'a full continuation page holds as many body rows as the first page',
		fn: async () => {
			const counts = await rowsPerSlide()
			assert(counts.length >= 3, `the table has to page at least twice to compare; got ${JSON.stringify(counts)}`)
			// Every page carries the two header rows; the last page is not full.
			const full = counts.slice(0, -1)
			for (const [idx, count] of full.entries())
				assertEqual(count, full[0], `page ${idx + 1} rows against page 1; all pages ${JSON.stringify(counts)}`)
		},
	},
])
