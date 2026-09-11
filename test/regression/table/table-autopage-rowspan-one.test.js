/**
 * `rowspan: 1` is "no span", and the auto-pager prices the row like any other.
 *
 * The pager exempted a cell from the row's line height when its `rowspan` option was truthy, so an
 * explicit `rowspan: 1` -- valid, and passed through by the span check -- priced its row at the
 * cell margins alone. Eighty two-cell rows paged onto 6 slides; the same rows with `rowspan: 1` on
 * every cell paged onto 2, with 46 rows on the first. The measured-fit layout already tests the
 * span itself (`rowSpan === 1`), so the two disagreed.
 */
import { assertEqual, build, defineRegressionSuite, listEntries } from '../../helpers.js'

/** How many slides 80 two-cell rows page onto, with `cellOptions` on every cell. */
async function pages(cellOptions) {
	const rows = Array.from({ length: 80 }, (_unused, i) => [
		{ text: `row ${i}`, options: cellOptions },
		{ text: 'value', options: cellOptions },
	])
	const { zip } = await build((p) => {
		p.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 6, autoPage: true })
	})
	return listEntries(zip).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).length
}

defineRegressionSuite('Auto-page rowspan 1', [
	{
		name: 'rows stating rowspan 1 page exactly as rows stating nothing',
		fn: async () => {
			assertEqual(await pages({ rowspan: 1 }), await pages({}), 'slides with rowspan: 1 against none')
		},
	},
])
