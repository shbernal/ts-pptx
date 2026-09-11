/**
 * The auto-pager wraps a cell at the size the cell is drawn at, whether the size is set on the
 * cell or on the table.
 *
 * `parseTextToLines` computed characters per line from the cell's own `fontSize` or the library
 * default, never the table's, while the row's line height and the emitted `sz` both used the
 * table's. Thirty rows of long text at a table `fontSize: 24` paged as 6 slides of 5 rows; the same
 * rows with `fontSize: 24` on each cell paged as 12 slides of 3. Both emitted `sz="2400"`, so the
 * table-level form ran rows off the bottom of every page.
 */
import { assertEqual, build, defineRegressionSuite, listEntries, readEntry } from '../../helpers.js'

const POS = { x: 0.5, y: 0.5, w: 4 }
const TEXT = 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma'

/** The row count of each page's table, in slide order. */
async function rowsPerPage(cellOptions, tableOptions) {
	const rows = Array.from({ length: 30 }, () => [{ text: TEXT, options: cellOptions }])
	const { zip } = await build((p) => {
		p.addSlide().addTable(rows, { ...POS, autoPage: true, ...tableOptions })
	})
	const slides = listEntries(zip)
		.filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
		.sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]))
	const counts = []
	for (const slide of slides) counts.push(((await readEntry(zip, slide)).match(/<a:tr /g) ?? []).length)
	return counts
}

defineRegressionSuite('Auto-page font size', [
	{
		name: 'a table-level fontSize pages exactly as the same size on every cell',
		fn: async () => {
			const onCells = await rowsPerPage({ fontSize: 24 }, {})
			const onTable = await rowsPerPage({}, { fontSize: 24 })
			assertEqual(JSON.stringify(onTable), JSON.stringify(onCells), 'rows per page, table-level against cell-level')
		},
	},
	{
		name: 'a cell fontSize still wins over the table one',
		fn: async () => {
			const onCells = await rowsPerPage({ fontSize: 24 }, {})
			const both = await rowsPerPage({ fontSize: 24 }, { fontSize: 10 })
			assertEqual(JSON.stringify(both), JSON.stringify(onCells), 'the cell size decides the wrap')
		},
	},
])
