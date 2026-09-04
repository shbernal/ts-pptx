import { Window } from 'happy-dom'
import { tableToSlides } from '../../../dist/html.js'
import { build, readEntry, assert, assertEqual, defineRegressionSuite } from '../../helpers.js'

// Acceptance: what an imported RAGGED HTML table looks like on the slide.
//
// An HTML table is not a rectangle of cells. A row states only the cells it *starts*, so a
// `colspan` fills several grid columns and a `rowspan` from above fills one the row never
// mentions; rows may also simply be short. `<a:tblGrid>` has no such model -- it declares N
// columns and every `<a:tr>` must carry N `<a:tc>` -- so something has to decide, per row, how
// many columns are already accounted for and how many blanks to append.
//
// `measureGridColumns` is that decision and is unit-tested directly in `html-table-grid.test.js`.
// What had no coverage was the END of the pipe: whether its answer composes with the emitter's
// own `gridSpan`/`vMerge` synthesis into a table that is rectangular, carries the right merges,
// and does not gain or lose a cell of authored text. That gap is the reason the rowspan-occupancy
// rule could not safely be touched -- the byte-identity corpus contains no imported HTML table at
// all, so nothing there would have caught a shift.
//
// Every case asserts the same three things: the grid width, the full cell-text matrix (which
// pins WHERE each authored string lands, not merely that it survived), and rectangularity.

/** A fresh window per case; no global DOM is installed and no state leaks between them. */
function windowWith(html) {
	const win = new Window()
	win.document.body.innerHTML = html
	return win
}

function tableOf(win, id = 't') {
	const table = win.document.getElementById(id)
	assert(table, `fixture is missing #${id}`)
	return table
}

function gridColWidths(xml) {
	return [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((m) => Number(m[1]))
}

/** Cell texts per row, in emitted order. */
function cellTexts(xml) {
	return [...xml.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((row) =>
		[...row[0].matchAll(/<a:tc[\s\S]*?<\/a:tc>/g)].map((cell) =>
			[...cell[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((run) => run[1]).join('|')
		)
	)
}

/** Assert the emitted table is a rectangle: every row carries exactly one cell per grid column. */
function assertRectangular(xml) {
	const cols = gridColWidths(xml).length
	const rows = [...xml.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((row) => [...row[0].matchAll(/<a:tc\b/g)].length)
	assert(cols > 0, `expected a non-empty grid; got: ${xml}`)
	assert(
		rows.every((count) => count === cols),
		`every row must carry ${cols} cells; got ${rows.join(', ')}`
	)
}

/** Convert one HTML table and return its slide XML. */
async function convert(html) {
	const win = windowWith(html)
	const { zip } = await build((pptx) => {
		tableToSlides(pptx, tableOf(win))
	})
	return await readEntry(zip, 'ppt/slides/slide1.xml')
}

/** The shared assertion: grid width, the whole text matrix, and a rectangular result. */
function assertGrid(xml, columns, texts) {
	assertEqual(gridColWidths(xml).length, columns, `expected ${columns} grid columns`)
	assertEqual(JSON.stringify(cellTexts(xml)), JSON.stringify(texts), 'cell texts, per row, in emitted order')
	assertRectangular(xml)
}

defineRegressionSuite('ragged HTML tables', [
	{
		name: 'a short row is padded on the right',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td>a</td><td>b</td><td>c</td></tr>
					<tr><td>d</td></tr>
				</tbody></table>`)
			assertGrid(xml, 3, [
				['a', 'b', 'c'],
				['d', '', ''],
			])
		},
	},
	{
		// The grid is the WIDEST row's reach, not the first row's -- a single spanning header over
		// a wider body is the common shape, and reading the width off row one would truncate it.
		name: 'a body wider than its header widens the grid',
		fn: async () => {
			const xml = await convert(`
				<table id="t">
					<thead><tr><th colspan="2">Wide</th></tr></thead>
					<tbody><tr><td>a</td><td>b</td><td>c</td><td>d</td></tr></tbody>
				</table>`)
			// Four cells in the header row, not two: the column a `gridSpan` covers is still an
			// `<a:tc hMerge="1">` in OOXML, and the two columns past the header's reach are padding.
			assertGrid(xml, 4, [
				['Wide', '', '', ''],
				['a', 'b', 'c', 'd'],
			])
			assert(/<a:tc\b[^>]*\bgridSpan="2"/.test(xml), 'the header keeps its colspan')
		},
	},
	{
		// The case a "sum this row's spans" count gets wrong: the held column comes AFTER the row's
		// own cells, so it is only found by looking past the final one. Counting the second row as
		// short would append a blank and produce a 3-cell row in a 2-column grid.
		name: 'a rowspan in the last column fills the row below without padding it',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td>a</td><td rowspan="2">Tall</td></tr>
					<tr><td>b</td></tr>
				</tbody></table>`)
			assertGrid(xml, 2, [
				['a', 'Tall'],
				['b', ''],
			])
			assert(/<a:tc\b[^>]*\bvMerge="1"/.test(xml), 'the continuation is a vMerge, not a padded blank')
		},
	},
	{
		name: 'a rowspan deeper than one row keeps holding its column',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td rowspan="3">Tall</td><td>a</td></tr>
					<tr><td>b</td></tr>
					<tr><td>c</td></tr>
					<tr><td>d</td><td>e</td></tr>
				</tbody></table>`)
			assertGrid(xml, 2, [
				['Tall', 'a'],
				['', 'b'],
				['', 'c'],
				['d', 'e'],
			])
			assert(/<a:tc\b[^>]*\browSpan="3"/.test(xml), 'the origin carries rowSpan="3"')
		},
	},
	{
		// Both rules at once: the row below is short AND partly held from above, so the padding
		// count is neither "grid width minus cells stated" nor "grid width minus columns held".
		name: 'a row that is both held from above and short is padded by the difference',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td rowspan="2">Tall</td><td>a</td><td>b</td><td>c</td></tr>
					<tr><td>d</td></tr>
				</tbody></table>`)
			assertGrid(xml, 4, [
				['Tall', 'a', 'b', 'c'],
				['', 'd', '', ''],
			])
		},
	},
	{
		// A rowspan and a colspan starting in the same cell: the hold covers both columns, so the
		// row below skips two and starts at column 2.
		name: 'a cell spanning both ways holds every column it covers',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td colspan="2" rowspan="2">Block</td><td>a</td></tr>
					<tr><td>b</td></tr>
					<tr><td>c</td><td>d</td><td>e</td></tr>
				</tbody></table>`)
			assertGrid(xml, 3, [
				['Block', '', 'a'],
				['', '', 'b'],
				['c', 'd', 'e'],
			])
			assert(/gridSpan="2"[^>]*rowSpan="2"|rowSpan="2"[^>]*gridSpan="2"/.test(xml), 'the origin carries both spans')
		},
	},
	{
		// Two independent rowspans of different depths, so the columns free up on different rows.
		// A single shared "rows remaining" counter instead of one per column gets this wrong.
		name: 'rowspans of different depths free their columns independently',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td rowspan="3">A3</td><td rowspan="2">B2</td><td>c</td></tr>
					<tr><td>d</td></tr>
					<tr><td>e</td><td>f</td></tr>
				</tbody></table>`)
			assertGrid(xml, 3, [
				['A3', 'B2', 'c'],
				['', '', 'd'],
				['', 'e', 'f'],
			])
		},
	},
	{
		// A rowspan reaching past the last row is the input a clamp exists for: held rows that do
		// not exist must not shift anything, and the emitted span must stay inside the table.
		name: 'a rowspan past the last row does not reach outside the table',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td>a</td><td rowspan="9">Tall</td></tr>
					<tr><td>b</td></tr>
				</tbody></table>`)
			assertGrid(xml, 2, [
				['a', 'Tall'],
				['b', ''],
			])
		},
	},
	{
		// A row of nothing but continuations states no cells at all. Its reach comes entirely from
		// the holds above it, so a row-length-based count would read it as empty and pad it to full
		// width -- doubling its cells.
		name: 'a row made only of continuations is already full',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td rowspan="2">A</td><td rowspan="2">B</td></tr>
					<tr></tr>
					<tr><td>c</td><td>d</td></tr>
				</tbody></table>`)
			assertGrid(xml, 2, [
				['A', 'B'],
				['', ''],
				['c', 'd'],
			])
		},
	},
	{
		// A nonsense span must cost its own cell nothing and the rest of the grid nothing: read
		// literally, a zero or negative colspan walks the column cursor backwards and shifts every
		// column after it.
		name: 'a nonsense span leaves the grid unshifted',
		fn: async () => {
			const xml = await convert(`
				<table id="t"><tbody>
					<tr><td colspan="0">a</td><td colspan="-2">b</td><td rowspan="abc">c</td></tr>
					<tr><td>d</td><td>e</td><td>f</td></tr>
				</tbody></table>`)
			assertGrid(xml, 3, [
				['a', 'b', 'c'],
				['d', 'e', 'f'],
			])
		},
	},
])
