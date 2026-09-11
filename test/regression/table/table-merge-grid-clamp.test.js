/**
 * The table emitter writes the grid the auto-pager and `tableLayout()` measure.
 *
 * The emitter built its merge grid by splicing covered cells into the authored rows, reading each
 * span as authored. `walkTableGrid`, which the pager and the measured fit read, clamps a colspan to
 * the grid's width and a rowspan to the rows left, so the table they described was not the one
 * written:
 *
 * - `[[{ text: 'A', rowspan: 3 }, 'B'], ['C']]` wrote `rowSpan="3"` into a two-row table.
 * - `[['A', 'B'], [{ text: 'D', colspan: 3 }]]` wrote three `<a:tc>` with `gridSpan="3"` against
 *   two `<a:gridCol>`.
 *
 * Either one is a table PowerPoint offers to repair.
 */
import { assert, assertEqual, captureDiagnostics, defineRegressionSuite, build, readEntry } from '../../helpers.js'

/** Each row's `<a:tc>` count, the `<a:gridCol>` count, and every span attribute with its position. */
async function structure(rows) {
	const { result, codes } = await captureDiagnostics(() =>
		build((p) => p.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 6 }))
	)
	const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
	const tbl = /<a:tbl>[\s\S]*<\/a:tbl>/.exec(xml)?.[0] ?? ''
	const gridCols = (tbl.match(/<a:gridCol /g) ?? []).length
	const rowXml = tbl.split('</a:tr>').filter((r) => r.includes('<a:tr'))
	const cellCounts = rowXml.map((r) => (r.match(/<a:tc[ >]/g) ?? []).length)
	const spans = rowXml.flatMap((r, rIdx) =>
		(r.match(/<a:tc[ >][^>]*>/g) ?? []).map((tc, cIdx) => ({
			rIdx,
			cIdx,
			rowSpan: Number(/rowSpan="(\d+)"/.exec(tc)?.[1] ?? 1),
			gridSpan: Number(/gridSpan="(\d+)"/.exec(tc)?.[1] ?? 1),
		}))
	)
	return { gridCols, rowCount: rowXml.length, cellCounts, spans, codes }
}

/** Every row carries one `<a:tc>` per grid column, and no span reaches past the table's edge. */
function assertRectangular({ gridCols, rowCount, cellCounts, spans }, label) {
	cellCounts.forEach((count, rIdx) =>
		assertEqual(count, gridCols, `${label}: row ${rIdx} has one <a:tc> per <a:gridCol>`)
	)
	for (const { rIdx, cIdx, rowSpan, gridSpan } of spans) {
		assert(
			rIdx + rowSpan <= rowCount,
			`${label}: (${rIdx},${cIdx}) rowSpan ${rowSpan} reaches past row ${rowCount - 1}`
		)
		assert(
			cIdx + gridSpan <= gridCols,
			`${label}: (${rIdx},${cIdx}) gridSpan ${gridSpan} reaches past column ${gridCols - 1}`
		)
	}
}

defineRegressionSuite('Table merge grid clamps spans to the table', [
	{
		name: 'a rowspan past the last row is written as far as the table goes',
		fn: async () => {
			const s = await structure([[{ text: 'A', options: { rowspan: 3 } }, 'B'], ['C']])
			assertEqual(s.rowCount, 2, 'two rows')
			assertRectangular(s, 'rowspan 3')
			assert(
				s.spans.some((span) => span.rIdx === 0 && span.cIdx === 0 && span.rowSpan === 2),
				`the origin carries rowSpan 2; got ${JSON.stringify(s.spans)}`
			)
			assert(s.codes.includes('table/span-out-of-range'), `the clamp is reported; got ${JSON.stringify(s.codes)}`)
		},
	},
	{
		name: 'a colspan past the last column is written as far as the grid goes',
		fn: async () => {
			const s = await structure([['A', 'B'], [{ text: 'D', options: { colspan: 3 } }]])
			assertEqual(s.gridCols, 2, 'two grid columns')
			assertRectangular(s, 'colspan 3')
			assert(
				s.spans.some((span) => span.rIdx === 1 && span.cIdx === 0 && span.gridSpan === 2),
				`the origin carries gridSpan 2; got ${JSON.stringify(s.spans)}`
			)
			assert(s.codes.includes('table/span-out-of-range'), `the clamp is reported; got ${JSON.stringify(s.codes)}`)
		},
	},
	{
		name: 'a row wider than the grid drops its extra cells and says so; a short row is filled',
		fn: async () => {
			const s = await structure([['A', 'B'], ['C', 'D', 'E'], ['F']])
			assertRectangular(s, 'ragged rows')
			assert(s.codes.includes('table/cell-past-grid'), `the dropped cell is reported; got ${JSON.stringify(s.codes)}`)
		},
	},
	{
		name: 'a well-formed merged table reports nothing',
		fn: async () => {
			const s = await structure([[{ text: 'A', options: { colspan: 2, rowspan: 2 } }, 'C'], ['F'], ['G', 'H', 'I']])
			assertRectangular(s, 'well-formed')
			assertEqual(JSON.stringify(s.codes), '[]', 'no diagnostic')
		},
	},
])
