import { defineRegressionSuite, build, captureDiagnostics, readEntry, assert, assertEqual } from '../../helpers.js'

// `colspan`/`rowspan` arrive from the calling program, and the merge-grid builder used to trust
// them: `new Array(colspan - 1).fill(undefined)` at `colspan: 4294967295` is not a slow path but
// a process kill (V8 aborts on the allocation -- `FATAL ERROR: … JavaScript heap out of memory`,
// with no exception to catch), and a negative or fractional span shifts every column after it
// while emitting a `gridSpan` PowerPoint cannot make sense of. The project's line is to warn on
// out-of-range input rather than emit a degenerate result, so each of those now falls back to a
// span of 1 with a `table/span-out-of-range` diagnostic.
//
// The two things worth asserting together are that the deck still *builds* (the allocation is
// bounded) and that the grid is internally consistent afterwards: the column count, the hMerge
// fillers and the emitted `gridSpan` attribute all read the corrected number, because they are
// corrected once, before the grid is built, rather than at each of the five sites that read a
// span.

async function tableXml(zip) {
	const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
	return /<a:tbl>[\s\S]*<\/a:tbl>/.exec(xml)[0]
}

/** Build a two-row table whose first cell carries `options`, returning the `<a:tbl>` and the codes. */
async function tableWithCellOptions(options) {
	const { result, codes } = await captureDiagnostics(() =>
		build((p) => {
			p.addSlide().addTable(
				[
					[{ text: 'A', options }, { text: 'B' }],
					[{ text: 'c' }, { text: 'd' }],
				],
				{ x: 0, y: 0, w: 6 }
			)
		})
	)
	return { tbl: await tableXml(result.zip), codes }
}

const OUT_OF_RANGE = 'table/span-out-of-range'

defineRegressionSuite('Table colspan/rowspan range checking', [
	{
		name: 'an unbounded colspan is refused instead of allocating for it',
		fn: async () => {
			// The reported shape: the largest `xsd:unsignedInt`. Reaching the assertions below at
			// all is most of the point — before the check this call took the process with it.
			const { tbl, codes } = await tableWithCellOptions({ colspan: 4294967295 })
			assert(codes.includes(OUT_OF_RANGE), `expected a ${OUT_OF_RANGE} diagnostic, got: ${codes.join(', ')}`)
			assertEqual((tbl.match(/<a:gridCol/g) || []).length, 2, 'the grid is the two columns actually authored')
			assertEqual((tbl.match(/gridSpan=/g) || []).length, 0, 'no gridSpan is emitted for a span of 1')
			assertEqual((tbl.match(/hMerge="1"/g) || []).length, 0, 'and no filler cells were synthesized')
		},
	},
	{
		name: 'a span past the ceiling is refused even though it would not exhaust memory',
		fn: async () => {
			// 1001 is one past `MAX_TABLE_SPAN`. It would allocate fine; it is still not a table.
			const { tbl, codes } = await tableWithCellOptions({ colspan: 1001 })
			assert(codes.includes(OUT_OF_RANGE), 'a span past the ceiling warns')
			assertEqual((tbl.match(/<a:gridCol/g) || []).length, 2, 'the grid is unchanged')
		},
	},
	{
		name: 'a negative, a fraction and a NaN each fall back to 1',
		fn: async () => {
			for (const colspan of [-2, 1.5, NaN, 0]) {
				const { tbl, codes } = await tableWithCellOptions({ colspan })
				assert(codes.includes(OUT_OF_RANGE), `colspan ${String(colspan)} warns`)
				assertEqual((tbl.match(/<a:gridCol/g) || []).length, 2, `colspan ${String(colspan)} leaves a 2-column grid`)
				assertEqual((tbl.match(/hMerge="1"/g) || []).length, 0, `colspan ${String(colspan)} synthesizes no fillers`)
			}
		},
	},
	{
		name: 'rowspan is checked the same way, and named as rowspan in the diagnostic',
		fn: async () => {
			const { result, messages, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[{ text: 'A', options: { rowspan: -3 } }, { text: 'B' }],
							[{ text: 'c' }, { text: 'd' }],
						],
						{ x: 0, y: 0, w: 6 }
					)
				})
			)
			assert(codes.includes(OUT_OF_RANGE), 'an out-of-range rowspan warns')
			assert(
				messages.some((m) => m.includes('rowspan')),
				`the message names the option that was wrong: ${messages.join(' | ')}`
			)
			const tbl = await tableXml(result.zip)
			assertEqual((tbl.match(/vMerge="1"/g) || []).length, 0, 'no rowspan continuation cells were synthesized')
		},
	},
	{
		name: 'one bad cell warns once, not once per site that reads its span',
		fn: async () => {
			const { codes } = await tableWithCellOptions({ colspan: -1 })
			const spanCodes = codes.filter((c) => c === OUT_OF_RANGE)
			assertEqual(spanCodes.length, 1, `expected exactly one diagnostic, got ${spanCodes.length}`)
		},
	},
	{
		// The second allocation, on the other path. The auto-pager sizes a per-column depth array
		// from a column count that is a sum of colspans, so `autoPage: true` reached the same abort
		// without ever getting as far as the merge grid. Both are fed by `addTableDefinition`, which
		// is why the check lives there rather than in either of them.
		name: 'the auto-paged path is guarded too, not just the merge grid',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[{ text: 'A', options: { colspan: 4294967295 } }, { text: 'B' }],
							[{ text: 'c' }, { text: 'd' }],
						],
						{ x: 0, y: 0, w: 6, autoPage: true }
					)
				})
			)
			assert(codes.includes(OUT_OF_RANGE), `expected a ${OUT_OF_RANGE} diagnostic, got: ${codes.join(', ')}`)
			const tbl = await tableXml(result.zip)
			assertEqual((tbl.match(/<a:gridCol/g) || []).length, 2, 'the paged table has the two columns authored')
		},
	},
	{
		name: 'a valid span is untouched and warns about nothing',
		fn: async () => {
			const { tbl, codes } = await tableWithCellOptions({ colspan: 2 })
			assertEqual(codes.filter((c) => c === OUT_OF_RANGE).length, 0, 'a legal span is not a finding')
			assert(/gridSpan="2"/.test(tbl), 'the colspan origin still declares gridSpan="2"')
			assertEqual((tbl.match(/hMerge="1"/g) || []).length, 1, 'and its filler cell is still synthesized')
		},
	},
])
