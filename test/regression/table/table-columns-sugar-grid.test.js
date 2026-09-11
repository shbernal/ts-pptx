/**
 * `columns[i]` styles the cells that start in grid column `i`, wherever they sit in their row.
 *
 * The sugar kept its own column cursor, `colCursor += colspan || 1`. It counted a colspan but not a
 * rowspan from an earlier row, so under `[[{ text: 'A', rowspan: 2 }, 'B'], ['C']]` it gave `C` the
 * first column's definition though `C` sits in the second. It also ran before the spans were
 * checked, so a string or negative colspan moved the cursor. The sugar now reads the placements the
 * emitter, the auto-pager and the measured fit read.
 */
import { assertEqual, captureDiagnostics, defineRegressionSuite, build, readEntry } from '../../helpers.js'

const RED = { fill: { color: 'FF0000' } }
const BLUE = { fill: { color: '0000FF' } }
const GREEN = { fill: { color: '00FF00' } }

/** Each written cell as `text:fill`, row by row, with merge cells as `(merge)`. */
async function cellFills(rows, columns) {
	const { result, codes } = await captureDiagnostics(() =>
		build((p) => p.addSlide().addTable(rows, { x: 0.5, y: 0.5, w: 6, columns }))
	)
	const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
	const cells = [...xml.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map((row) =>
		[...row[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((tc) => {
			const text = [...tc[0].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('')
			const fill = /<a:srgbClr val="([0-9A-F]{6})"\/><\/a:solidFill>\s*<\/a:tcPr>/.exec(tc[0])?.[1] ?? '-'
			return `${text || '(merge)'}:${fill}`
		})
	)
	return { cells, codes }
}

defineRegressionSuite('Table columns sugar follows the grid', [
	{
		name: 'a cell below a rowspan takes the definition of the column it sits in',
		fn: async () => {
			const { cells } = await cellFills([[{ text: 'A', options: { rowspan: 2 } }, 'B'], ['C']], [RED, BLUE])
			assertEqual(
				JSON.stringify(cells),
				JSON.stringify([
					['A:FF0000', 'B:0000FF'],
					['(merge):FF0000', 'C:0000FF'],
				]),
				'C sits in the second column and takes its fill'
			)
		},
	},
	{
		name: 'a colspan moves the cells after it, as it always did',
		fn: async () => {
			const { cells } = await cellFills(
				[
					[{ text: 'A', options: { colspan: 2 } }, 'C'],
					['D', 'E', 'F'],
				],
				[RED, BLUE, GREEN]
			)
			assertEqual(
				JSON.stringify(cells),
				JSON.stringify([
					['A:FF0000', '(merge):FF0000', 'C:00FF00'],
					['D:FF0000', 'E:0000FF', 'F:00FF00'],
				]),
				'the merged cell is the first column, C the third'
			)
		},
	},
	{
		name: 'a colspan that is not a span moves nothing, and is reported once',
		fn: async () => {
			const { cells, codes } = await cellFills([[{ text: 'A', options: { colspan: -2 } }, 'B']], [RED, BLUE])
			assertEqual(JSON.stringify(cells), JSON.stringify([['A:FF0000', 'B:0000FF']]), 'A and B keep their columns')
			assertEqual(
				codes.filter((code) => code === 'table/span-out-of-range').length,
				1,
				`one report for the one bad span; got ${JSON.stringify(codes)}`
			)
		},
	},
])
