/**
 * The auto-pager never ends a page on a row whose rowspan reaches past it.
 *
 * The break guard protected only rows covered from ABOVE. When a page filled partway through the
 * row that starts a rowspan, that partial row became the page's last row with its `rowspan` still
 * set, so the emitter wrote `rowSpan="2"` with no row under it -- the state PowerPoint reports as
 * corrupt -- and the next page opened with the rest of the merged row. Twenty-two plain rows, then
 * a 40-word cell spanning two rows in 1in columns, put that `rowSpan` on the last `<a:tr>` of slide 2.
 *
 * Checked through the read model, which reports each cell's `rowSpan` against the rows its own
 * table actually has, so the assertion is about the emitted structure and not about the pager.
 */
import TsPptx from '../../../dist/node.js'
import { Presentation } from '../../../dist/read.js'
import { assert, captureDiagnostics, defineRegressionSuite } from '../../helpers.js'

const WORDS = Array.from({ length: 40 }, (_unused, i) => `word${i}`).join(' ')

/** Every page's table, read back, with each origin cell's rowSpan checked against the page. */
async function spanProblems(rows) {
	const pres = new TsPptx()
	pres.addSlide().addTable(rows, { x: 0.5, y: 0.5, colW: [1, 1], autoPage: true })
	const presentation = await Presentation.load(await pres.toBytes())
	const problems = []
	let pages = 0
	presentation.slides.forEach((slide, slideIdx) => {
		for (const shape of slide.shapes) {
			// Only a graphic frame carries `table`; the cast is for the other members of the union.
			const table = /** @type {any} */ (shape).table
			if (!table) continue
			pages++
			const tableRows = table.rows
			tableRows.forEach((row, r) => {
				row.cells.forEach((cell, c) => {
					if (cell.isMergeContinuation) return
					if (r + cell.rowSpan > tableRows.length)
						problems.push(
							`slide ${slideIdx + 1} (${r},${c}) spans ${cell.rowSpan} rows of ${tableRows.length - r} left`
						)
				})
			})
		}
	})
	return { problems, pages }
}

defineRegressionSuite('Auto-page break before a rowspan', [
	{
		name: 'a page never ends on a row whose rowspan reaches past it',
		fn: async () => {
			const plain = (n) => Array.from({ length: n }, (_unused, i) => [`p${i}`, 'v'])
			const rows = [...plain(22), [{ text: WORDS, options: { rowspan: 2 } }, 'X'], ['Y'], ...plain(5)]
			const { problems, pages } = await spanProblems(rows)
			assert(pages >= 2, `the table has to page to test a break; got ${pages} page(s)`)
			assert(problems.length === 0, `\n  ${problems.join('\n  ')}`)
		},
	},
	{
		name: 'a rowspan group taller than a page is kept whole and reported',
		fn: async () => {
			const tall = Array.from({ length: 12 }, () => WORDS).join(' ')
			const rows = [['a', 'b'], [{ text: tall, options: { rowspan: 3 } }, 'X'], ['Y'], ['Z'], ['c', 'd']]
			const { result, codes } = await captureDiagnostics(() => spanProblems(rows))
			assert(result.problems.length === 0, `\n  ${result.problems.join('\n  ')}`)
			assert(codes.includes('table/autopage-rowspan-too-tall'), `reported; got ${JSON.stringify(codes)}`)
		},
	},
])
