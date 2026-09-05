/**
 * A word wider than its column costs no extra line.
 *
 * The auto-pager wraps by pushing the buffered line and starting a new one whenever the next
 * word would not fit. On the FIRST word of a cell the buffer is empty, so a word wider than the
 * column on its own pushed an empty line before it -- the flush at the end of the same loop
 * guards for exactly that (`lineCells.length > 0`) and the wrap did not.
 *
 * The cell's text survived, because the pager concatenates its lines back together. Its height
 * did not: `_lines.length` is what prices the row, so every such cell was budgeted one line too
 * tall and the table paged early. Twelve identical single-word rows in a 0.4in column took two
 * slides when the word was long and one when it was short.
 *
 * The word still overflows its column after this. Nothing here breaks inside a word, and
 * whether it should is a separate question this does not answer.
 */
import { assert, assertEqual, build, defineRegressionSuite, listEntries, readEntry } from '../../helpers.js'

const POS = { x: 0.5, y: 0.5, colW: [0.4], autoPage: true }

/** Twelve single-word rows, all the same word. */
async function paginate(word) {
	const rows = Array.from({ length: 12 }, () => [{ text: word }])
	const { zip } = await build((p) => {
		p.addSlide().addTable(rows, POS)
	})
	return zip
}

const slideCount = (zip) => listEntries(zip).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length

defineRegressionSuite('Auto-page: a word wider than its column', [
	{
		name: 'costs no more lines than a word that fits',
		fn: async () => {
			const fits = slideCount(await paginate('ab'))
			const overWide = slideCount(await paginate('Supercalifragilistic'))
			assertEqual(overWide, fits, 'an over-wide word must not buy a blank line and page the table early')
		},
	},
	{
		name: 'still reaches the slide whole',
		fn: async () => {
			// The guard must skip the empty push, not the word: dropping the buffer would lose text.
			const zip = await paginate('Supercalifragilistic')
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(xml.includes('Supercalifragilistic'), `the over-wide word must survive; got: ${xml}`)
		},
	},
])
