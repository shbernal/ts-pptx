// The merge encoding PowerPoint writes, on table-merge-encoding.pptx, as the target for the read
// path's structural edits. Each slide holds one table:
//
//   slide 1  merge-2x2       3x3, (1,1)-(2,2) merged
//   slide 2  merge-inserted  slide 1's merge after Rows.Add(2) and Columns.Add(2) through it
//   slide 3  merge-deleted   a 4x4 with (1,1)-(3,3) merged, after Rows(2) and Columns(2) deleted
//   slide 4  merge-1x2       3x3, (1,1)-(1,2) merged
//   slide 5  merge-2x1       3x3, (1,1)-(2,1) merged
//
// PowerPoint repeats a region's spans on its covered cells: the ones in its first row carry its
// rowSpan, the ones in its first column its gridSpan, and the inner ones only the two flags. Every
// case compares the span attributes of the whole grid, so an edit that gets one cell's form wrong
// fails even when the table is still consistent.

import { describe, test } from 'vitest'

import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const SPAN_ATTRIBUTES = ['rowSpan', 'gridSpan', 'hMerge', 'vMerge']

/** The table on slide `index` of the fixture, read fresh. */
async function tableOn(index) {
	for (const shape of (await openFixture('table-merge-encoding')).slides[index].shapes) {
		if (shape.shapeType === 'graphicFrame' && shape.table) return shape.table
	}
	throw new Error(`slide ${index + 1} holds no table`)
}

/** Every cell's span attributes, row by row, as `name=value` in a fixed order. */
function spanGrid(table) {
	return table.rows.map((row) =>
		row.cells.map((cell) =>
			SPAN_ATTRIBUTES.filter((name) => cell.element_.getAttribute(name))
				.map((name) => `${name}=${cell.element_.getAttribute(name)}`)
				.join(' ')
		)
	)
}

describe('Table.mergeCells writes the spans PowerPoint writes (table-merge-encoding.pptx)', () => {
	test.for([
		{ slide: 0, name: 'a 2x2 merge', to: [1, 1] },
		{ slide: 3, name: 'a 1x2 merge', to: [0, 1] },
		{ slide: 4, name: 'a 2x1 merge', to: [1, 0] },
	])('$name', async ({ slide, to }) => {
		const table = await tableOn(slide)
		const powerPoint = spanGrid(table)
		table.unmergeCell(0, 0)
		assert(
			spanGrid(table).every((row) => row.every((cell) => cell === '')),
			'unmerging leaves no span attribute to compare against'
		)
		table.mergeCells(0, 0, to[0], to[1])
		assertEqual(JSON.stringify(spanGrid(table)), JSON.stringify(powerPoint))
	})
})

describe("structural edits through a PowerPoint merge match PowerPoint's own (table-merge-encoding.pptx)", () => {
	test('inserting a row and a column through the merge', async () => {
		const table = await tableOn(0)
		table.addRow(1)
		table.addColumn(1)
		assertEqual(JSON.stringify(spanGrid(table)), JSON.stringify(spanGrid(await tableOn(1))))
	})

	test('deleting a row and a column through the merge', async () => {
		const table = await tableOn(1)
		table.removeRow(1)
		table.removeColumn(1)
		assertEqual(JSON.stringify(spanGrid(table)), JSON.stringify(spanGrid(await tableOn(2))))
	})
})

describe('Table.mergeCells keeps the text of the cells it covers', () => {
	test("the origin holds each covered cell's paragraphs, row-major, as PowerPoint writes them", async () => {
		// Slide 1's merged origin holds `1,1`, `1,2`, `2,1` and `2,2` as four paragraphs: that is
		// PowerPoint's own Merge Cells, and it is what this reproduces. Merging used to empty the
		// covered cells outright, so the three texts other than the origin's were destroyed.
		const powerPoint = await tableOn(0)
		const expected = powerPoint.cell(0, 0).text
		assertEqual(expected, '1,1\n1,2\n2,1\n2,2', 'the fixture is what this test thinks it is')

		const table = await tableOn(0)
		table.unmergeCell(0, 0)
		// Unmerging does not put the text back -- it never left the origin -- so the covered cells
		// are re-filled to reconstruct the pre-merge table before merging again.
		assertEqual(table.cell(0, 0).text, expected, 'unmerging leaves the origin holding all four')
		table.cell(0, 0).text = '1,1'
		table.cell(0, 1).text = '1,2'
		table.cell(1, 0).text = '2,1'
		table.cell(1, 1).text = '2,2'

		table.mergeCells(0, 0, 1, 1)
		assertEqual(table.cell(0, 0).text, expected, 'merging gathers them back in the same order')
		for (const [r, c] of [
			[0, 1],
			[1, 0],
			[1, 1],
		]) {
			assertEqual(table.cell(r, c).text, '', `the covered cell (${r},${c}) is emptied`)
		}
		assertEqual(table.cell(2, 0).text, '3,1', 'a cell outside the rectangle is untouched')
	})

	test('an empty covered cell adds no paragraph, and formatting comes with the text', async () => {
		// Both from a COM probe of PowerPoint's Merge Cells, which the committed fixture cannot
		// show: its four merged cells all carry plain text. A cell with nothing in it contributes
		// nothing rather than a blank line, and a bold, coloured run keeps both.
		const table = await tableOn(0)
		table.unmergeCell(0, 0)
		table.cell(0, 0).text = 'origin'
		table.cell(0, 1).text = ''
		table.cell(1, 0).text = 'tail'
		table.cell(1, 1).text = ''

		table.mergeCells(0, 0, 1, 1)
		assertEqual(table.cell(0, 0).text, 'origin\ntail', 'the two empty cells add nothing')
		const runs = table.cell(0, 0).textFrame.paragraphs.flatMap((paragraph) => paragraph.runs)
		assertEqual(runs.length, 2, 'two runs, one per non-empty cell')
	})
})
