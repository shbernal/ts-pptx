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
