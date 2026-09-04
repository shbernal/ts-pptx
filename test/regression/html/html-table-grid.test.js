import { describe, test } from 'vitest'
import { assert } from '../../helpers.js'
import { extendColBasis, measureGridColumns } from '../../../src/gen/table/html-dom.ts'
import { createRowSpanOccupancy } from '../../../src/gen/table/grid.ts'

// Acceptance: an HTML table's rows do not state their own width — a `colspan` fills several grid
// columns and a `rowspan` from above fills one the row never mentions — but `<a:tblGrid>` declares
// a column count that every `<a:tr>` must match. These are the two DOM-independent decisions that
// translate between the two models, unit-tested directly (the pattern docs/project-target.md
// prescribes for this file's DOM-bound neighbours).

/** Shorthand: a row of plain single-track cells. */
const plain = (count) => Array.from({ length: count }, () => ({}))

describe('measureGridColumns', () => {
	test('a rectangular table is as wide as its rows', () => {
		const { columns, filled } = measureGridColumns([plain(3), plain(3)])
		assert(columns === 3, `expected 3 columns; got ${columns}`)
		assert(JSON.stringify(filled) === '[3,3]', `every row reaches all 3; got ${JSON.stringify(filled)}`)
	})

	test('a colspan fills the columns it covers', () => {
		const { columns, filled } = measureGridColumns([[{ colspan: 2 }, {}], plain(3)])
		assert(columns === 3, `a 2-span plus one cell is 3 columns; got ${columns}`)
		assert(JSON.stringify(filled) === '[3,3]', `the spanning row is not short; got ${JSON.stringify(filled)}`)
	})

	test('a rowspan fills the following row without that row mentioning it', () => {
		// The emitter synthesizes the vMerge continuation cell itself, so the second row is already
		// complete at 2 columns — counting it as short here would pad it into a 3-column row.
		const { columns, filled } = measureGridColumns([[{ rowspan: 2 }, {}], [{}]])
		assert(columns === 2, `expected 2 columns; got ${columns}`)
		assert(JSON.stringify(filled) === '[2,2]', `the rowspan fills the second row; got ${JSON.stringify(filled)}`)
	})

	test('a rowspan deeper than one row keeps holding its column', () => {
		const { filled } = measureGridColumns([[{ rowspan: 3 }, {}], [{}], [{}], plain(2)])
		assert(JSON.stringify(filled) === '[2,2,2,2]', `a 3-deep span covers 2 further rows; got ${JSON.stringify(filled)}`)
	})

	test('a rowspan in the last column still counts as filled', () => {
		// The held column comes *after* the row's own cells, so it is only found by looking past
		// the final cell — the case a "sum the spans" count misses.
		const { filled } = measureGridColumns([[{}, { rowspan: 2 }], [{}]])
		assert(JSON.stringify(filled) === '[2,2]', `trailing held columns count; got ${JSON.stringify(filled)}`)
	})

	test('the grid is as wide as the widest row, and short rows are reported short', () => {
		const { columns, filled } = measureGridColumns([[{ colspan: 3 }], plain(4), [{}]])
		assert(columns === 4, `the widest row sets the grid; got ${columns}`)
		assert(JSON.stringify(filled) === '[3,4,1]', `each row's reach; got ${JSON.stringify(filled)}`)
	})

	test('a table with no cells has no columns', () => {
		const { columns } = measureGridColumns([[], []])
		assert(columns === 0, `empty rows occupy nothing; got ${columns}`)
	})

	test('a nonsense span covers exactly one column', () => {
		// A negative or zero span read literally would walk the column cursor backwards (or not at
		// all) and shift every column after it.
		const { columns, filled } = measureGridColumns([[{ colspan: -2 }, { colspan: 0 }, { rowspan: NaN }]])
		assert(columns === 3, `three bad spans are still three columns; got ${columns}`)
		assert(JSON.stringify(filled) === '[3]', `the row reaches 3; got ${JSON.stringify(filled)}`)
	})

	test('a fractional span truncates rather than fractionally advancing', () => {
		const { columns } = measureGridColumns([[{ colspan: 2.9 }]])
		assert(columns === 2, `2.9 covers 2 whole columns; got ${columns}`)
	})
})

describe('extendColBasis', () => {
	test('a basis already covering the grid is returned unchanged', () => {
		assert(JSON.stringify(extendColBasis([1, 2, 3], 3)) === '[1,2,3]', 'nothing to extend')
	})

	test('missing columns take the average of the stated ones', () => {
		assert(JSON.stringify(extendColBasis([10, 30], 4)) === '[10,30,20,20]', 'extra columns read as ordinary')
	})

	test('an all-equal basis extends to an equal split', () => {
		assert(JSON.stringify(extendColBasis([1, 1], 3)) === '[1,1,1]', 'the equal-split fallback stays equal')
	})

	test('a zero-sum basis extends with 1 rather than 0', () => {
		// A zero would make the extra column a zero-width sliver; 1 makes the proportional calc
		// downstream give it an ordinary share.
		assert(JSON.stringify(extendColBasis([0, 0], 3)) === '[0,0,1]', 'no column may be extended to zero width')
	})

	test('an empty basis extends to all ones', () => {
		assert(JSON.stringify(extendColBasis([], 2)) === '[1,1]', 'with nothing stated, every column is equal')
	})
})

// The rowspan bookkeeping underneath both grid walks: `measureGridColumns` above, which measures
// how wide an imported HTML table's grid *is*, and `walkTableGrid`, which places authored cells
// into a grid whose width is already declared. Neither can be expressed as the other without
// giving one of them a mode it does not want, so what they share is the bookkeeping, not the walk
// — and it is unit-tested here rather than only through its two callers, because getting the
// ageing off by one silently shifts every column below.
describe('createRowSpanOccupancy', () => {
	test('a column is free until something holds it', () => {
		const occ = createRowSpanOccupancy()
		assert(occ.nextFree(0) === 0, 'nothing is held to begin with')
		occ.hold(0, 1, 2)
		assert(occ.nextFree(0) === 1, 'column 0 is held, so the next free one is 1')
	})

	test('a hold covers every column of its colspan', () => {
		const occ = createRowSpanOccupancy()
		occ.hold(0, 3, 2)
		occ.endRow()
		assert(occ.nextFree(0) === 3, 'all three columns stay held into the next row')
	})

	test('a rowspan of 2 holds exactly one further row', () => {
		const occ = createRowSpanOccupancy()
		occ.hold(0, 1, 2)
		occ.endRow()
		assert(occ.nextFree(0) === 1, 'still held one row on')
		occ.endRow()
		assert(occ.nextFree(0) === 0, 'and free the row after that')
	})

	test('holds of different depths expire independently', () => {
		// One shared "rows remaining" counter instead of one per column gets this wrong.
		const occ = createRowSpanOccupancy()
		occ.hold(0, 1, 3)
		occ.hold(1, 1, 2)
		occ.endRow()
		assert(occ.nextFree(0) === 2, 'both columns still held')
		occ.endRow()
		assert(occ.nextFree(0) === 1, 'the 2-deep hold has expired and the 3-deep one has not')
	})

	test('a rowspan of 1 holds nothing beyond its own row', () => {
		const occ = createRowSpanOccupancy()
		occ.hold(0, 1, 1)
		occ.endRow()
		assert(occ.nextFree(0) === 0, 'an unspanned cell leaves nothing behind')
	})

	test('nextFree walks past a run of held columns, not just one', () => {
		const occ = createRowSpanOccupancy()
		occ.hold(0, 1, 2)
		occ.hold(1, 1, 2)
		occ.hold(2, 1, 2)
		occ.endRow()
		assert(occ.nextFree(0) === 3, 'three consecutive holds are skipped in one call')
	})

	test('a column past the end of the record is free', () => {
		// The record is sparse and unbounded on purpose: a caller that knows its width clamps what
		// it holds, and then nothing past that width is ever held.
		const occ = createRowSpanOccupancy()
		occ.hold(0, 1, 2)
		assert(occ.nextFree(9) === 9, 'a column nothing ever touched is free')
	})
})
