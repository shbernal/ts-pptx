/**
 * `addTable` treats its `rows` argument as a read-only input.
 *
 * This is a guard, not a fix: the property already holds, and it holds for one reason --
 * `gen/define/table.ts` copies every cell and every option bag (`{ ...cell, options: {
 * ...cell.options } }`) before the auto-pager sees them. The pager then writes to its own
 * working copy freely, which is only safe while that copy exists. Nothing stated that
 * dependency, so nothing would have caught its removal; the chart path has had the same guard
 * for its own normalization (`test/regression/chart/chart-input-immutability.test.js`) and the
 * table path had none.
 *
 * Freezing rather than comparing afterwards is deliberate: a write to a frozen object throws in
 * a module, so this fails at the site of the mutation rather than at an equality check that
 * cannot say which key moved.
 */
import { expect } from 'vitest'
import { assert, assertEqual, build, defineRegressionSuite, readEntry } from '../../helpers.js'

const POS = { x: 0.5, y: 0.5, w: 9 }

/** Freeze a rows array and every cell and option bag inside it. */
function deepFreezeRows(rows) {
	for (const row of rows)
		for (const cell of row) {
			if (cell.options) Object.freeze(cell.options)
			Object.freeze(cell)
		}
	rows.forEach(Object.freeze)
	return Object.freeze(rows)
}

/** Twelve rows of one long-ish cell, enough to page under a short height. */
function longRows() {
	return Array.from({ length: 12 }, (_, i) => [
		{ text: `row ${i} with several words in it`, options: { fontSize: 14 } },
	])
}

defineRegressionSuite('Table input immutability', [
	{
		name: 'a frozen row set survives addTable with autoPage on and a table-level charWeight',
		fn: async () => {
			const rows = deepFreezeRows(longRows())
			const { zip } = await build((p) => {
				p.addSlide().addTable(rows, { ...POS, h: 2, autoPage: true, autoPageCharWeight: 0.4 })
			})
			// Frozen input is only proof if the table actually reached the part.
			assert((await readEntry(zip, 'ppt/slides/slide1.xml')).includes('<a:tbl>'), 'the table must reach the slide')
		},
	},
	{
		name: 'a frozen row set survives addTable when the table states no charWeight',
		fn: async () => {
			// Deleting a key from a frozen bag throws exactly as assigning one does, so the two
			// cases are separate: an option the table states and an option it does not.
			const rows = deepFreezeRows(longRows())
			await build((p) => {
				p.addSlide().addTable(rows, { ...POS, h: 2, autoPage: true })
			})
		},
	},
	{
		name: 'the caller`s cell options come back exactly as they went in',
		fn: async () => {
			const rows = [[{ text: 'a b c d e f g h', options: { autoPageCharWeight: 0.2, fontSize: 14 } }]]
			const before = structuredClone(rows)
			await build((p) => {
				p.addSlide().addTable(rows, { ...POS, autoPage: true, autoPageCharWeight: 0.6 })
			})
			expect(rows).toEqual(before)
			assertEqual(rows[0][0].options.autoPageCharWeight, 0.2, 'the cell`s own weight must survive by value')
		},
	},
])
