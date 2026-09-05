/**
 * `autoPageCharWeight` is readable on a cell as well as on the table, and the cell's wins.
 *
 * The auto-pager measures a cell twice through two call sites of `parseTextToLines`, and they
 * had resolved this option differently. The row-height probe read `cell.options` directly; the
 * main loop stamped the table's value over the cell's first -- and `delete`d the cell's when the
 * table stated none -- so in the loop that actually decides where a page breaks, a weight set on
 * a cell alone did nothing at all. `TableCellProps` declares the option, so that was an option
 * accepted and dropped.
 *
 * Both sites now resolve it the same way, with the more specific statement winning, which is the
 * precedence every other paired table option uses (`fontSize` through `resolveCellFontSize`).
 * A stated `0` is part of that: it is the caller asking for the default weight on this cell, not
 * the caller saying nothing and inheriting the table's.
 *
 * Every case reads the page count rather than the wrapped lines, because that is the thing the
 * weight exists to move and the thing a caller sees.
 */
import { assert, assertEqual, build, defineRegressionSuite, listEntries } from '../../helpers.js'

const POS = { x: 0.5, y: 0.5, w: 3, h: 2 }
const TEXT = 'alpha beta gamma delta epsilon zeta eta theta'

/**
 * How many slides a 14-row table pages onto under the given weights.
 * @param cellWeight - `autoPageCharWeight` on every cell, or `undefined` for none
 * @param tableWeight - `autoPageCharWeight` on the table, or `undefined` for none
 */
async function pages(cellWeight, tableWeight) {
	const rows = Array.from({ length: 14 }, () => [
		{ text: TEXT, options: cellWeight === undefined ? {} : { autoPageCharWeight: cellWeight } },
	])
	const { zip } = await build((p) => {
		p.addSlide().addTable(
			rows,
			tableWeight === undefined
				? { ...POS, autoPage: true }
				: { ...POS, autoPage: true, autoPageCharWeight: tableWeight }
		)
	})
	return listEntries(zip).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).length
}

defineRegressionSuite('Auto-page character weight', [
	{
		name: 'a weight on the table alone widens the lines and shortens the table',
		fn: async () => {
			const none = await pages(undefined, undefined)
			const wide = await pages(undefined, 0.9)
			assert(wide < none, `a table charWeight of 0.9 must page shorter; got ${wide} against ${none}`)
		},
	},
	{
		name: 'a weight on a cell alone does the same -- it used to do nothing',
		fn: async () => {
			const none = await pages(undefined, undefined)
			const wide = await pages(0.9, undefined)
			assert(wide < none, `a cell charWeight of 0.9 must page shorter; got ${wide} against ${none}`)
		},
	},
	{
		name: 'a cell that states a weight wins over a table that states another, either direction',
		fn: async () => {
			assertEqual(await pages(0.9, -0.9), await pages(0.9, undefined), 'a wide cell under a narrow table')
			assertEqual(await pages(-0.9, 0.9), await pages(-0.9, undefined), 'a narrow cell under a wide table')
		},
	},
	{
		name: 'a cell weight of 0 is a stated weight, not silence to inherit through',
		fn: async () => {
			assertEqual(await pages(0, 0.9), await pages(undefined, undefined), 'a stated 0 must not take the table`s 0.9')
		},
	},
])
