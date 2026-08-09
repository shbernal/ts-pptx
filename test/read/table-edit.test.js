// Editing a table in an existing deck: cell property setters and structural edits.
//
// One invariant dominates every case here, and it is the reason this file asserts on
// serialized XML rather than on the read model it just mutated. `CT_TableCellProperties` is
// a SEQUENCE, so an append-only setter produces an out-of-order `a:tcPr` — which PowerPoint
// reports as a corrupt file rather than as a bad edit. A getter would happily read back a
// value from an element in the wrong place, so round-tripping through `save()` and looking
// at the bytes is what actually proves the edit is legal.
//
// The structural half has the analogous invariant: `a:tblGrid/a:gridCol` count must equal
// every row's `a:tc` count, and a `@gridSpan="n"` origin must be followed by exactly n-1
// `@hMerge` cells (likewise `@rowSpan` / `@vMerge` down a column). Each case re-checks the
// whole rectangle rather than only the cells it touched, because the failures worth catching
// are the ones an edit causes somewhere it was not looking.

import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import JSZip from 'jszip'
import { firstTable } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** Author a deck, load it for editing, and return the presentation plus its first table. */
async function editable(build) {
	const pres = new TsPptx()
	build(pres)
	const buf = /** @type {Uint8Array} */ (await pres.stream())
	const presentation = await Presentation.load(buf)
	const table = firstTable(presentation)
	assert(table, 'the authored table is found')
	return { presentation, table }
}

/** Reload a presentation from its own saved bytes, and return it with its first table. */
async function reload(presentation) {
	const reloaded = await Presentation.load(await presentation.save())
	return { presentation: reloaded, table: firstTable(reloaded) }
}

/**
 * Run `fn` and return the `code` of the error it threw.
 *
 * Asserting on the code rather than the message: a code is API and a message explicitly is
 * not, so a message assertion breaks on any wording improvement. Returns `null` when nothing
 * threw, which fails the comparison with a legible diff rather than a bare `false`.
 */
function codeOfThrow(fn) {
	try {
		fn()
	} catch (err) {
		return err.code ?? `(threw without a code: ${err.message})`
	}
	return null
}

/** Save the edited deck and return its slide part. */
async function savedSlide(presentation) {
	const zip = await JSZip.loadAsync(await presentation.save())
	return zip.file('ppt/slides/slide1.xml').async('string')
}

/** A plain 3x3 table with no styling to get in the way. */
function plainTable(p) {
	p.addSlide().addTable(
		[
			[{ text: 'A1' }, { text: 'B1' }, { text: 'C1' }],
			[{ text: 'A2' }, { text: 'B2' }, { text: 'C2' }],
			[{ text: 'A3' }, { text: 'B3' }, { text: 'C3' }],
		],
		{ x: 1, y: 1, w: 9, colW: [3, 3, 3] }
	)
}

/** `CT_TableCellProperties`' child sequence — what an edited `a:tcPr` must still match. */
const TCPR_SEQUENCE = [
	'lnL',
	'lnR',
	'lnT',
	'lnB',
	'lnTlToBr',
	'lnBlToTr',
	'cell3D',
	'noFill',
	'solidFill',
	'gradFill',
	'blipFill',
	'pattFill',
	'grpFill',
	'headers',
	'extLst',
]

/** The direct element children of the first `a:tcPr`, in document order. */
function tcPrChildren(xml, index = 0) {
	const blocks = xml.match(/<a:tcPr(?:\/>|[^>]*>[\s\S]*?<\/a:tcPr>)/g) || []
	const block = blocks[index]
	assert(block, `expected an a:tcPr at index ${index}; got: ` + xml)
	let flat = block
	for (const name of TCPR_SEQUENCE) {
		flat = flat.replace(new RegExp(`<a:${name}\\b[^>]*>[\\s\\S]*?</a:${name}>`, 'g'), `<a:${name}/>`)
	}
	return [...flat.matchAll(/<a:(\w+)[^>]*?\/>/g)].map((m) => m[1]).filter((name) => TCPR_SEQUENCE.includes(name))
}

/**
 * Each `a:tcPr` with its border elements stripped, so only the cell's own fill remains.
 *
 * Every border wraps its own `a:noFill` or `a:solidFill` for its stroke, so asking "does
 * this cell have a noFill" of the raw block answers about the borders, not the cell.
 */
function tcPrFillOnly(xml) {
	return (xml.match(/<a:tcPr(?:\/>|[^>]*>[\s\S]*?<\/a:tcPr>)/g) || []).map((block) =>
		block.replace(/<a:ln(?:L|R|T|B|TlToBr|BlToTr)\b[\s\S]*?<\/a:ln(?:L|R|T|B|TlToBr|BlToTr)>/g, '')
	)
}

/** Assert `a:tcPr`'s children are a subsequence of the schema order. */
function assertTcPrOrder(xml, index = 0) {
	const children = tcPrChildren(xml, index)
	const positions = children.map((name) => TCPR_SEQUENCE.indexOf(name))
	for (let i = 1; i < positions.length; i++) {
		assert(
			positions[i] > positions[i - 1],
			`a:tcPr children are out of CT_TableCellProperties order: ${children.join(',')}`
		)
	}
	return children
}

/**
 * Assert the whole grid is still rectangular and every span's continuations are present.
 * This is the check that catches a structural edit breaking something it did not touch.
 */
function assertGridConsistent(table) {
	const colCount = table.columnCount
	const grid = table.rows.map((row) => row.cells)
	for (const [r, cells] of grid.entries()) {
		assertEqual(cells.length, colCount, `row ${r} has one a:tc per grid column`)
	}
	for (const [r, cells] of grid.entries()) {
		for (const [c, cell] of cells.entries()) {
			if (cell.isMergeContinuation) continue
			for (let cc = c + 1; cc < c + cell.gridSpan; cc++) {
				const covered = grid[r][cc]
				assert(covered, `(${r},${cc}) exists to cover the gridSpan at (${r},${c})`)
				assert(covered.isMergeContinuation, `(${r},${cc}) is flagged as covered by (${r},${c})`)
			}
			for (let rr = r + 1; rr < r + cell.rowSpan; rr++) {
				const covered = grid[rr]?.[c]
				assert(covered, `(${rr},${c}) exists to cover the rowSpan at (${r},${c})`)
				assert(covered.isMergeContinuation, `(${rr},${c}) is flagged as covered by (${r},${c})`)
			}
		}
	}
}

describe('TableCell setters — a:tcPr attributes', () => {
	test('anchor, vert, horzOverflow and anchorCtr all set, and null clears', async () => {
		const { presentation, table } = await editable(plainTable)
		const cell = table.cell(0, 0)
		cell.setAnchor('ctr')
		cell.setVerticalText('vert270')
		cell.setHorzOverflow('overflow')
		cell.setAnchorCtr(true)

		let xml = await savedSlide(presentation)
		const tag = xml.match(/<a:tcPr[^>]*>/)[0]
		for (const expected of ['anchor="ctr"', 'vert="vert270"', 'horzOverflow="overflow"', 'anchorCtr="1"']) {
			assert(tag.includes(expected), `expected ${expected}; got: ` + tag)
		}

		// Re-read the saved bytes rather than trusting the live model, so the assertion is
		// about what a consumer would actually open.
		const { presentation: reloaded, table: reloadedTable } = await reload(presentation)
		const back = reloadedTable.cell(0, 0)
		assertEqual(back.anchor, 'ctr', 'anchor reads back')
		assertEqual(back.verticalText, 'vert270', 'vert reads back')
		assertEqual(back.anchorCtr, true, 'anchorCtr reads back')

		back.setAnchor(null)
		back.setVerticalText(null)
		back.setAnchorCtr(false)
		xml = await savedSlide(reloaded)
		const cleared = xml.match(/<a:tcPr[^>]*>/)[0]
		for (const gone of ['anchor=', 'vert=', 'anchorCtr=']) {
			assert(!cleared.includes(gone), `expected ${gone} removed; got: ` + cleared)
		}
	})

	test('a value outside its schema enum throws rather than being written', async () => {
		const { table } = await editable(plainTable)
		const cell = table.cell(0, 0)
		// Throwing, not warn-and-drop: a caller editing one attribute would otherwise be left
		// looking at an unchanged deck with nothing to explain it.
		const cases = [
			{ value: 'middle', code: 'table/invalid-cell-anchor', apply: () => cell.setAnchor('middle') },
			{ value: 'sideways', code: 'table/invalid-cell-vert', apply: () => cell.setVerticalText('sideways') },
			{ value: 'ellipsis', code: 'table/invalid-cell-overflow', apply: () => cell.setHorzOverflow('ellipsis') },
		]
		for (const { value, code, apply } of cases) {
			assertEqual(codeOfThrow(apply), code, `the error names its condition for ${value}`)
		}
		assertEqual(cell.anchor, null, 'nothing was written')
	})

	test('margins set per side, and a null side returns to the schema default', async () => {
		const { presentation, table } = await editable(plainTable)
		table.cell(0, 0).setMarginsEmu({ left: 0, top: 12700 })

		const xml = await savedSlide(presentation)
		let tag = xml.match(/<a:tcPr[^>]*>/)[0]
		assert(tag.includes('marL="0"'), 'the left inset is flush; got: ' + tag)
		assert(tag.includes('marT="12700"'), 'the top inset is set; got: ' + tag)
		// Untouched sides keep the writer's own values rather than being reset.
		assert(tag.includes('marR="91440"'), 'the right inset is untouched; got: ' + tag)

		const { presentation: reloaded, table: reloadedTable } = await reload(presentation)
		reloadedTable.cell(0, 0).setMarginsEmu({ left: null })
		tag = (await savedSlide(reloaded)).match(/<a:tcPr[^>]*>/)[0]
		assert(!tag.includes('marL='), 'a null side removes the attribute; got: ' + tag)
	})

	test('a non-finite margin throws instead of writing NaN', async () => {
		const { table } = await editable(plainTable)
		assertEqual(
			codeOfThrow(() => table.cell(0, 0).setMarginsEmu({ left: Number.NaN })),
			'table/invalid-cell-margin',
			'NaN is rejected rather than reaching the attribute'
		)
	})
})

describe('TableCell setters — fill and borders keep a:tcPr in schema order', () => {
	test('a fill set after the borders still lands after them in the XML', async () => {
		// The bug this defends against: appending `a:solidFill` to a `a:tcPr` that already has
		// borders happens to be correct, but appending a BORDER to one that already has a fill
		// is not — and a setter that appends cannot tell the difference.
		const { presentation, table } = await editable(plainTable)
		const cell = table.cell(0, 0)
		cell.setFillColor('#FF0000')
		cell.setBorder('top', { widthPt: 2, color: '00FF00' })

		const xml = await savedSlide(presentation)
		const children = assertTcPrOrder(xml)
		assert(children.includes('lnT'), 'the border is present; got: ' + children.join(','))
		assert(children.includes('solidFill'), 'the fill is present; got: ' + children.join(','))
		assert(children.indexOf('lnT') < children.indexOf('solidFill'), 'and the border precedes the fill')
	})

	test('a diagonal set last still lands between lnB and the fill', async () => {
		const { presentation, table } = await editable(plainTable)
		const cell = table.cell(0, 0)
		cell.setFillColor('#EEEEEE')
		cell.setBorder('tlToBr', { widthPt: 1, color: 'C00000' })

		const children = assertTcPrOrder(await savedSlide(presentation))
		assert(children.includes('lnTlToBr'), 'the diagonal is present; got: ' + children.join(','))
	})

	test('setBorder writes width, colour and dash, and null removes the edge', async () => {
		const { presentation, table } = await editable(plainTable)
		table.cell(0, 0).setBorder('left', { widthPt: 3, color: '#336699', dash: 'sysDot' })

		let xml = await savedSlide(presentation)
		const lnL = xml.match(/<a:lnL[\s\S]*?<\/a:lnL>/)[0]
		assert(lnL.includes('w="38100"'), '3pt is 38100 EMU; got: ' + lnL)
		assert(lnL.includes('<a:srgbClr val="336699"/>'), 'the colour is normalized past the #; got: ' + lnL)
		assert(lnL.includes('<a:prstDash val="sysDot"/>'), 'the dash is written; got: ' + lnL)

		const { presentation: reloaded, table: reloadedTable } = await reload(presentation)
		const cell = reloadedTable.cell(0, 0)
		assertEqual(cell.borders.left.widthPt, 3, 'the width reads back in points')
		assertEqual(cell.borders.left.dash, 'sysDot', 'and the dash')

		cell.setBorder('left', null)
		xml = await savedSlide(reloaded)
		const firstCell = xml.match(/<a:tc[ >][\s\S]*?<\/a:tc>/)[0]
		assert(!firstCell.includes('<a:lnL'), 'null removes the element entirely; got: ' + firstCell)
	})

	test('a scheme colour wins over a literal, and noFill is distinct from clearing', async () => {
		const { presentation, table } = await editable(plainTable)
		table.cell(0, 0).setBorder('top', { color: 'FF0000', schemeColor: 'accent1' })
		table.cell(0, 1).setBorder('top', { noFill: true })

		const xml = await savedSlide(presentation)
		const lnTs = [...xml.matchAll(/<a:lnT[\s\S]*?<\/a:lnT>/g)].map((m) => m[0])
		assert(lnTs[0].includes('<a:schemeClr val="accent1"/>'), 'the token wins; got: ' + lnTs[0])
		assert(!lnTs[0].includes('srgbClr'), 'and the literal is not also written; got: ' + lnTs[0])
		assert(lnTs[1].includes('<a:noFill/>'), 'an explicitly suppressed edge; got: ' + lnTs[1])
	})

	test('noFill() and setFillColor(null) are different edits', async () => {
		const { presentation, table } = await editable(plainTable)
		table.cell(0, 0).setFillSchemeColor('accent2')
		table.cell(1, 0).setFillSchemeColor('accent2')

		const { presentation: reloaded, table: t } = await reload(presentation)
		t.cell(0, 0).noFill() // explicit transparent — suppresses inherited shading
		t.cell(1, 0).setFillColor(null) // back to inheriting

		const tcPrs = tcPrFillOnly(await savedSlide(reloaded))
		assert(tcPrs[0].includes('<a:noFill/>'), 'the first cell is explicitly transparent; got: ' + tcPrs[0])
		assert(!tcPrs[3].includes('<a:solidFill>'), 'the second cell has no fill of its own; got: ' + tcPrs[3])
		assert(!tcPrs[3].includes('<a:noFill/>'), 'and is not explicitly transparent either; got: ' + tcPrs[3])
	})
})

describe('Table structural edits — rows', () => {
	test('addRow appends an empty auto-height row and removeRow takes one away', async () => {
		const { presentation, table } = await editable(plainTable)
		assertEqual(table.rowCount, 3, 'starting rows')

		const added = table.addRow()
		assertEqual(added.heightEmu, 0, 'a new row is auto-height')
		assertEqual(table.rowCount, 4, 'the row landed')
		assertEqual(table.rows[3].cells.length, 3, 'with one cell per grid column')
		assertEqual(table.rows[3].cells[0].text, '', 'and they are empty')
		assertGridConsistent(table)

		table.removeRow(0)
		assertEqual(table.rowCount, 3, 'a row was removed')
		assertEqual(table.cell(0, 0).text, 'A2', 'and it was the first one')
		assertGridConsistent(table)

		const xml = await savedSlide(presentation)
		assertEqual((xml.match(/<a:tr\b/g) || []).length, 3, 'the saved part agrees')
	})

	test('addRow(index) inserts in the middle', async () => {
		const { table } = await editable(plainTable)
		table.addRow(1)
		assertEqual(table.rowCount, 4, 'the row landed')
		assertEqual(table.cell(0, 0).text, 'A1', 'the row above is untouched')
		assertEqual(table.cell(1, 0).text, '', 'the new row is where it was asked for')
		assertEqual(table.cell(2, 0).text, 'A2', 'and the old row 1 moved down')
		assertGridConsistent(table)
	})

	test('an out-of-range index throws', async () => {
		const { table } = await editable(plainTable)
		for (const fn of [() => table.addRow(9), () => table.removeRow(3), () => table.removeRow(-1)]) {
			assertEqual(codeOfThrow(fn), 'table/row-index-out-of-range', 'the error names its condition')
		}
	})

	test('inserting through a vertical merge extends it rather than splitting it', async () => {
		// The corrupt-file case: an origin claiming more rows than it has continuations.
		const { table } = await editable((p) => {
			p.addSlide().addTable(
				[[{ text: 'tall', options: { rowspan: 3 } }, { text: 'B1' }], [{ text: 'B2' }], [{ text: 'B3' }]],
				{ x: 1, y: 1, w: 9 }
			)
		})
		assertEqual(table.cell(0, 0).rowSpan, 3, 'the span starts at 3')

		table.addRow(1)
		assertEqual(table.cell(0, 0).rowSpan, 4, 'the span grew with the table')
		assert(table.cell(1, 0).isMergeContinuation, 'the new cell continues the span')
		assertGridConsistent(table)
	})

	test('removing a rowspan origin promotes its first continuation', async () => {
		const { table } = await editable((p) => {
			p.addSlide().addTable(
				[[{ text: 'tall', options: { rowspan: 3 } }, { text: 'B1' }], [{ text: 'B2' }], [{ text: 'B3' }]],
				{ x: 1, y: 1, w: 9 }
			)
		})

		table.removeRow(0)
		assertEqual(table.rowCount, 2, 'the row is gone')
		// The merged region survives one row shorter; only the removed row's own text is lost.
		assert(!table.cell(0, 0).isMergeContinuation, 'the first continuation became the origin')
		assertEqual(table.cell(0, 0).rowSpan, 2, 'and inherited the remaining extent')
		assertGridConsistent(table)
	})

	test('removing a row that only continues a merge shortens it', async () => {
		const { table } = await editable((p) => {
			p.addSlide().addTable(
				[[{ text: 'tall', options: { rowspan: 3 } }, { text: 'B1' }], [{ text: 'B2' }], [{ text: 'B3' }]],
				{ x: 1, y: 1, w: 9 }
			)
		})

		table.removeRow(2)
		assertEqual(table.cell(0, 0).rowSpan, 2, 'the span shrank to match')
		assertGridConsistent(table)
	})
})

describe('Table structural edits — columns', () => {
	test('addColumn updates a:tblGrid and every row together', async () => {
		const { presentation, table } = await editable(plainTable)
		assertEqual(table.columnCount, 3, 'starting columns')

		table.addColumn(1, 457200)
		assertEqual(table.columnCount, 4, 'the gridCol landed')
		assertEqual(table.columnWidths[1], 457200, 'with the width asked for')
		assertEqual(table.cell(0, 1).text, '', 'the new column is empty')
		assertEqual(table.cell(0, 2).text, 'B1', 'and the old column 1 moved right')
		assertGridConsistent(table)

		const xml = await savedSlide(presentation)
		assertEqual((xml.match(/<a:gridCol\b/g) || []).length, 4, 'the saved grid agrees')
		// The pair that must stay in step: one a:tc per gridCol, per row.
		for (const tr of xml.match(/<a:tr\b[\s\S]*?<\/a:tr>/g)) {
			assertEqual((tr.match(/<a:tc[ >]/g) || []).length, 4, 'every row has four cells')
		}
	})

	test('removeColumn takes the gridCol and the cells with it', async () => {
		const { presentation, table } = await editable(plainTable)
		table.removeColumn(0)
		assertEqual(table.columnCount, 2, 'the gridCol is gone')
		assertEqual(table.cell(0, 0).text, 'B1', 'and so is the first column')
		assertGridConsistent(table)

		const xml = await savedSlide(presentation)
		for (const tr of xml.match(/<a:tr\b[\s\S]*?<\/a:tr>/g)) {
			assertEqual((tr.match(/<a:tc[ >]/g) || []).length, 2, 'every row lost a cell')
		}
	})

	test('inserting through a horizontal merge widens it', async () => {
		const { table } = await editable((p) => {
			p.addSlide().addTable([[{ text: 'wide', options: { colspan: 3 } }], ['A2', 'B2', 'C2']], { x: 1, y: 1, w: 9 })
		})
		assertEqual(table.cell(0, 0).gridSpan, 3, 'the span starts at 3')

		table.addColumn(1)
		assertEqual(table.cell(0, 0).gridSpan, 4, 'the span grew with the table')
		assertEqual(table.columnCount, 4, 'and the grid did too')
		assertGridConsistent(table)
	})

	test('removing a column inside a merge narrows it and keeps the content', async () => {
		const { table } = await editable((p) => {
			p.addSlide().addTable([[{ text: 'wide', options: { colspan: 3 } }], ['A2', 'B2', 'C2']], { x: 1, y: 1, w: 9 })
		})

		table.removeColumn(1)
		assertEqual(table.columnCount, 2, 'the grid narrowed')
		assertEqual(table.cell(0, 0).gridSpan, 2, 'the span narrowed with it')
		// A covered cell is dropped rather than the origin, so the region keeps its text.
		assertEqual(table.cell(0, 0).text, 'wide', 'and the merged region kept its content')
		assertGridConsistent(table)
	})

	test('an out-of-range column index throws', async () => {
		const { table } = await editable(plainTable)
		assertEqual(
			codeOfThrow(() => table.removeColumn(3)),
			'table/column-index-out-of-range',
			'the error names its condition'
		)
	})
})

describe('Table structural edits — merging', () => {
	test('mergeCells makes one cell of a rectangle and unmergeCell puts it back', async () => {
		const { presentation, table } = await editable(plainTable)
		table.mergeCells(0, 0, 1, 1)

		const origin = table.cell(0, 0)
		assertEqual(origin.gridSpan, 2, 'the origin spans two columns')
		assertEqual(origin.rowSpan, 2, 'and two rows')
		assertEqual(origin.text, 'A1', 'and keeps its content')
		assert(table.cell(0, 1).isMergeContinuation, '(0,1) is covered')
		assert(table.cell(1, 0).isMergeContinuation, '(1,0) is covered')
		assert(table.cell(1, 1).isMergeContinuation, '(1,1) is covered')
		assertEqual(table.cell(1, 1).text, '', 'a covered cell is emptied — it is never rendered')
		assertGridConsistent(table)

		const xml = await savedSlide(presentation)
		assert(xml.includes('gridSpan="2"') && xml.includes('rowSpan="2"'), 'the spans reach the part')

		table.unmergeCell(0, 0)
		assertEqual(table.cell(0, 0).gridSpan, 1, 'the span is gone')
		assertEqual(table.cell(0, 0).rowSpan, 1, 'both of them')
		assert(!table.cell(1, 1).isMergeContinuation, 'and the covered cells are free again')
		assertGridConsistent(table)
	})

	test('a range that cuts through an existing merge is rejected, not silently widened', async () => {
		const { table } = await editable(plainTable)
		table.mergeCells(0, 0, 0, 1)

		// Asking for (0,1)-(0,2) would half-cover the existing merge. Widening it to fit would
		// silently give the caller a region they did not ask for.
		assertEqual(
			codeOfThrow(() => table.mergeCells(0, 1, 0, 2)),
			'table/merge-range-invalid',
			'the overlapping range is rejected'
		)
		assertEqual(table.cell(0, 0).gridSpan, 2, 'and the existing merge is untouched')
		assertGridConsistent(table)
	})

	test('a single-cell range and a covered-cell unmerge both throw', async () => {
		const { table } = await editable(plainTable)
		assertEqual(
			codeOfThrow(() => table.mergeCells(1, 1, 1, 1)),
			'table/merge-range-invalid',
			'a one-cell merge is rejected'
		)

		table.mergeCells(0, 0, 0, 1)
		let message = ''
		try {
			table.unmergeCell(0, 1)
		} catch (err) {
			message = err.message
		}
		// The one place a message is asserted rather than a code: the point of this error is
		// that it tells the caller which cell to address instead, so the coordinates are the
		// contract, not incidental wording.
		assert(message.includes('(0,0)'), 'the message names the origin to use; got: ' + message)
	})

	test('unmerging an unmerged cell is a no-op', async () => {
		const { table } = await editable(plainTable)
		table.unmergeCell(1, 1)
		assertEqual(table.cell(1, 1).gridSpan, 1, 'nothing changed')
		assertGridConsistent(table)
	})
})
