// Write→read fidelity for the two table read-model additions in
// src/read/api/table.ts: TableCell.borders (a:tcPr/a:lnL|lnR|lnT|lnB…) and
// Table.styleId (a:tblPr/a:tableStyleId). Both are BLIND SPOTs the writer
// already emits — the `border` cell option and the `tableStyle` table option —
// so each is proven by authoring a deck with the write API, reading it back
// through the deep model, and asserting the extracted values match what was
// written. The writer's bytes are the fixture; the shared harness loads them.
//
// NOTE: the writer emits a full four-side border set on EVERY cell, defaulting
// an unspecified side to <a:ln w="0"><a:noFill/></a:ln>. So a written table's
// cells never read `borders === null`; the null path (a cell whose a:tcPr
// carries no border element at all) is covered from the PowerPoint fixtures in
// table.test.js instead.

import { describe, test } from 'vitest'
import { TableStyle } from '../../dist/node.js'
import { authorRead, firstTable, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** A 2×2 table whose top-left cell carries a full four-side border set. */
function borderedTable(pres) {
	const border = [
		{ type: 'solid', color: 'FF0000', width: 3 }, // top
		{ type: 'dash', color: '00FF00', width: 1 }, // right
		{ type: 'solid', color: '0000FF', width: 2 }, // bottom
		{ type: 'none' }, // left (suppressed)
	]
	const rows = [
		[{ text: 'A', options: { border } }, { text: 'B' }],
		[{ text: 'C' }, { text: 'D' }],
	]
	pres.addSlide().addTable(rows, { x: 1, y: 1, w: 8, colW: [4, 4] })
}

describe('Table.styleId — a:tblPr/a:tableStyleId', () => {
	test('a table authored with a built-in tableStyle reads its GUID back', async () => {
		const styleId = TableStyle.MEDIUM_STYLE_2_ACCENT_1
		const { presentation } = await authorRead((pres) => {
			pres
				.addSlide()
				.addTable([[{ text: 'x' }, { text: 'y' }]], { x: 1, y: 1, w: 6, tableStyle: styleId, hasHeader: true })
		})
		const table = firstTable(presentation)
		assert(table, 'authored table is read back')
		assertEqual(table.styleId, styleId, 'styleId matches the authored GUID')
		assertEqual(table.firstRowHeader, true, 'the header flag still reads alongside the style id')
	})

	test('a table authored without a tableStyle reports a null styleId', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'x' }]], { x: 1, y: 1, w: 4 })
		})
		assertEqual(firstTable(presentation).styleId, null, 'no a:tableStyleId → null')
	})
})

describe('TableCell.borders — a:tcPr/a:lnL|lnR|lnT|lnB', () => {
	test('a fully bordered cell reads each side (width / dash / colour / suppressed)', async () => {
		const { presentation } = await authorRead(borderedTable)
		const borders = firstTable(presentation).cell(0, 0).borders
		assert(borders, 'the bordered cell surfaces borders')

		// Writer maps the [top,right,bottom,left] option tuple onto a:lnT/lnR/lnB/lnL.
		assertEqual(borders.top.widthPt, 3, 'top width in points')
		assertEqual(borders.top.dash, 'solid', 'a solid border reads prstDash "solid"')
		assertEqual(borders.top.color, 'FF0000', 'top colour resolves to the authored hex')
		assertEqual(borders.top.noFill, false, 'a drawn border is not noFill')

		assertEqual(borders.right.widthPt, 1, 'right width in points')
		assertEqual(borders.right.dash, 'sysDash', 'a dash border reads prstDash "sysDash"')
		assertEqual(borders.right.color, '00FF00', 'right colour')

		assertEqual(borders.bottom.widthPt, 2, 'bottom width in points')
		assertEqual(borders.bottom.color, '0000FF', 'bottom colour')

		// A type:'none' side is emitted as <a:lnL w="0"><a:noFill/></a:lnL>.
		assert(borders.left, 'the suppressed side is still present as an element')
		assertEqual(borders.left.noFill, true, 'the none side reads noFill')
		assertEqual(borders.left.color, null, 'a noFill side has no colour')

		// The writer authors only the four edges, never the diagonals.
		assertEqual(borders.tlToBr, null, 'no diagonal ╲ authored')
		assertEqual(borders.blToTr, null, 'no diagonal ╱ authored')
	})

	test('a cell given no border option still reads a four-side noFill set', async () => {
		// The writer defaults every unspecified cell to four w=0 noFill edges, so
		// cell D reads a border object with all four sides suppressed (not null).
		const borders = firstTable(await authorRead(borderedTable).then((r) => r.presentation)).cell(1, 1).borders
		assert(borders, 'even an unstyled cell carries an authored border set')
		for (const side of ['left', 'right', 'top', 'bottom']) {
			assert(borders[side], `${side} edge is present`)
			assertEqual(borders[side].noFill, true, `${side} edge reads noFill`)
			assertEqual(borders[side].color, null, `${side} edge has no colour`)
		}
	})

	test('a written border reports an empty transform list, not a missing one', async () => {
		// The distinction issue #26 is about: `transforms: []` has to mean "this edge
		// stated none", which is only readable as such because the field exists at all.
		const { presentation } = await authorRead(borderedTable)
		const top = firstTable(presentation).cell(0, 0).borders.top
		assert(top.resolvedColor, 'a drawn border carries a full ResolvedColor')
		assertEqual(top.resolvedColor.hex, 'FF0000', 'base hex is the authored literal')
		assertEqual(top.resolvedColor.transforms.length, 0, 'a bare srgbClr carries no transform children')
		assertEqual(top.resolvedColor.effectiveHex, top.color, 'the flat color mirrors resolvedColor.effectiveHex')
		assertEqual(firstTable(presentation).cell(0, 0).borders.left.resolvedColor, null, 'a noFill edge resolves none')
	})

	test.skipIf(!validatorInstalled)('the authored bordered/styled decks are schema-valid', async () => {
		const bordered = await authorRead(borderedTable)
		assertEqual((await schemaErrors(bordered.buf)).length, 0, 'bordered deck validates')
		const styled = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'x' }]], { x: 1, y: 1, w: 4, tableStyle: TableStyle.LIGHT_STYLE_1 })
		})
		assertEqual((await schemaErrors(styled.buf)).length, 0, 'styled deck validates')
	})
})
