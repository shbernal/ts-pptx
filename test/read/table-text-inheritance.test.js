// Where a table cell's text takes what it does not state, on table-text-inheritance.pptx. Every
// tier that could answer is given a value no other tier has:
//
//   p:defaultTextStyle lvl1   28pt, C00000, Courier New
//   master p:otherStyle lvl1  14pt, 0070C0, Georgia, italic
//   theme                     dk1 7030A0, minor Latin font Verdana
//
// and every run on the slide is bare (`<a:rPr lang="en-US"/>`). TextBox is a text box, the control.
// StyledTable carries Medium Style 2 - Accent 1, whose tcTxStyle names the minor font with dk1 for
// the body and lt1 bold for the header row. NoGridTable carries "No Style, No Grid", whose
// tcTxStyle names the minor font and tx1. NoStyleTable carries no table style at all.
//
// Exported to PNG, the text box paints 28pt Courier New in C00000, and every cell paints 14pt
// italic Verdana in 7030A0, apart from StyledTable's header, which paints white and bold. So a
// cell takes its size and italic from p:otherStyle and never reads p:defaultTextStyle, and its
// face and colour come from the table style's text style, or from the minor font and tx1 when no
// table style applies, never from p:otherStyle.

import { describe, test } from 'vitest'

import { assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

/** The first run of every cell of the table named `name`, row by row. */
async function cellRuns(name) {
	for (const shape of (await openFixture('table-text-inheritance')).slides[0].shapes) {
		if (shape.name === name && shape.shapeType === 'graphicFrame' && shape.table) {
			return shape.table.rows.map((row) => row.cells.map((cell) => cell.textFrame.paragraphs[0].runs[0]))
		}
	}
	throw new Error(`no table named ${name}`)
}

/** A run's resolved character properties, in one comparable string. */
function resolved(run) {
	return [
		run.resolvedSizePt,
		run.resolvedFontFace,
		run.resolvedColor?.effectiveHex ?? null,
		run.resolvedBold,
		run.resolvedItalic,
	].join(' ')
}

const BODY = '14 Verdana 7030A0  true'

describe('table cell text inheritance (table-text-inheritance.pptx)', () => {
	test('the text box control resolves through p:defaultTextStyle', async () => {
		for (const shape of (await openFixture('table-text-inheritance')).slides[0].shapes) {
			if (shape.name !== 'TextBox' || shape.shapeType !== 'autoShape' || !shape.textFrame) continue
			assertEqual(resolved(shape.textFrame.paragraphs[0].runs[0]), '28 Courier New C00000  ')
			return
		}
		throw new Error('no text box named TextBox')
	})

	test('a styled table takes its colour, face and bold from the table style and its size and italic from p:otherStyle', async () => {
		const [header, body] = await cellRuns('StyledTable')
		for (const run of header) assertEqual(resolved(run), '14 Verdana FFFFFF true true', 'firstRow: lt1, bold')
		for (const run of body) assertEqual(resolved(run), BODY, 'wholeTbl: dk1')
	})

	test('"No Style, No Grid" names tx1, which maps to dk1', async () => {
		for (const row of await cellRuns('NoGridTable')) {
			for (const run of row) assertEqual(resolved(run), BODY)
		}
	})

	test('a table with no style still takes the minor font and tx1, not p:otherStyle', async () => {
		for (const row of await cellRuns('NoStyleTable')) {
			for (const run of row) assertEqual(resolved(run), BODY)
		}
	})
})
