// A `preserve` import and the source master's p:otherStyle, which a table cell's text takes its size
// and italic from (see table-text-inheritance.test.js).
//
// table-text-inheritance.pptx's p:otherStyle lvl1 is 14pt and italic; default-text-style.pptx's is
// 18pt and upright. Pasting the one slide into the other in PowerPoint
// (test/read/fixtures/authoring/probe-table-text-paste.ps1) with "Keep Source Formatting" leaves
// every cell 14pt italic, with StyledTable's header still bold from its table style; with "Use
// Destination Theme" every cell turns 18pt upright. `preserve` is the first. PowerPoint gets there
// by keeping the slide on a copy of the source master, writing nothing onto the cells; a `preserve`
// import rebinds to the destination master, so it bakes the values onto the runs instead.

import { describe, test } from 'vitest'

import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const TABLES = ['StyledTable', 'NoGridTable', 'NoStyleTable']

/** The first run of every cell of the table named `name` on `slide`, row by row. */
function cellRuns(slide, name) {
	const shape = slide.shapes.find((s) => s.name === name && s.shapeType === 'graphicFrame' && s.table)
	assert(shape, `expected a table named ${name}`)
	return shape.table.rows.map((row) => row.cells.map((cell) => cell.textFrame.paragraphs[0].runs[0]))
}

/** Save `target` and reopen it at the slide whose partname is `partName`. */
async function reopened(target, partName) {
	const slide = (await Presentation.load(await target.save())).slides.find((s) => s.partName === partName)
	assert(slide, 'the imported slide is found after a save')
	return slide
}

describe("importSlide({ theme: 'preserve' }) keeps what table cells took from the source p:otherStyle", () => {
	test('every cell keeps the source size and italic in a deck whose p:otherStyle differs', async () => {
		const target = await openFixture('default-text-style')
		const imported = target.importSlide(await openFixture('table-text-inheritance'), 0, { theme: 'preserve' })
		const slide = await reopened(target, imported.partName)
		for (const name of TABLES) {
			for (const run of cellRuns(slide, name).flat()) {
				assertEqual(run.resolvedSizePt, 14, `${name}: the source 14pt, not the destination's 18pt`)
				assertEqual(run.resolvedItalic, true, `${name}: the source italic`)
			}
		}
	})

	// StyledTable's header takes its bold from Medium Style 2's firstRow text style, which states no
	// italic. The destination does not define that style, so the header resolves nothing for bold
	// through the read model; PowerPoint paints it from its built-in definition of the same id.
	test('what the table style states is not baked over, and what it leaves is', async () => {
		const target = await openFixture('default-text-style')
		const imported = target.importSlide(await openFixture('table-text-inheritance'), 0, { theme: 'preserve' })
		const slide = await reopened(target, imported.partName)
		for (const run of cellRuns(slide, 'StyledTable')[0]) {
			const rPr = run.element_.getElementsByTagName('a:rPr')[0]
			assertEqual(rPr.getAttribute('b'), null, "the table style's bold is left to it")
			assertEqual(rPr.getAttribute('i'), '1', 'the italic the table style leaves to p:otherStyle is baked')
		}
	})

	test('a table naming a style the source does not define gets its size, and its weight left alone', async () => {
		const target = await openFixture('default-text-style')
		const source = await openFixture('table-text-inheritance')
		const frame = source.slides[0].shapes.find((s) => s.name === 'NoGridTable')
		frame.element_.getElementsByTagName('a:tableStyleId')[0].textContent = '{00000000-0000-0000-0000-000000000000}'
		frame.markDirty()
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const slide = await reopened(target, imported.partName)
		for (const run of cellRuns(slide, 'NoGridTable').flat()) {
			const rPr = run.element_.getElementsByTagName('a:rPr')[0]
			assertEqual(rPr.getAttribute('sz'), '1400', 'a table style never states a size')
			assertEqual(rPr.getAttribute('i'), null, 'what an unknown style states is unknown')
		}
	})

	test('a cell run that states its own size keeps it', async () => {
		const target = await openFixture('default-text-style')
		const source = await openFixture('table-text-inheritance')
		cellRuns(source.slides[0], 'NoStyleTable')[1][0].fontSizePt = 40
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const slide = await reopened(target, imported.partName)
		assertEqual(cellRuns(slide, 'NoStyleTable')[1][0].resolvedSizePt, 40, 'an own size is not baked over')
	})

	test('importShape bakes a lifted table the same way', async () => {
		const target = await openFixture('default-text-style')
		const source = await openFixture('table-text-inheritance')
		const index = source.slides[0].shapes.findIndex((s) => s.name === 'NoGridTable')
		target.importShape(target.slides[0], source.slides[0], index, { theme: 'preserve' })
		const slide = await reopened(target, target.slides[0].partName)
		for (const run of cellRuns(slide, 'NoGridTable').flat()) {
			assertEqual(run.resolvedSizePt, 14, 'the source 14pt')
			assertEqual(run.resolvedItalic, true, 'the source italic')
		}
	})
})
