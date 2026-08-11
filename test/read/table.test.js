// Phase 4 table read/edit tests for `ts-pptx/read` (src/read/api/table.ts).
//
// Contract under test: a table graphic frame exposes Table → rows → cells with
// geometry/merge metadata read from the live DOM; cell text edits (via the
// convenience setter and via the reused Run setters) mutate only the owning
// slide part, survive a save → reopen round-trip, and stay schema-valid.

import { readFile } from 'node:fs/promises'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { bytesEqual, assert, assertEqual, partBodies, assertUnchangedExcept } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { fixturePath, openFixture } from './corpus.js'

const validatorInstalled = await validatorAvailable()

/** First table on any slide of the fixture. */
function firstTable(presentation) {
	for (const slide of presentation.slides) {
		for (const shape of slide.shapes) {
			if (shape.shapeType === 'graphicFrame' && shape.table) return shape.table
		}
	}
	return null
}

/** All tables on any slide of the fixture, in document order. */
function allTables(presentation) {
	return presentation.slides
		.flatMap((slide) => slide.shapes)
		.filter((shape) => shape.shapeType === 'graphicFrame' && shape.table)
		.map((shape) => shape.table)
}

/**
 * The fixture's `TableWithFormattedCells` table (slide 3): its cells carry an
 * explicit `a:tcPr/a:solidFill` of `accent3` with `lumMod`/`lumOff`, so it is the
 * one table that exercises the cell-fill accessors.
 */
function formattedTable(presentation) {
	return allTables(presentation).find((table) => table.cell(0, 0)?.fillSchemeColor === 'accent3') ?? null
}

describe('Table read model', () => {
	test('reads rows, columns, cells, and grid geometry', async () => {
		const table = firstTable(await openFixture('table'))
		assert(table, 'table fixture has a table')
		// The first table in the fixture is 3 rows × 4 columns of "cell".
		assertEqual(table.rowCount, 3, 'row count')
		assertEqual(table.columnCount, 4, 'column count')
		assertEqual(table.rows.length, 3, 'rows array length')
		assertEqual(table.rows[0].cells.length, 4, 'first row cell count')
		assert(
			table.columnWidths.every((w) => w > 0),
			'each column has a positive width'
		)
		assert(table.rows[0].heightEmu > 0, 'row height resolves')
	})

	test('reads cell text and merge metadata', async () => {
		const table = firstTable(await openFixture('table'))
		const cell = table.cell(0, 0)
		assert(cell, 'cell (0,0) exists')
		assertEqual(cell.text, 'cell', 'cell text')
		assertEqual(cell.gridSpan, 1, 'default gridSpan is 1')
		assertEqual(cell.rowSpan, 1, 'default rowSpan is 1')
		assertEqual(cell.isMergeContinuation, false, 'plain cell is not a merge continuation')
		assert(cell.textFrame, 'cell exposes a text frame')
	})

	test('firstRowHeader / bandedRows reflect a:tblPr flags', async () => {
		// The fixture has both a plain banded table and one with firstRow="1".
		const presentation = await openFixture('table')
		const tables = presentation.slides
			.flatMap((slide) => slide.shapes)
			.filter((shape) => shape.shapeType === 'graphicFrame')
			.filter((frame) => frame.table)
			.map((frame) => frame.table)
		assert(tables.length >= 2, `expected ≥2 tables, got ${tables.length}`)
		assert(
			tables.some((table) => table.bandedRows),
			'at least one table is banded'
		)
		assert(
			tables.some((table) => table.firstRowHeader),
			'at least one table has a header first row'
		)
	})
})

describe('Table cell editing', () => {
	test('cell.text setter replaces text and survives a reload', async () => {
		const presentation = await openFixture('table')
		firstTable(presentation).cell(0, 0).text = 'EDITED'
		const reopened = await Presentation.load(await presentation.save())
		assertEqual(firstTable(reopened).cell(0, 0).text, 'EDITED', 'edited cell text reloads')
	})

	test('cell.text setter preserves the first run formatting', async () => {
		const presentation = await openFixture('table')
		// First-table cells carry sz="1400"; the replacement run should keep it.
		const before = firstTable(presentation).cell(0, 0).textFrame.paragraphs[0].runs[0].fontSizePt
		assertEqual(before, 14, 'precondition: cell run is 14pt')
		firstTable(presentation).cell(0, 0).text = 'KEEP'
		const reopened = await Presentation.load(await presentation.save())
		const run = firstTable(reopened).cell(0, 0).textFrame.paragraphs[0].runs[0]
		assertEqual(run.text, 'KEEP', 'text replaced')
		assertEqual(run.fontSizePt, 14, 'first-run formatting preserved')
	})

	test('editing a cell via Run setters works (per-run formatting)', async () => {
		const presentation = await openFixture('table')
		const run = firstTable(presentation).cell(1, 1).textFrame.paragraphs[0].runs[0]
		run.text = 'RUN'
		run.bold = true
		const reopened = await Presentation.load(await presentation.save())
		const reread = firstTable(reopened).cell(1, 1).textFrame.paragraphs[0].runs[0]
		assertEqual(reread.text, 'RUN', 'run text reloads')
		assertEqual(reread.bold, true, 'run bold reloads')
	})

	test('editing a cell leaves every other part byte-identical', async () => {
		const input = await readFile(fixturePath('table'))
		const presentation = await Presentation.load(input)
		firstTable(presentation).cell(0, 0).text = 'EDITED'
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await presentation.save())
		const dirty = 'ppt/slides/slide1.xml'
		assert(!bytesEqual(inputBodies.get(dirty), outputBodies.get(dirty)), 'edited slide differs')
		assertUnchangedExcept(inputBodies, outputBodies, [dirty])
	})

	test.skipIf(!validatorInstalled)('an edited table stays schema-valid', async () => {
		const presentation = await openFixture('table')
		const table = firstTable(presentation)
		table.cell(0, 0).text = 'A'
		table.cell(0, 1).textFrame.paragraphs[0].runs[0].text = 'B'
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

describe('Table cell styling', () => {
	// The fixture's `TableWithFormattedCells` cells carry
	// `a:tcPr/a:solidFill/a:schemeClr val="accent3"` with `lumMod`/`lumOff`.
	// Populated `verticalText`/`anchor`/`marginsEmu` paths need a PowerPoint-authored
	// fixture; `table-cell-style.pptx` is it — see the suite below.
	test('resolvedFill resolves a scheme fill through the theme, applying lum transforms', async () => {
		const cell = formattedTable(await openFixture('table')).cell(0, 0)
		const fill = cell.resolvedFill
		assert(fill, 'formatted cell has a resolved fill')
		assertEqual(fill.hex, 'A5A5A5', 'base accent3 hex')
		assertEqual(fill.effectiveHex, 'C9C9C9', 'effective hex after lumMod 60% / lumOff 40%')
		assert(
			fill.transforms.some((t) => t.name === 'lumMod') && fill.transforms.some((t) => t.name === 'lumOff'),
			'lumMod and lumOff transforms are reported'
		)
	})

	test('fillSchemeColor exposes the raw scheme token', async () => {
		assertEqual(formattedTable(await openFixture('table')).cell(0, 0).fillSchemeColor, 'accent3', 'raw scheme token')
	})

	test('a cell with no own fill inherits the table style graph, but exposes no own scheme token', async () => {
		// TableDefault's cells have no a:tcPr fill; the table is Medium Style 2 - Accent 1
		// with banded rows, so its body cells shade from the style's band1H/band2H parts.
		const cell = firstTable(await openFixture('table')).cell(0, 0)
		assertEqual(cell.fillSchemeColor, null, 'fillSchemeColor reads only the own fill -> null')
		// Row 0 of a banded, header-less table is the first body band (band1H); the resolved
		// hex matches what PowerPoint renders (verified via COM against this fixture).
		assertEqual(cell.resolvedFill?.effectiveHex, 'CFD5EA', 'resolvedFill falls back to the band1H style fill')
	})

	test('a cell whose table style defines no fill still reports null', async () => {
		// The slide-2 tables use the built-in "No Style, No Grid" style ({2D5ABB26-…}):
		// it resolves, but its parts define no cell fill, so a bare cell inherits nothing.
		const noStyle = allTables(await openFixture('table')).find(
			(t) => t.styleId === '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
		)
		assert(noStyle, 'fixture has a "No Style, No Grid" table')
		assert(noStyle.resolvedStyle, 'that style still resolves in tableStyles.xml')
		assertEqual(noStyle.cell(0, 0).resolvedFill, null, 'a fill-less style yields no resolved cell fill')
	})

	test('verticalText / anchor / marginsEmu are null when the cell sets none', async () => {
		// No a:tcPr in the fixture carries @vert, @anchor, or @marL/@marR/@marT/@marB,
		// so every cell exercises the unset (null) branch of these accessors.
		const cell = firstTable(await openFixture('table')).cell(0, 0)
		assertEqual(cell.verticalText, null, 'no @vert -> null')
		assertEqual(cell.anchor, null, 'no @anchor -> null')
		assertEqual(cell.marginsEmu, null, 'no tcPr margins -> null')
	})

	test('borders is null for a cell with no a:tcPr border element', async () => {
		// The fixture's cells carry no a:lnL/lnR/lnT/lnB, so both border-getter null
		// branches are exercised: the plain first-table cell has no a:tcPr at all,
		// and the formatted cell has an a:tcPr (a solid fill) but no border child.
		assertEqual(firstTable(await openFixture('table')).cell(0, 0).borders, null, 'no a:tcPr -> null borders')
		assertEqual(formattedTable(await openFixture('table')).cell(0, 0).borders, null, 'a:tcPr without a border -> null')
	})

	test('element_ escape hatches expose the underlying a:tbl / a:tr / a:tc nodes', async () => {
		const table = firstTable(await openFixture('table'))
		const row = table.rows[0]
		const cell = row.cells[0]
		assertEqual(table.element_.nodeName, 'a:tbl', 'table element_ is the a:tbl')
		assertEqual(row.element_.nodeName, 'a:tr', 'row element_ is the a:tr')
		assertEqual(cell.element_.nodeName, 'a:tc', 'cell element_ is the a:tc')
	})
})

describe('Table cell styling (populated a:tcPr paths)', () => {
	// The `table-cell-style.pptx` fixture is a PowerPoint-authored 2×2 table whose
	// cells each isolate one appearance accessor:
	//   (0,0) <a:tcPr vert="vert270"/>                              -> verticalText
	//   (0,1) <a:tcPr anchor="b"/>                                  -> anchor
	//   (1,0) <a:tcPr marL/marR/marT/marB=228600/342900/114300/457200/> -> marginsEmu
	//   (1,1) <a:tcPr/>                                             -> negative control
	// The isolation lets each test assert the populated value AND that the other
	// two accessors stay null on the same cell.
	test('verticalText reports the a:tcPr @vert token (and leaves anchor/margins null)', async () => {
		const cell = firstTable(await openFixture('table-cell-style')).cell(0, 0)
		assertEqual(cell.verticalText, 'vert270', 'populated @vert')
		assertEqual(cell.anchor, null, 'no @anchor on the vert cell')
		assertEqual(cell.marginsEmu, null, 'no tcPr margins on the vert cell')
	})

	test('anchor reports the a:tcPr @anchor token (and leaves vert/margins null)', async () => {
		const cell = firstTable(await openFixture('table-cell-style')).cell(0, 1)
		assertEqual(cell.anchor, 'b', 'populated @anchor')
		assertEqual(cell.verticalText, null, 'no @vert on the anchor cell')
		assertEqual(cell.marginsEmu, null, 'no tcPr margins on the anchor cell')
	})

	test('marginsEmu reports all four a:tcPr insets in EMU (and leaves vert/anchor null)', async () => {
		const cell = firstTable(await openFixture('table-cell-style')).cell(1, 0)
		const m = cell.marginsEmu
		assert(m, 'populated marginsEmu object')
		assertEqual(m.left, 228600, 'marL EMU')
		assertEqual(m.right, 342900, 'marR EMU')
		assertEqual(m.top, 114300, 'marT EMU')
		assertEqual(m.bottom, 457200, 'marB EMU')
		assertEqual(cell.verticalText, null, 'no @vert on the margins cell')
		assertEqual(cell.anchor, null, 'no @anchor on the margins cell')
	})

	test('a bare a:tcPr cell reports null for vert / anchor / margins', async () => {
		const cell = firstTable(await openFixture('table-cell-style')).cell(1, 1)
		assertEqual(cell.verticalText, null, 'no @vert -> null')
		assertEqual(cell.anchor, null, 'no @anchor -> null')
		assertEqual(cell.marginsEmu, null, 'no tcPr margins -> null')
	})

	test.skipIf(!validatorInstalled)('the fixture is schema-valid', async () => {
		const errors = await validateBuf(await readFile(fixturePath('table-cell-style')))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

describe('Table cell horzOverflow', () => {
	// `table-cell-horzoverflow.pptx` is a 1×2 table (`HorzOverflowTable`) whose first cell
	// carries <a:tcPr horzOverflow="overflow"/> and whose second is a bare <a:tcPr/>.
	// PowerPoint re-serialized both (authoring/author-table-cell-horzoverflow.ps1), which is
	// what makes this an oracle rather than a hand-built string: the attribute has no COM or
	// UI surface, so it is injected and then handed back to PowerPoint to save.
	//
	// There is deliberately no explicit-`clip` cell. `clip` is the schema default and
	// PowerPoint strips it on save, so a PowerPoint-authored one cannot exist.
	test('horzOverflow reports the a:tcPr @horzOverflow token', async () => {
		const cell = firstTable(await openFixture('table-cell-horzoverflow')).cell(0, 0)
		assertEqual(cell.horzOverflow, 'overflow', 'populated @horzOverflow')
		assertEqual(cell.verticalText, null, 'no @vert on the overflow cell')
		assertEqual(cell.anchor, null, 'no @anchor on the overflow cell')
	})

	test('a cell without @horzOverflow reports null', async () => {
		const cell = firstTable(await openFixture('table-cell-horzoverflow')).cell(0, 1)
		assertEqual(cell.horzOverflow, null, 'no @horzOverflow -> null')
	})

	test.skipIf(!validatorInstalled)('the fixture is schema-valid', async () => {
		const errors = await validateBuf(await readFile(fixturePath('table-cell-horzoverflow')))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
