// Phase 4 table read/edit tests for `pptxgenjs/read` (src/read/api/table.ts).
//
// Contract under test: a table graphic frame exposes Table → rows → cells with
// geometry/merge metadata read from the live DOM; cell text edits (via the
// convenience setter and via the reused Run setters) mutate only the owning
// slide part, survive a save → reopen round-trip, and stay schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

/** First table on any slide of the fixture. */
function firstTable(presentation) {
	for (const slide of presentation.slides) {
		for (const shape of slide.shapes) {
			if (shape.shapeType === 'graphicFrame' && shape.table) return shape.table
		}
	}
	return null
}

async function partBodies(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	const bodies = new Map()
	for (const entry of Object.values(zip.files)) {
		if (entry.dir) continue
		bodies.set(entry.name, await entry.async('uint8array'))
	}
	return bodies
}

function bytesEqual(a, b) {
	return a.length === b.length && a.every((value, index) => value === b[index])
}

describe('Table read model', () => {
	test('reads rows, columns, cells, and grid geometry', async () => {
		const table = firstTable(await open('table'))
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
		const table = firstTable(await open('table'))
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
		const presentation = await open('table')
		const tables = presentation.slides
			.flatMap((slide) => slide.shapes)
			.filter((shape) => shape.shapeType === 'graphicFrame' && shape.table)
			.map((shape) => shape.table)
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
		const presentation = await open('table')
		firstTable(presentation).cell(0, 0).text = 'EDITED'
		const reopened = await Presentation.load(await presentation.save())
		assertEqual(firstTable(reopened).cell(0, 0).text, 'EDITED', 'edited cell text reloads')
	})

	test('cell.text setter preserves the first run formatting', async () => {
		const presentation = await open('table')
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
		const presentation = await open('table')
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
		for (const [name, body] of inputBodies) {
			if (name === dirty) continue
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be untouched`)
		}
	})

	test.skipIf(!validatorInstalled)('an edited table stays schema-valid', async () => {
		const presentation = await open('table')
		const table = firstTable(presentation)
		table.cell(0, 0).text = 'A'
		table.cell(0, 1).textFrame.paragraphs[0].runs[0].text = 'B'
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
