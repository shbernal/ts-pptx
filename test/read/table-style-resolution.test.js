// Table-style resolution: a cell with no own fill inherits the shading its
// `a:tableStyleId` supplies from `ppt/tableStyles.xml` (header / banded-row /
// wholeTbl parts), resolved through the slide theme. Every expected hex below was
// verified against what PowerPoint renders, read back per cell over COM from the
// `table-styles.pptx` fixture (three tables, each a distinct built-in style, all
// cells with an empty `<a:tcPr/>`, `firstRow` + `bandRow` on).
import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation, isGraphicFrame } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function tables() {
	const pres = await Presentation.load(await readFile(path.join(__dirname, 'fixtures', 'table-styles.pptx')))
	return pres.slides[0].shapes
		.filter(isGraphicFrame)
		.filter((s) => s.table)
		.map((s) => s.table)
}

function fillHex(table, row, col) {
	return table.cell(row, col)?.resolvedFill?.effectiveHex ?? null
}

describe('Table.resolvedStyle names the referenced tableStyles.xml entry', () => {
	test('each table resolves its GUID to the matching style id + name', async () => {
		const [medium2, medium4, light2] = await tables()
		assertEqual(medium2.resolvedStyle?.styleId, '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}', 'medium2 styleId')
		assertEqual(medium2.resolvedStyle?.name, 'Medium Style 2 - Accent 3', 'medium2 name')
		assertEqual(medium4.resolvedStyle?.name, 'Medium Style 4 - Accent 4', 'medium4 name')
		assertEqual(light2.resolvedStyle?.name, 'Light Style 2 - Accent 1', 'light2 name')
		assert(medium2.resolvedStyle?.element_, 'the raw a:tblStyle element is exposed')
	})
})

describe('TableCell.resolvedFill composes the style graph in precedence order', () => {
	test('firstRow header, banded body rows, and wholeTbl (Medium Style 2 - Accent 3)', async () => {
		const [medium2] = await tables()
		// accent3 = 196B24. Header = accent3 full; body row 1 = band1H (tint 40%);
		// body row 2 = band2H (empty) which falls through to wholeTbl (tint 20%).
		assertEqual(fillHex(medium2, 0, 0), '196B24', 'firstRow -> accent3 full')
		assertEqual(fillHex(medium2, 1, 0), 'CCD4CC', 'first body row -> band1H tint 40%')
		assertEqual(fillHex(medium2, 2, 0), 'E7EBE8', 'second body row -> band2H empty -> wholeTbl tint 20%')
		// The whole row shares one fill (no column conditions).
		assertEqual(fillHex(medium2, 1, 2), 'CCD4CC', 'banding is per-row, not per-cell')
	})

	test('a missing band2H part falls straight to wholeTbl (Medium Style 4 - Accent 4)', async () => {
		const [, medium4] = await tables()
		// This style defines no band2H, so the second body row also uses wholeTbl.
		assertEqual(fillHex(medium4, 0, 0), 'E7F0F7', 'firstRow -> accent4 tint 20%')
		assertEqual(fillHex(medium4, 1, 0), 'CCDFEF', 'first body row -> band1H tint 40%')
		assertEqual(fillHex(medium4, 2, 0), 'E7F0F7', 'second body row -> wholeTbl (no band2H)')
	})

	test('a fillRef header resolves through the theme; a noFill body reads null (Light Style 2 - Accent 1)', async () => {
		const [, , light2] = await tables()
		// firstRow supplies its fill via a:fillRef idx="1" (style-matrix), not a direct
		// a:fill -> resolves to accent1 full. The body rows are wholeTbl a:noFill.
		assertEqual(fillHex(light2, 0, 0), '156082', 'firstRow fillRef -> accent1')
		assertEqual(fillHex(light2, 1, 0), null, 'body row -> wholeTbl noFill -> null')
		assertEqual(fillHex(light2, 2, 0), null, 'body row -> wholeTbl noFill -> null')
	})
})
