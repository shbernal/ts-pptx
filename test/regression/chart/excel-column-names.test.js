import { describe, expect, test } from 'vitest'
import { getExcelColName } from '../../../src/gen/chart/data-refs.ts'
import { build, readEntry, listEntries } from '../../helpers.js'

// Worksheet column names are bijective base 26: A to Z, AA to ZZ, AAA to XFD. The conversion used
// to handle two letters at most, so column 703 came out as `undefinedA`, and a chart with that many
// series pointed its later formulas, and its workbook's own `<dimension>`, at no column at all.
describe('worksheet column names', () => {
	test.each([
		[1, 'A'],
		[26, 'Z'],
		[27, 'AA'],
		[52, 'AZ'],
		[53, 'BA'],
		[702, 'ZZ'],
		[703, 'AAA'],
		[728, 'AAZ'],
		[16384, 'XFD'],
	])('column %i is %s', (index, name) => {
		expect(getExcelColName(index)).toBe(name)
	})

	test('a column past XFD, the last one a worksheet has, is refused', () => {
		let code = null
		try {
			getExcelColName(16385)
		} catch (err) {
			code = err?.code ?? null
		}
		expect(code).toBe('chart/too-many-columns')
	})

	test('a chart with more than 702 series references real columns', async () => {
		const data = Array.from({ length: 703 }, (_, i) => ({ name: `S${i + 1}`, labels: ['a'], values: [i] }))
		const { zip } = await build((p) => p.addSlide().addChart(data, { type: 'bar', x: 1, y: 1, w: 8, h: 4 }))
		const chartPart = listEntries(zip).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
		expect(chartPart).toBeDefined()
		const xml = await readEntry(zip, chartPart)
		expect(xml).not.toContain('undefined')
		expect(xml).toContain('Sheet1!$AAA$1')
	})
})
