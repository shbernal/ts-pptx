import TsPptx, { ChartType } from '../../../dist/node.js'
import JSZip from 'jszip'
import { defineRegressionSuite, build, listEntries, assert, assertIncludes, assertNotIncludes } from '../../helpers.js'

// 1x1 PNG (red pixel) for image-only deck case
const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='

function chartsOrEmbeddingsEntries(zip) {
	return listEntries(zip).filter(
		(p) =>
			p.startsWith('ppt/charts/') || p === 'ppt/charts' || p.startsWith('ppt/embeddings/') || p === 'ppt/embeddings'
	)
}

defineRegressionSuite('Chart embedding parts [legacy bug-17]', [
	{
		name: 'empty deck (text-only) does not create ppt/charts or ppt/embeddings dirs',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addText('hello world', { x: 1, y: 1, w: 4, h: 1 })
			})
			const stray = chartsOrEmbeddingsEntries(zip)
			assert(
				stray.length === 0,
				'expected no charts/embeddings entries for chart-free deck; got: ' + JSON.stringify(stray)
			)
		},
	},
	{
		name: 'image-only deck does not create ppt/charts or ppt/embeddings dirs',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addImage({ data: PNG_DATA, x: 1, y: 1, w: 1, h: 1 })
			})
			const stray = chartsOrEmbeddingsEntries(zip)
			assert(
				stray.length === 0,
				'expected no charts/embeddings entries for image-only deck; got: ' + JSON.stringify(stray)
			)
		},
	},
	{
		name: 'chart-present deck still creates chart and embedding parts (regression)',
		fn: async () => {
			const pres = new TsPptx()
			const slide = pres.addSlide()
			const data = [
				{
					name: 'Series 1',
					labels: ['Cat A', 'Cat B', 'Cat C'],
					values: [10, 20, 30],
				},
			]
			slide.addChart(data, { type: ChartType.bar, x: 1, y: 1, w: 6, h: 3 })
			const buf = await pres.toBytes()
			const zip = await JSZip.loadAsync(buf)
			const entries = listEntries(zip)
			const chartEntries = entries.filter((p) => p.startsWith('ppt/charts/'))
			const embedEntries = entries.filter((p) => p.startsWith('ppt/embeddings/'))
			assert(
				chartEntries.length > 0,
				'expected ppt/charts/ entries when chart present; got: ' + JSON.stringify(entries)
			)
			assert(
				embedEntries.length > 0,
				'expected ppt/embeddings/ entries when chart present; got: ' + JSON.stringify(entries)
			)
		},
	},
	{
		// The workbook is the half of a chart nothing else checks. PowerPoint does not parse the
		// embedding when it opens the deck, so a workbook Excel refuses (0x3EC) still gives a deck
		// that opens and paints — the failure only surfaces on "Edit Data", by which point the
		// chart's own cache is already clean, because `numCachePt` dropped the bad point on the way
		// into `chart.xml`. `<v>Infinity</v>` was reaching every family's sheet regardless.
		//
		// A gap is `<v></v>`, which is what a `null` value has always produced here and which Excel
		// reads back as an empty cell; a non-finite number now takes that same path.
		name: 'no family writes a non-finite number into the embedded workbook',
		fn: async () => {
			const labels = ['a', 'b', 'c']
			const cases = [
				['bar', ChartType.bar, [{ name: 'S1', labels, values: [1, Infinity, 3] }]],
				['pie', ChartType.pie, [{ name: 'S1', labels, values: [1, -Infinity, 3] }]],
				['line, NaN', ChartType.line, [{ name: 'S1', labels, values: [1, NaN, 3] }]],
				[
					'scatter',
					ChartType.scatter,
					[
						{ name: 'X', labels, values: [1, 2, 3] },
						{ name: 'Y', labels, values: [4, Infinity, 6] },
					],
				],
				[
					'bubble, in the size column',
					ChartType.bubble,
					[
						{ name: 'X', labels, values: [1, 2, 3] },
						{ name: 'Y', labels, values: [4, 5, 6], sizes: [7, Infinity, 9] },
					],
				],
				[
					'multi-level categories',
					ChartType.bar,
					[{ name: 'S1', labels: [labels, ['Grp', '', '']], values: [1, Infinity, 3] }],
				],
			]

			for (const [what, type, data] of cases) {
				const { buf } = await build((p) => {
					p.addSlide().addChart(data, { type, x: 1, y: 1, w: 6, h: 4 })
				})
				const xlsx = (await JSZip.loadAsync(buf)).file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
				assert(xlsx, `${what}: expected an embedded workbook`)
				const sheet = await (
					await JSZip.loadAsync(await xlsx.async('arraybuffer'))
				)
					.file('xl/worksheets/sheet1.xml')
					.async('string')
				assertNotIncludes(sheet, 'Infinity', `${what}: no infinity reaches the workbook`)
				assertNotIncludes(sheet, 'NaN', `${what}: nor a NaN`)
				assertIncludes(sheet, '<v></v>', `${what}: the bad value leaves the same empty cell a gap does`)
			}

			// The spelling a gap has always had, so the fix above is the only thing that changed.
			const { buf } = await build((p) => {
				p.addSlide().addChart([{ name: 'S1', labels: ['a', 'b', 'c'], values: [1, null, 3] }], {
					type: ChartType.bar,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
				})
			})
			const xlsx = (await JSZip.loadAsync(buf)).file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx')
			const sheet = await (
				await JSZip.loadAsync(await xlsx.async('arraybuffer'))
			)
				.file('xl/worksheets/sheet1.xml')
				.async('string')
			assertIncludes(sheet, '<c r="B3"><v></v></c>', 'a null value is an empty cell')
		},
	},
	{
		// `table1.xml`'s bubble range used the *column* count where a row count belongs, so one
		// workbook stated two different extents for the same sheet. `<tableParts>` is deliberately
		// never emitted, so Excel does not read the part today — but it is relationship-linked
		// from `sheet1.xml.rels`, and the formula was wrong on its face.
		name: "a bubble workbook's table range matches its own sheet dimension",
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart(
					[
						{ name: 'X-Axis', values: [10, 11, 12, 13] },
						{ name: 'Y1', values: [1, 6, 7, 8], sizes: [4, 5, 6, 7] },
					],
					{ type: ChartType.bubble, x: 1, y: 1, w: 6, h: 4 }
				)
			})
			const xlsx = await JSZip.loadAsync(
				await (await JSZip.loadAsync(buf)).file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx').async('arraybuffer')
			)
			const table = await xlsx.file('xl/tables/table1.xml').async('string')
			const sheet = await xlsx.file('xl/worksheets/sheet1.xml').async('string')
			const tableRef = /<table[^>]*\bref="([^"]+)"/.exec(table)?.[1]
			const sheetRef = /<dimension ref="([^"]+)"/.exec(sheet)?.[1]
			assert(tableRef === sheetRef, `one workbook, one extent; table says ${tableRef}, sheet says ${sheetRef}`)
			// Four X values plus the header row, across 1 + 2 columns.
			assert(tableRef === 'A1:C5', `expected A1:C5; got ${tableRef}`)
		},
	},
	{
		name: 'a bubble series with no name still writes the required tableColumn name',
		fn: async () => {
			const { buf } = await build((p) => {
				p.addSlide().addChart([{ values: [10, 11] }, { values: [1, 6], sizes: [4, 5] }], {
					type: ChartType.bubble,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
				})
			})
			const xlsx = await JSZip.loadAsync(
				await (await JSZip.loadAsync(buf)).file('ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx').async('arraybuffer')
			)
			const table = await xlsx.file('xl/tables/table1.xml').async('string')
			const columns = [...table.matchAll(/<tableColumn [^>]*\/>/g)].map((m) => m[0])
			columns.forEach((col) => assert(col.includes('name='), '`name` is required on a tableColumn; got: ' + col))
		},
	},
])
