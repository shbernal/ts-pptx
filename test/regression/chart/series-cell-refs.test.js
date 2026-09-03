import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, assertNotIncludes } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// The embedded workbook lays every series out behind the FIRST series' label columns, one row
// per that series' categories. The chart XML derived both numbers from each series' own labels,
// so labelling only the first series -- the shape the plot builders' own worked example shows --
// made the two sides disagree from series 1 on: a `<c:tx>` naming series A's header, a `<c:val>`
// range running backwards over series A's column, and `Sheet1!$A$2:$$1` for the categories.
// Rendering survives on the inconsistent cache; "Edit Data" opens onto the wrong columns.

/** `[{ name, labels?, values }]` with labels on the first series only. */
const FIRST_LABELLED = [
	{ name: 'Region 1', labels: [['April', 'May', 'June']], values: [17, 26, 53] },
	{ name: 'Region 2', values: [55, 43, 70] },
	{ name: 'Region 3', values: [12, 34, 21] },
]

/** Every `<c:f>` formula in the part, in document order. */
function formulas(xml) {
	return [...xml.matchAll(/<c:f>([^<]*)<\/c:f>/g)].map((m) => m[1])
}

/** The `<c:ser>` blocks, in document order. */
function seriesBlocks(xml) {
	return [...xml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((m) => m[0])
}

/**
 * Assert that every reference in the part addresses a real cell: a range never runs backwards,
 * and no column letter is missing.
 */
function assertWellFormedRefs(xml, label) {
	for (const f of formulas(xml)) {
		assert(/^Sheet1!\$[A-Z]+\$\d+(:\$[A-Z]+\$\d+)?$/.test(f), `${label}: malformed reference ${JSON.stringify(f)}`)
		const range = f.match(/^Sheet1!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$/)
		if (range) {
			assert(Number(range[4]) >= Number(range[2]), `${label}: reversed row range ${JSON.stringify(f)}`)
		}
	}
}

defineRegressionSuite('Series worksheet references', [
	{
		name: 'a bar chart labelled on the first series only addresses each series its own column',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(FIRST_LABELLED, { type: ChartType.bar, x: 1, y: 1, w: 6, h: 3 })
			})
			const xml = await chartXml(zip)
			assertWellFormedRefs(xml, 'bar')
			const sers = seriesBlocks(xml)
			assert(sers.length === 3, `expected three series; got ${sers.length}`)
			// Series names sit in the header row of each series' own column: B1, C1, D1.
			const names = sers.map((ser) => ser.match(/<c:f>(Sheet1!\$[A-Z]+\$1)<\/c:f>/)?.[1])
			assert(
				JSON.stringify(names) === JSON.stringify(['Sheet1!$B$1', 'Sheet1!$C$1', 'Sheet1!$D$1']),
				'each series names its own header cell; got ' + JSON.stringify(names)
			)
			// And each `<c:val>` spans the three category rows of its own column.
			const vals = sers.map((ser) => ser.match(/<c:val><c:numRef><c:f>([^<]*)<\/c:f>/)?.[1])
			assert(
				JSON.stringify(vals) === JSON.stringify(['Sheet1!$B$2:$B$4', 'Sheet1!$C$2:$C$4', 'Sheet1!$D$2:$D$4']),
				'each series reads its own column over the sheet rows; got ' + JSON.stringify(vals)
			)
		},
	},
	{
		name: 'an unlabelled series states no categories rather than an empty range',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(FIRST_LABELLED, { type: ChartType.line, x: 1, y: 1, w: 6, h: 3 })
			})
			const xml = await chartXml(zip)
			assertNotIncludes(xml, '$$', 'no reference may be missing a column letter')
			const cats = seriesBlocks(xml).map((ser) => ser.includes('<c:cat>'))
			assert(
				JSON.stringify(cats) === JSON.stringify([true, false, false]),
				'only the labelled series carries a <c:cat>; got ' + JSON.stringify(cats)
			)
		},
	},
	{
		name: 'the value cache point count matches the sheet, not the series own label count',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(FIRST_LABELLED, { type: ChartType.bar, x: 1, y: 1, w: 6, h: 3 })
			})
			const xml = await chartXml(zip)
			for (const ser of seriesBlocks(xml)) {
				const cache = ser.match(/<c:val>[\s\S]*?<\/c:val>/)?.[0] ?? ''
				const ptCount = Number(cache.match(/<c:ptCount val="(\d+)"\/>/)?.[1])
				const pts = [...cache.matchAll(/<c:pt idx="\d+">/g)].length
				assert(ptCount === 3, `expected ptCount 3; got ${ptCount} in ${cache}`)
				assert(pts === 3, `expected three cached points; got ${pts}`)
			}
		},
	},
	{
		name: 'stock and surface charts address their series the same way',
		fn: async () => {
			for (const type of [ChartType.stock, ChartType.surface]) {
				const series =
					type === ChartType.stock
						? [
								{ name: 'High', labels: [['Mon', 'Tue', 'Wed']], values: [10, 12, 11] },
								{ name: 'Low', values: [7, 8, 6] },
								{ name: 'Close', values: [9, 11, 7] },
							]
						: FIRST_LABELLED
				const { zip } = await build((p) => {
					p.addSlide().addChart(series, { type, x: 1, y: 1, w: 6, h: 3 })
				})
				const xml = await chartXml(zip)
				assertWellFormedRefs(xml, String(type))
				assertNotIncludes(xml, '$$', `${String(type)}: no reference may be missing a column letter`)
			}
		},
	},
])
