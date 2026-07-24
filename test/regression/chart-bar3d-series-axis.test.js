import { ChartType } from '../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertIncludes,
	assertNotIncludes,
} from '../helpers.js'

// A `bar3d` chart emits a third axis — the series axis (`<c:serAx>`, `makeSerAxis`)
// — that no other chart type produces. Its title, gridlines, tick-label skip,
// number format, and time-unit options were entirely uncovered. These cases lock
// the serAx emission and its input validation (garbage time units warn and are
// dropped rather than corrupting the chart).

function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, path)
}

/** The `<c:serAx>…</c:serAx>` block (the bar3d series axis). */
function serAxBlock(xml) {
	const match = xml.match(/<c:serAx>[\s\S]*?<\/c:serAx>/)
	assert(match, 'expected a <c:serAx> block in bar3d chart; got: ' + xml)
	return match[0]
}

const DATA = [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]

defineRegressionSuite('Chart bar3d series axis', [
	{
		name: 'bar3d emits a c:serAx with a series-axis title and tick-label frequency',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar3d,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					showSerAxisTitle: true,
					serAxisTitle: 'Depth',
					serAxisLabelFrequency: '2',
				})
			})
			const xml = await chartXml(zip)
			assertIncludes(xml, '<c:bar3DChart>', 'chart group is bar3DChart')
			const serAx = serAxBlock(xml)
			assertIncludes(serAx, '<a:t>Depth</a:t>', 'serAxisTitle renders into the axis title')
			assertIncludes(serAx, '<c:tickLblSkip val="2"/>', 'serAxisLabelFrequency flows into tickLblSkip')
			assertIncludes(serAx, '<c:delete val="0"/>', 'axis is shown by default (delete=0)')
		},
	},
	{
		name: 'serAxisHidden deletes the series axis (delete=1)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar3d,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					serAxisHidden: true,
				})
			})
			const serAx = serAxBlock(await chartXml(zip))
			assertIncludes(serAx, '<c:delete val="1"/>', 'serAxisHidden sets delete=1')
		},
	},
	{
		name: 'serLabelFormatCode drives the axis numFmt and enables valid time units',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar3d,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					serLabelFormatCode: '0.0',
					serAxisBaseTimeUnit: 'days',
					serAxisMajorTimeUnit: 'months',
					serAxisMinorTimeUnit: 'years',
				})
			})
			const serAx = serAxBlock(await chartXml(zip))
			assertIncludes(serAx, '<c:numFmt formatCode="0.0" sourceLinked="0"/>', 'serLabelFormatCode drives numFmt')
			assertIncludes(serAx, '<c:baseTimeUnit  val="days"/>', 'base time unit emitted (lowercased)')
			assertIncludes(serAx, '<c:majorTimeUnit val="months"/>', 'major time unit emitted')
			assertIncludes(serAx, '<c:minorTimeUnit val="years"/>', 'minor time unit emitted')
		},
	},
	{
		name: 'serAxisLabelPos flows through as the tickLblPos value (each value round-trips)',
		fn: async () => {
			// Regression: the value was previously only read as a truthiness test — a
			// missing pair of parentheses (`a || b === c ? x : y`) made every set value
			// emit val="low". Each explicit position must now round-trip verbatim.
			for (const pos of ['none', 'low', 'high', 'nextTo']) {
				const { zip } = await build((p) => {
					p.addSlide().addChart(DATA, {
						type: ChartType.bar3d,
						x: 1,
						y: 1,
						w: 6,
						h: 4,
						serAxisLabelPos: pos,
					})
				})
				const serAx = serAxBlock(await chartXml(zip))
				assertIncludes(serAx, `<c:tickLblPos val="${pos}"/>`, `serAxisLabelPos '${pos}' round-trips into tickLblPos`)
			}
		},
	},
	{
		name: 'serAxisLabelPos unset defaults the tickLblPos by bar direction',
		fn: async () => {
			// bar3d defaults to barDir 'col', so the unset default is 'low'.
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, { type: ChartType.bar3d, x: 1, y: 1, w: 6, h: 4 })
			})
			const serAx = serAxBlock(await chartXml(zip))
			assertIncludes(serAx, '<c:tickLblPos val="low"/>', 'unset serAxisLabelPos falls back to the col default (low)')
		},
	},
	{
		name: 'a garbage time unit warns and is dropped (chart stays valid)',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar3d,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					serLabelFormatCode: '0',
					serAxisMajorTimeUnit: 'fortnights',
				})
			})
			const serAx = serAxBlock(await chartXml(zip))
			assertNotIncludes(serAx, 'fortnights', 'invalid time unit is dropped, not emitted')
			assertNotIncludes(serAx, '<c:majorTimeUnit', 'the whole majorTimeUnit element is omitted after validation')
		},
	},
])
