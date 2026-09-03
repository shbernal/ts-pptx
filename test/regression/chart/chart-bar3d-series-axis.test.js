import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, assertIncludes, assertNotIncludes } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// A `bar3d` chart emits a third axis — the series axis (`<c:serAx>`, `makeSerAxis`)
// — that no other chart type produces. Its title, gridlines, tick-label skip and
// number format were entirely uncovered. These cases lock the serAx emission, and
// the shape of `CT_SerAx` that leaves it with no unit element to carry.

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
		name: 'serLabelFormatCode drives the axis numFmt, and the axis carries no unit of any kind',
		fn: async () => {
			// `CT_SerAx` ends at `tickLblSkip`/`tickMarkSkip`/`extLst`: it has no slot for the
			// numeric units nor for the three time units, so the axis emits none and the options
			// that used to write them are gone from the type.
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar3d,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					serLabelFormatCode: '0.0',
				})
			})
			const serAx = serAxBlock(await chartXml(zip))
			assertIncludes(serAx, '<c:numFmt formatCode="0.0" sourceLinked="0"/>', 'serLabelFormatCode drives numFmt')
			for (const tag of ['baseTimeUnit', 'majorTimeUnit', 'minorTimeUnit', 'majorUnit', 'minorUnit']) {
				assertNotIncludes(serAx, `<c:${tag}`, `CT_SerAx has no ${tag}`)
			}
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
			// The category axis is where the time units live, and this is its date-axis arm.
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.line,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					catLabelFormatCode: '0',
					catAxisMajorTimeUnit: 'fortnights',
				})
			})
			const xml = await chartXml(zip)
			assertNotIncludes(xml, 'fortnights', 'invalid time unit is dropped, not emitted')
			assertNotIncludes(xml, '<c:majorTimeUnit', 'the whole majorTimeUnit element is omitted after validation')
		},
	},
	{
		name: 'the series axis line takes a width and a dash, like the other two axes',
		fn: async () => {
			// The series axis could be shown and coloured but not sized or dashed: the emitter
			// hardcoded one point and `solid` where `c:catAx` and `c:valAx` read the caller's
			// options. Nothing said so -- the options simply did not exist -- so a 3-D chart
			// styled to match its siblings came out with one axis drawn differently.
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, {
					type: ChartType.bar3d,
					x: 1,
					y: 1,
					w: 6,
					h: 4,
					serAxisLineShow: true,
					serAxisLineColor: '4472C4',
					serAxisLineSize: 3,
					serAxisLineStyle: 'dash',
				})
			})
			const serAx = serAxBlock(await chartXml(zip))
			assertIncludes(serAx, '<a:ln w="38100"', 'serAxisLineSize 3pt reaches a:ln/@w')
			assertIncludes(serAx, '<a:prstDash val="dash"/>', 'serAxisLineStyle reaches a:prstDash')
			assertIncludes(serAx, '<a:srgbClr val="4472C4"/>', 'and the colour is still honoured')
		},
	},
	{
		name: 'a series axis stating no width or dash is unchanged',
		fn: async () => {
			// The defaults are the values that were hardcoded, so every existing deck emits
			// exactly what it did before the options existed.
			const { zip } = await build((p) => {
				p.addSlide().addChart(DATA, { type: ChartType.bar3d, x: 1, y: 1, w: 6, h: 4 })
			})
			const serAx = serAxBlock(await chartXml(zip))
			assertIncludes(serAx, '<a:ln w="12700"', 'the default width is one point')
			assertIncludes(serAx, '<a:prstDash val="solid"/>', 'and the default dash is solid')
		},
	},
])
