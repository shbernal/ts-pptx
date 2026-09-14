import TsPptx, { ChartType, InvalidOptionError } from '../../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	assert,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
	captureDiagnostics,
} from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// A combo subchart's `options` is a whole `ChartOpts`, merged over the chart's options when the part
// is written. Only thirteen hand-listed keys used to be checked, and none of the defaults that depend
// on the chart type ran for a subchart. Each case states an option on a subchart and holds it to what
// the same option does on a chart of its own.

const BASE = { x: 1, y: 1, w: 6, h: 3 }
const BAR = [{ name: 'Bar', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]
const LINE = [{ name: 'Line', labels: ['A', 'B', 'C'], values: [3, 2, 1] }]
const XY = [
	{ name: 'X', values: [1, 2, 3] },
	{ name: 'Y', values: [4, 5, 6], labels: ['a', 'b', 'c'] },
]

/** Build one chart and return its part with the diagnostics it raised. */
async function chartWith(data, options = {}) {
	const {
		result: xml,
		codes,
		messages,
	} = await captureDiagnostics(async () => {
		const { zip } = await build((p) => {
			p.addSlide().addChart(data, { ...BASE, ...options })
		})
		return chartXml(zip)
	})
	return { xml, codes, messages }
}

/** A bar and line combo, with `barOptions` on the bar subchart. */
const barAndLine = (barOptions, chartOptions = {}) =>
	chartWith(
		[
			{ type: ChartType.bar, data: BAR, options: barOptions },
			{ type: ChartType.line, data: LINE, options: {} },
		],
		chartOptions
	)

/** The first match of `pattern` in `xml`, or `''`. */
const first = (xml, pattern) => xml.match(pattern)?.[0] ?? ''

defineRegressionSuite('Combo subchart normalization', [
	{
		name: 'a subchart shadow is checked as a chart shadow is',
		fn: async () => {
			const shadow = { type: 'weird', angle: 9999 }
			const combo = await barAndLine({ shadow })
			const single = await chartWith(BAR, { type: ChartType.bar, shadow })
			assertNotIncludes(combo.xml, 'weirdShdw', 'no element named after the unknown type')
			const outer = first(single.xml, /<a:outerShdw[^>]*>/)
			assert(outer.length > 0, 'the single chart draws an outer shadow')
			assertIncludes(combo.xml, outer, 'the subchart draws the same shadow')
			for (const code of ['shadow/invalid-type', 'shadow/angle-out-of-range']) {
				assert(combo.codes.includes(code), `${code} is raised; got ${JSON.stringify(combo.codes)}`)
			}
			assertEqual(shadow.type, 'weird', "the caller's shadow is not rewritten")
		},
	},
	{
		name: 'a subchart dataBorder takes the border defaults',
		fn: async () => {
			const dataBorder = { width: -2, color: 'red' }
			const combo = await barAndLine({ dataBorder })
			const single = await chartWith(BAR, { type: ChartType.bar, dataBorder })
			const line = first(first(single.xml, /<c:barChart>[\s\S]*?<\/c:barChart>/), /<a:ln [^>]*>.*?<\/a:ln>/)
			assertIncludes(line, 'w="9525"', 'the single chart takes the 0.75pt default')
			assertIncludes(first(combo.xml, /<c:barChart>[\s\S]*?<\/c:barChart>/), line, 'and so does the subchart')
			assert(
				combo.codes.includes('chart/option-out-of-range'),
				`the width is reported; got ${JSON.stringify(combo.codes)}`
			)
		},
	},
	{
		name: 'a subchart radarStyle is checked, and its alias resolved',
		fn: async () => {
			const { xml, codes } = await chartWith([
				{ type: ChartType.radar, data: BAR, options: { radarStyle: 'bogus' } },
				{ type: ChartType.radar, data: LINE, options: { radarStyle: 'markers' } },
			])
			assertNotIncludes(xml, 'val="bogus"', 'the unknown style is not written')
			assertIncludes(xml, '<c:radarStyle val="standard"/>', 'it takes the default')
			assertIncludes(xml, '<c:radarStyle val="marker"/>', 'the alias takes its wire spelling')
			assert(codes.includes('chart/invalid-option-value'), `and warns; got ${JSON.stringify(codes)}`)
		},
	},
	{
		// The `dataLabelFormatScatter = 'custom'` default keyed on the chart's `_type`, which for a combo
		// is the subchart list, so a scatter subchart with `showLabel` drew nothing.
		name: 'a scatter subchart with showLabel draws the labels a scatter chart draws',
		fn: async () => {
			const combo = await chartWith([
				{ type: ChartType.bar, data: BAR, options: {} },
				{
					type: ChartType.scatter,
					data: XY,
					options: { showLabel: true, secondaryValAxis: true, secondaryCatAxis: true },
				},
			])
			const single = await chartWith(XY, { type: ChartType.scatter, showLabel: true })
			const count = (xml) => (xml.match(/<c:dLbl>/g) ?? []).length
			assert(count(single.xml) > 0, 'the single scatter draws labels')
			assertEqual(count(combo.xml), count(single.xml), 'the scatter subchart draws as many')
		},
	},
	{
		// The stacked default gap was decided once for the whole chart, so a clustered bar under a
		// stacked chart level took 50.
		name: 'the default gap width is decided per subchart',
		fn: async () => {
			const clustered = await barAndLine({ barGrouping: 'clustered' }, { barGrouping: 'stacked' })
			const bars = first(clustered.xml, /<c:barChart>[\s\S]*?<\/c:barChart>/)
			assertIncludes(bars, '<c:grouping val="clustered"/>', 'the subchart keeps its grouping')
			assertIncludes(bars, '<c:gapWidth val="150"/>', 'and the clustered gap a clustered bar chart takes')

			const stacked = await barAndLine({ barGrouping: 'stacked' })
			assertIncludes(stacked.xml, '<c:gapWidth val="50"/>', 'a stacked subchart still takes 50')
			const stated = await barAndLine({ barGrouping: 'clustered' }, { barGrouping: 'stacked', barGapWidthPct: 80 })
			assertIncludes(stated.xml, '<c:gapWidth val="80"/>', 'a stated chart-level gap is kept')
		},
	},
	{
		name: 'a subchart chartColorsOpacity is checked when the chart is added',
		fn: async () => {
			let thrown = null
			try {
				new TsPptx().addSlide().addChart(
					[
						{ type: ChartType.bar, data: BAR, options: { chartColorsOpacity: NaN } },
						{ type: ChartType.line, data: LINE, options: {} },
					],
					BASE
				)
			} catch (err) {
				thrown = err
			}
			assert(thrown instanceof InvalidOptionError, `a NaN opacity throws at addChart; got ${thrown}`)
			assertEqual(thrown.code, 'chart/option-non-finite', 'with the shared code')
			const over = await barAndLine({ chartColorsOpacity: 150 })
			assertIncludes(over.xml, '<a:alpha val="100000"/>', 'an opacity above 100 clamps')
		},
	},
	{
		// `scrubGridLine` mutates its argument, and the chart-level `barSeriesLine` was handed to it as
		// the caller's own object.
		name: "a barSeriesLine is scrubbed without rewriting the caller's object",
		fn: async () => {
			const chartLine = { width: 0, cap: 'bevel' }
			const subLine = { width: 0, cap: 'bevel' }
			await chartWith(BAR, { type: ChartType.bar, barGrouping: 'stacked', barSeriesLine: chartLine })
			await barAndLine({ barGrouping: 'stacked', barSeriesLine: subLine })
			assertEqual(JSON.stringify(chartLine), '{"width":0,"cap":"bevel"}', 'the chart-level object is untouched')
			assertEqual(JSON.stringify(subLine), '{"width":0,"cap":"bevel"}', 'the subchart object is untouched')
		},
	},
])
