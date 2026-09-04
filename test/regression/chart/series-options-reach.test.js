import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// `seriesOptions` is indexed by the series' OWN position -- the number it carries in `<c:idx>` and
// `<c:order>` -- not by a position in the `data` array. The two differ on exactly the plots covered
// here: a scatter and a bubble chart take their shared X values from `data[0]`, which is not a
// series and therefore has no entry, so `seriesOptions[0]` styles `data[1]`.
//
// Every case reaches past the palette's own answer: the overrides are colours no default palette
// contains, so a series painted with one could not have got there by accident.

/** Scatter/bubble shape: `data[0]` is the shared X row; every later row is one series. */
const XY = [
	{ name: 'X', values: [1, 2, 3] },
	{ name: 'Y1', values: [4, 5, 6], sizes: [1, 2, 3] },
	{ name: 'Y2', values: [7, 8, 9], sizes: [3, 2, 1] },
]
const BASE = { x: 1, y: 1, w: 6, h: 3 }

/** The `<a:srgbClr val>` of each `<c:ser>`'s own `<c:spPr>` fill, in series order. */
function seriesFills(xml) {
	return xml
		.split('<c:ser>')
		.slice(1)
		.map((ser) => /<c:spPr>.*?<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"\/>/s.exec(ser)?.[1])
}

defineRegressionSuite('seriesOptions reaches the XY and stock plots', [
	{
		name: "a scatter's seriesOptions[0] colours the first Y series, not the X row",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(XY, {
					...BASE,
					type: ChartType.scatter,
					seriesOptions: [{ color: 'FF00FF' }, { color: '00FFFF' }],
				})
			})
			const fills = seriesFills(await chartXml(zip))
			assert(
				fills[0] === 'FF00FF' && fills[1] === '00FFFF',
				'both Y series should take their override in order; got ' + JSON.stringify(fills)
			)
		},
	},
	{
		name: "a scatter's per-series lineSize reaches its stroke",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(XY, {
					...BASE,
					type: ChartType.scatter,
					lineSize: 2,
					seriesOptions: [{ lineSize: 6 }, { lineSize: 0 }],
				})
			})
			const xml = await chartXml(zip)
			const sers = xml.split('<c:ser>').slice(1)
			// 6pt = 76200 EMU; a stated 0 is the caller's "no outline" and emits `<a:noFill/>`.
			assert(sers[0].includes('<a:ln w="76200"'), 'the first series should stroke at 6pt; got ' + sers[0].slice(0, 400))
			assert(
				/<c:spPr>.*?<a:ln><a:noFill\/><\/a:ln>/s.test(sers[1]),
				'the second series asked for no outline; got ' + sers[1].slice(0, 400)
			)
		},
	},
	{
		// The scatter label builders read the font through `labelFontAttrs`/`labelFontChildren`, and
		// both run inside the series loop -- so a per-series override has a referent there where the
		// bubble plot, whose `<c:dLbls>` is one chart-level block, has none.
		name: "a scatter's per-series data-label font overrides the chart's",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(XY, {
					...BASE,
					type: ChartType.scatter,
					showLabel: true,
					dataLabelFormatScatter: 'XY',
					dataLabelColor: '000000',
					seriesOptions: [{ dataLabelColor: 'AB12CD' }],
				})
			})
			const sers = (await chartXml(zip)).split('<c:ser>').slice(1)
			assert(sers[0].includes('AB12CD'), "the first series' labels take its own colour; got " + sers[0].slice(0, 900))
			assert(!sers[1].includes('AB12CD'), 'the second series keeps the chart-level colour')
		},
	},
	{
		name: "a bubble series' colour and outline come from its own entry",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(XY, {
					...BASE,
					type: ChartType.bubble,
					seriesOptions: [{ color: 'FF00FF', lineSize: 6 }],
				})
			})
			const xml = await chartXml(zip)
			const sers = xml.split('<c:ser>').slice(1)
			assert(seriesFills(xml)[0] === 'FF00FF', 'the first bubble series takes its override colour')
			assert(sers[0].includes('<a:ln w="76200"'), 'and its 6pt outline; got ' + sers[0].slice(0, 400))
		},
	},
	{
		// A stock chart's price series draw no line by design and its `<c:dLbls>` is a constant, so
		// `color` has exactly two referents: the volume bar (series 0 of a `vhlc`/`vohlc` style) and
		// the dot marking the close series where there are no up-down bars.
		name: "a stock chart's volume bar and close marker take their series override",
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addChart(
					[
						{ name: 'Volume', labels: ['D1', 'D2'], values: [100, 200] },
						{ name: 'High', labels: ['D1', 'D2'], values: [12, 14] },
						{ name: 'Low', labels: ['D1', 'D2'], values: [8, 9] },
						{ name: 'Close', labels: ['D1', 'D2'], values: [10, 13] },
					],
					{
						...BASE,
						type: ChartType.stock,
						stockStyle: 'vhlc',
						seriesOptions: [{ color: 'FF00FF' }, {}, {}, { color: '00FFFF' }],
					}
				)
			})
			const xml = await chartXml(zip)
			assert(xml.includes('FF00FF'), 'the volume bar takes seriesOptions[0].color')
			assert(xml.includes('00FFFF'), 'the close marker takes seriesOptions[3].color')
		},
	},
])
