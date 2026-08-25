import { ChartType } from '../../../dist/node.js'
import { defineRegressionSuite, build, assert, assertEqual } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// The default series palettes are shorter than the number of series or data points a caller may
// hand in, so every palette lookup has to wrap. It used to draw a *random* colour past the end
// instead, which made the same deck emit different bytes on every build -- invisible to the
// byte-identity harness, because no showcase deck reaches that far into a palette.
//
// The cases below go past the end on purpose. They assert the two properties the wraparound has
// to hold: the colour repeats from the start of the palette, and a second build of the same
// chart is byte-identical to the first.

const SERIES_COUNT = 12 // longer than either default palette

/** One `{ name, labels, values }` row per series, all with the same two categories. */
function seriesRows(count) {
	return Array.from({ length: count }, (_unused, idx) => ({
		name: `Series ${idx + 1}`,
		labels: ['A', 'B'],
		values: [idx + 1, idx + 2],
	}))
}

/** The `<a:srgbClr val>` of each `<c:ser>`'s own `<c:spPr>` fill, in series order. */
function seriesFills(xml) {
	return xml
		.split('<c:ser>')
		.slice(1)
		.map((ser) => /<c:spPr>.*?<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"\/>/s.exec(ser)?.[1])
}

/** The `<a:srgbClr val>` of each `<c:dPt>` fill, in point order. */
function pointFills(xml) {
	return [...xml.matchAll(/<c:dPt>.*?<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"\/>/gs)].map((m) => m[1])
}

async function barXml(seriesTotal = SERIES_COUNT) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(seriesRows(seriesTotal), { type: ChartType.bar, x: 1, y: 1, w: 6, h: 4 })
	})
	return chartXml(zip)
}

async function pieXml(sliceTotal = SERIES_COUNT) {
	const labels = Array.from({ length: sliceTotal }, (_unused, idx) => `Slice ${idx + 1}`)
	const values = Array.from({ length: sliceTotal }, (_unused, idx) => idx + 1)
	const { zip } = await build((p) => {
		p.addSlide().addChart([{ name: 'Share', labels, values }], { type: ChartType.pie, x: 1, y: 1, w: 6, h: 4 })
	})
	return chartXml(zip)
}

defineRegressionSuite('chart palette wraparound', [
	{
		name: 'bar series past the end of the palette repeat it from the start',
		fn: async () => {
			const fills = seriesFills(await barXml())
			assertEqual(fills.length, SERIES_COUNT, 'one fill per series')
			assert(
				fills.every((fill) => typeof fill === 'string'),
				'every series carries a solid fill; got: ' + JSON.stringify(fills)
			)
			const period = new Set(fills).size
			assert(period < SERIES_COUNT, `the palette is shorter than ${SERIES_COUNT} series; got ${period} distinct`)
			fills.forEach((fill, idx) => {
				assertEqual(fill, fills[idx % period], `series ${idx} repeats the colour of series ${idx % period}`)
			})
		},
	},
	{
		name: 'pie points past the end of the palette repeat it from the start',
		fn: async () => {
			const fills = pointFills(await pieXml())
			assertEqual(fills.length, SERIES_COUNT, 'one c:dPt per slice')
			const period = new Set(fills).size
			assert(period < SERIES_COUNT, `the palette is shorter than ${SERIES_COUNT} slices; got ${period} distinct`)
			fills.forEach((fill, idx) => {
				assertEqual(fill, fills[idx % period], `slice ${idx} repeats the colour of slice ${idx % period}`)
			})
		},
	},
	{
		// The property the byte-identity harness cannot check, because its corpus never reaches
		// past a palette's end.
		name: 'building the same over-long chart twice emits identical XML',
		fn: async () => {
			assertEqual(await barXml(), await barXml(), 'bar chart XML is reproducible')
			assertEqual(await pieXml(), await pieXml(), 'pie chart XML is reproducible')
		},
	},
])
