import { ChartType } from '../../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	captureDiagnostics,
	assert,
	assertEqual,
} from '../../helpers.js'

// The fill slot of the two `c:spPr` elements a chart part carries: the plot area's (written by
// `makeChartPlotAreaPropsXml`) and the chartSpace-level one (STEP 5 of `makeXmlCharts`).
//
// `ChartPropsFillLine.fill` is typed `ShapeFillProps`, so both areas accept every fill kind the
// shape path does -- `c:spPr` is `a:CT_ShapeProperties`, the same optional `EG_FillProperties`
// group. Both sites used to gate on `fill?.color`, which meant every spelling carrying no colour
// (`gradient`, `pattern`, `image`, and `inherit` once #10 added it) fell to the `<a:noFill/>` arm
// and silently did nothing. See #11.
//
// What makes the gate delicate, and why half of these cases assert that nothing changed:
// `normalizeChartOptions` defaults `plotArea.fill` to `{}`, so EVERY chart reaches the emitter
// with a fill object present. A plain `fill ? ... : '<a:noFill/>'` -- the shape path's shape --
// would therefore paint every plot area ever authored a default grey. `isStatedFill` gates on
// `color || type` for that reason, and the baseline/`{}`/transparency-only cases below are what
// hold that line.

/** The chart part; charts always land at chart1.xml here since each case builds one. */
function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, path)
}

/**
 * The fill slot of both `c:spPr` elements, as raw XML (`''` when no fill child was emitted).
 *
 * `EG_FillProperties` sits between the geometry and `a:ln` in `CT_ShapeProperties`, and `a:ln` is
 * always present here -- both sites emit `<a:ln><a:noFill/></a:ln>` when no border is asked for.
 * So the fill slot is exactly what lies between `<c:spPr>` and the first `<a:ln`, which also means
 * an inherited fill reads back as the empty string rather than needing an absence assertion.
 */
function areaFills(xml) {
	const split = xml.indexOf('</c:chart>')
	assert(split > 0, 'expected a </c:chart> to separate the plot area from the chartSpace spPr')
	const slot = (part) => {
		const open = part.indexOf('<c:spPr>')
		assert(open >= 0, 'expected a <c:spPr> in ' + JSON.stringify(part.slice(0, 80)))
		return part.slice(open + '<c:spPr>'.length, part.indexOf('<a:ln', open))
	}
	// The plot area's spPr is the LAST one before `</c:chart>` -- the axes and series carry their
	// own -- so search backwards from the closing `</c:plotArea>`.
	const beforeChart = xml.slice(0, split)
	return {
		plot: slot(beforeChart.slice(beforeChart.lastIndexOf('<c:spPr>'))),
		space: slot(xml.slice(split)),
	}
}

const SERIES = [{ name: 'S1', labels: ['A', 'B'], values: [1, 2] }]
const BASE = { x: 1, y: 1, w: 6, h: 3, type: ChartType.bar }

async function fillsFor(opts) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(SERIES, { ...BASE, ...opts })
	})
	return areaFills(await chartXml(zip))
}

const GRADIENT = {
	kind: 'linear',
	angle: 90,
	stops: [
		{ position: 0, color: '0088CC' },
		{ position: 100, color: 'FFFFFF' },
	],
}

defineRegressionSuite('Chart area and plot area fills', [
	{
		// The default, and the case a presence-gated fix would have broken: `plotArea.fill` is `{}`
		// by the time the emitter sees it, and that has to keep meaning transparent.
		name: 'an unstated fill leaves both areas transparent',
		fn: async () => {
			for (const [label, opts] of [
				['omitted entirely', {}],
				['an empty fill object', { plotArea: { fill: {} }, chartArea: { fill: {} } }],
				[
					'transparency with no colour',
					{ plotArea: { fill: { transparency: 50 } }, chartArea: { fill: { transparency: 50 } } },
				],
			]) {
				const { plot, space } = await fillsFor(opts)
				assertEqual(plot, '<a:noFill/>', `plot area fill with ${label}`)
				assertEqual(space, '<a:noFill/>', `chart area fill with ${label}`)
			}
		},
	},
	{
		// The colour path, unchanged, present so the cases below are read against a working baseline
		// rather than against nothing.
		name: 'a colour still reaches both areas as a solid fill',
		fn: async () => {
			const { plot, space } = await fillsFor({
				plotArea: { fill: { color: 'EEEEEE' } },
				chartArea: { fill: { color: '003366', transparency: 25 } },
			})
			assertEqual(plot, '<a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill>', 'plot area solid fill')
			assertEqual(
				space,
				'<a:solidFill><a:srgbClr val="003366"><a:alpha val="75000"/></a:srgbClr></a:solidFill>',
				'chart area solid fill, with the transparency it now has a colour to apply to'
			)
		},
	},
	{
		// The bug proper. Both of these used to emit `<a:noFill/>`.
		name: 'gradient and pattern fills reach the areas they were asked for',
		fn: async () => {
			const { plot, space } = await fillsFor({
				chartArea: { fill: { type: 'gradient', gradient: GRADIENT } },
				plotArea: { fill: { type: 'pattern', pattern: { preset: 'pct25', fgColor: '336699', bgColor: 'FFFFFF' } } },
			})
			assertEqual(
				space,
				'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="0088CC"/></a:gs>' +
					'<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>',
				'chart area gradient fill'
			)
			assertEqual(
				plot,
				'<a:pattFill prst="pct25"><a:fgClr><a:srgbClr val="336699"/></a:fgClr>' +
					'<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
				'plot area pattern fill'
			)

			// Each option reaches only its own area: the other one keeps the default.
			assertEqual(plot.includes('gradFill'), false, 'the chart area gradient did not leak into the plot area')
			assertEqual(space.includes('pattFill'), false, 'the plot area pattern did not leak into the chart area')
		},
	},
	{
		// `'none'` and `'inherit'` are the pair #10 introduced, and a chart is the one surface where
		// they can be told apart without the shape path's "omission means no-fill" default in the way.
		name: "type 'none' paints an explicit no-fill, type 'inherit' emits no fill child at all",
		fn: async () => {
			const none = await fillsFor({ chartArea: { fill: { type: 'none' } }, plotArea: { fill: { type: 'none' } } })
			assertEqual(none.plot, '<a:noFill/>', "plot area with type 'none'")
			assertEqual(none.space, '<a:noFill/>', "chart area with type 'none'")

			const inherit = await fillsFor({
				chartArea: { fill: { type: 'inherit' } },
				plotArea: { fill: { type: 'inherit' } },
			})
			assertEqual(inherit.plot, '', "plot area with type 'inherit' emits nothing into the fill slot")
			assertEqual(inherit.space, '', "chart area with type 'inherit' emits nothing into the fill slot")
		},
	},
	{
		// The one kind that cannot work here. A blip fill needs a media relationship on the chart
		// part and nothing registers one -- `registerImageFillMedia` only runs for shape and slide
		// fills. The pixels are the same as before the fix; what is new is that you are told.
		name: "type 'image' is not supported on a chart and says so instead of failing silently",
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				fillsFor({ chartArea: { fill: { type: 'image', image: { path: 'demos/media/starlabs.png' } } } })
			)
			assertEqual(result.space, '<a:noFill/>', 'the chart area falls back to no fill')
			assert(
				codes.includes('image-fill/unresolved-media'),
				'expected an image-fill/unresolved-media diagnostic; got: ' + JSON.stringify(codes)
			)
		},
	},
])
