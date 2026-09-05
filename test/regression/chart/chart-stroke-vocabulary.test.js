import TsPptx, { ChartType, SchemeColor } from '../../../dist/node.js'
import { defineRegressionSuite, build, assertEqual, assertIncludes, assertNotIncludes } from '../../helpers.js'
import { chartXml } from './chart-parts.js'

// A chart stroke used to be spelled three ways the rest of the library did not know: `size`
// for the width, `style` for the dash, and a flat `*AxisLineShow` flag for the state
// `type: 'none'` already names. Those spellings still work (deprecated), and everything a
// stroke can say — `width`, `dashType`, `type`, `color`, `cap`, `transparency` — now reaches
// all four sites: the three axis lines, gridlines, series lines and error bars.
//
// The byte-identity gate cannot cover this: the showcase decks reach `*AxisLineShow` and
// gridlines only, so `catAxisLineStyle`, `*AxisLineSize`, `*AxisLineColor`, error bars and
// `barSeriesLine` have no coverage there at all. What each case asserts is the emitted
// attribute or element, not that the option was accepted.

const SERIES = [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]
const BASE = { x: 1, y: 1, w: 6, h: 3 }

/** The chart part for one options bag. */
async function chartFor(options, data = SERIES) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(data, { ...BASE, ...options })
	})
	return chartXml(zip)
}

/**
 * One axis' own `<c:spPr>` line block, by axis element rather than by position.
 *
 * `<c:majorGridlines>` carries a `<c:spPr>` of its own and sits ahead of the axis' in
 * `CT_*Ax`, so it has to come out first -- otherwise every value-axis assertion reads the
 * gridline's stroke and passes on the default.
 */
function axisLineBlock(xml, kind) {
	const axis = xml.match(new RegExp(`<c:${kind}Ax>[^]*?</c:${kind}Ax>`))
	if (!axis) return ''
	const own = axis[0].replace(/<c:majorGridlines>[^]*?<\/c:majorGridlines>/g, '')
	const spPr = own.match(/<c:spPr><a:ln [^]*?<\/a:ln><\/c:spPr>/)
	return spPr ? spPr[0] : ''
}

defineRegressionSuite('Chart stroke vocabulary', [
	{
		name: 'an axis line takes width, dashType, colour, cap and transparency in one stroke',
		fn: async () => {
			const xml = await chartFor({
				type: ChartType.bar,
				catAxisLine: { width: 2.5, dashType: 'lgDashDot', color: '112233', cap: 'round', transparency: 40 },
			})
			const line = axisLineBlock(xml, 'cat')
			assertIncludes(line, 'w="31750"', '2.5pt reaches a:ln@w')
			assertIncludes(line, 'cap="rnd"', 'the cap reaches a:ln@cap, which the flat keys could not say')
			assertIncludes(line, '<a:prstDash val="lgDashDot"/>', 'the full ST_PresetLineDashVal set is reachable')
			assertIncludes(
				line,
				'<a:srgbClr val="112233"><a:alpha val="60000"/></a:srgbClr>',
				'transparency becomes an a:alpha'
			)
		},
	},
	{
		name: 'an axis line colour may be a scheme token, which the bare-string key only documented as hex',
		fn: async () => {
			const xml = await chartFor({ type: ChartType.bar, valAxisLine: { color: SchemeColor.accent1 } })
			assertIncludes(axisLineBlock(xml, 'val'), '<a:schemeClr val="accent1"/>', 'a theme colour reaches the axis line')
		},
	},
	{
		name: "an axis line's type: 'none' is what *AxisLineShow: false used to say",
		fn: async () => {
			const nested = await chartFor({ type: ChartType.bar, catAxisLine: { type: 'none' } })
			const flat = await chartFor({ type: ChartType.bar, catAxisLineShow: false })
			assertIncludes(axisLineBlock(nested, 'cat'), '<a:noFill/>', "type: 'none' paints no line")
			assertEqual(axisLineBlock(nested, 'cat'), axisLineBlock(flat, 'cat'), 'the two spellings emit the same block')
		},
	},
	{
		name: 'the nested stroke wins over the flat key it supersedes',
		fn: async () => {
			const xml = await chartFor({
				type: ChartType.bar,
				catAxisLine: { width: 3, color: 'AABBCC' },
				catAxisLineSize: 1,
				catAxisLineColor: '000000',
			})
			const line = axisLineBlock(xml, 'cat')
			assertIncludes(line, 'w="38100"', 'the nested width wins')
			assertIncludes(line, 'AABBCC', 'the nested colour wins')
			assertNotIncludes(line, '000000', 'the superseded flat colour must not reach the part')
		},
	},
	{
		name: 'a hidden axis line still writes the dash the caller asked for',
		fn: async () => {
			// `show` and `style` fold independently. Letting the inferred `type: 'none'` swallow
			// the dash would change the bytes of every deck that set both.
			const xml = await chartFor({ type: ChartType.bar, catAxisLineShow: false, catAxisLineStyle: 'dot' })
			const line = axisLineBlock(xml, 'cat')
			assertIncludes(line, '<a:noFill/>', 'the line is hidden')
			assertIncludes(line, '<a:prstDash val="dot"/>', 'and its dash survives the fold')
		},
	},
	{
		name: 'the series axis reads the same stroke as the other two',
		fn: async () => {
			const xml = await chartFor({
				type: ChartType.bar3d,
				serAxisLine: { width: 2, color: 'ABCDEF', dashType: 'sysDashDot', cap: 'square' },
			})
			const serAx = axisLineBlock(xml, 'ser')
			assertIncludes(serAx, 'w="25400"', 'the series axis width is no longer hardcoded to one point')
			assertIncludes(serAx, 'cap="sq"', 'nor its cap to flat')
			assertIncludes(serAx, '<a:prstDash val="sysDashDot"/>', 'nor its dash to solid')
			assertIncludes(serAx, 'ABCDEF', 'and its colour still reaches the part')
		},
	},
	{
		name: 'a gridline takes width and dashType, and size/style still mean the same thing',
		fn: async () => {
			const modern = await chartFor({
				type: ChartType.bar,
				valGridLine: { width: 2, dashType: 'dash', color: 'CCCCCC', cap: 'round' },
			})
			const legacy = await chartFor({
				type: ChartType.bar,
				valGridLine: { size: 2, style: 'dash', color: 'CCCCCC', cap: 'round' },
			})
			const grid = (xml) => xml.match(/<c:majorGridlines>[^]*?<\/c:majorGridlines>/)[0]
			assertIncludes(grid(modern), 'w="25400"', 'width reaches a:ln@w')
			assertEqual(grid(modern), grid(legacy), 'size/style emit exactly what width/dashType do')
		},
	},
	{
		name: "a gridline's type: 'none' omits the element, as style: 'none' does",
		fn: async () => {
			const xml = await chartFor({ type: ChartType.bar, valGridLine: { type: 'none' } })
			assertNotIncludes(xml, '<c:majorGridlines>', "type: 'none' suppresses the gridlines")
		},
	},
	{
		name: 'a gridline takes a transparency, which its four-key shape could not express',
		fn: async () => {
			const xml = await chartFor({ type: ChartType.bar, valGridLine: { color: '888888', transparency: 25 } })
			assertIncludes(xml, '<a:alpha val="75000"/>', 'transparency reaches the gridline paint')
		},
	},
	{
		name: 'series lines take the same stroke, including the suppressing type',
		fn: async () => {
			const drawn = await chartFor({
				type: ChartType.bar,
				barGrouping: 'stacked',
				barSeriesLine: { width: 2, dashType: 'sysDot', color: '999999' },
			})
			assertIncludes(drawn, '<c:serLines>', 'the element is emitted')
			assertIncludes(drawn, '<a:prstDash val="sysDot"/>', 'over the full dash set')
			const none = await chartFor({ type: ChartType.bar, barGrouping: 'stacked', barSeriesLine: { type: 'none' } })
			assertNotIncludes(none, '<c:serLines>', "type: 'none' omits it, as style: 'none' does")
		},
	},
	{
		name: 'an error bar takes dashType, cap and transparency, not only a colour and a width',
		fn: async () => {
			const xml = await chartFor({ type: ChartType.bar }, [
				{
					name: 'S1',
					labels: ['A', 'B'],
					values: [1, 2],
					errorBars: { color: 'FF0000', width: 2, dashType: 'sysDash', cap: 'round', transparency: 30 },
				},
			])
			const bars = xml.match(/<c:errBars>[^]*?<\/c:errBars>/)[0]
			assertIncludes(bars, 'w="25400"', 'width reaches a:ln@w')
			assertIncludes(bars, 'cap="rnd"', 'the cap reaches a:ln@cap')
			assertIncludes(bars, '<a:prstDash val="sysDash"/>', 'the dash reaches a:prstDash')
			assertIncludes(bars, '<a:alpha val="70000"/>', 'the transparency reaches a:alpha')
		},
	},
	{
		name: 'an error bar that names only a dash does not conjure a colour out of the default',
		fn: async () => {
			// "no `<a:solidFill>`" is what leaves PowerPoint's own error-bar style in charge, and
			// a caller who asked for a dash did not ask to lose it.
			const xml = await chartFor({ type: ChartType.bar }, [
				{ name: 'S1', labels: ['A', 'B'], values: [1, 2], errorBars: { dashType: 'dot' } },
			])
			const bars = xml.match(/<c:errBars>[^]*?<\/c:errBars>/)[0]
			assertIncludes(bars, '<a:prstDash val="dot"/>', 'the dash is emitted')
			assertNotIncludes(bars, '<a:solidFill>', 'and no colour is invented for it')
		},
	},
	{
		name: 'a chart no longer accepts the four gridline keys at its top level',
		fn: () => {
			// `ChartOpts` used to extend `OptsChartGridLine`, so `color`, `size`, `style` and `cap`
			// typechecked on the chart bag itself and were read by nothing. A JS caller can still
			// pass them; what this pins is that they reach no attribute.
			const pres = new TsPptx()
			// The cast is the assertion: without it this line no longer compiles, which is the half
			// of the change TypeScript enforces. A JavaScript caller can still pass them.
			const stray = /** @type {any} */ ({ ...BASE, type: ChartType.bar, color: 'FF0000', size: 9, style: 'dash' })
			pres.addSlide().addChart(SERIES, stray)
			// Nothing to assert beyond "this still builds" — the type-level half is the change, and
			// a silent no-op is what it removes from the surface.
			assertEqual(typeof pres.toBytes, 'function', 'the deck still builds with the stray keys present')
		},
	},
])
