import { ChartType } from '../../dist/node.js'
import {
	defineRegressionSuite,
	build,
	readEntry,
	listEntries,
	assert,
	assertEqual,
	assertIncludes,
	assertNotIncludes,
} from '../helpers.js'

// The normalization half of `gen/define/chart.ts` -- everything `addChartDefinition` does to the
// options bag before a byte of chart XML exists: the defensive copy, the enum corrections keyed to
// chart type, the bounded-integer clamps, the plotArea/chartArea/dataBorder defaults, and the combo
// pass that re-runs all of it per subchart. `chart-option-validation.test.js` covers the marker-size
// and gridLine scrubs; this covers the rest of the same surface.
//
// Assertions go through the emitted chart part rather than the options object, because the options
// object is internal -- the only contract a caller can observe is which attributes reach the part.
// Where a correction is *not* visible in the bytes (a dropped `dataLabelPosition` on a chart type
// that has no `<c:dLblPos>` to begin with) the case asserts the absence instead, and a sibling case
// in the same block shows the same attribute present, so the absence is a real signal.
//
// Left deliberately red, all "unreachable by construction" in the sense of docs/testing.md. They
// fall into three groups:
//
//   1. THIRTEEN self-cancelling ternaries, of the form
//        `options.showLegend = options.showLegend || !options.showLegend ? options.showLegend : false`
//      The condition `a || !a` is true for every value of `a`, so the alternative is dead and the
//      whole statement is an identity assignment. Eleven of them are the `show*` block in
//      `normalizeChartPlotAreaOptions` (L174-187), plus `v3DRAngAx` (L200) and `dataLabelBkgrdColors`
//      (L516). No input reaches the `: false` / `: true` arms because none exists. Worth collapsing
//      one day, but that is a `src` edit and so gated on a byte-identity baseline.
//   2. Combo fallbacks whose input the chart-level pass has already filled in. By the time
//      `normalizeComboSubchartOptions` runs, `barDir`, `bar3DShape`, `lineDataSymbol`,
//      `barGapWidthPct`, `barGapDepthPct` and `lineDataSymbolSize` all carry a normalized
//      chart-level value, and the merged bag inherits it -- so `fixed.barDir || ''` (L361, L363,
//      L365), the `?? chartOptions.barGap*Pct` fallbacks (L377, L378) and the `!= null` guard's else
//      (L383) have no reachable input. Same reason `options.barGrouping || ''` (L136, L139) cannot
//      be empty: `normalizeChartBarGrouping` ran first and a bar chart always leaves with one.
//   3. Options-bag guards the builder makes unreachable. `SlideBuilder.addChart` resolves its
//      options to `arg2 ?? {}`, so the object it hands down is never missing -- which leaves
//      `tmpOpt = ... : opt` (L481) and `copyChartOptions(tmpOpt && ... ? tmpOpt : {})` (L487)
//      without an input. The sibling guards on `data` (L483) and on a combo entry's own
//      `data`/`options` ARE reachable from untyped JS and are covered below.
//
// One shape that looks unreachable and is not, since it cost a first draft of this header: the
// second arm of `!border.width || isNaN(border.width)`. `NaN` is falsy, so it seems the first arm
// must always catch it -- but a branch counter records that the second operand was *evaluated*, not
// that it was true, and any truthy width gets there. Both instances (plotArea and dataBorder) are
// covered by passing a width rather than omitting one.

function chartXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, path)
}

/** The chartEx part, for the chart types that emit one (waterfall, funnel, ...). */
function chartExXml(zip) {
	const path = listEntries(zip).find((p) => /^ppt\/charts\/chartEx\d+\.xml$/.test(p))
	assert(path, 'expected a ppt/charts/chartExN.xml entry; got: ' + JSON.stringify(listEntries(zip)))
	return readEntry(zip, path)
}

/** Build one chart and return its part. */
async function chartFrom(data, options) {
	const { zip } = await build((p) => {
		p.addSlide().addChart(data, options)
	})
	return chartXml(zip)
}

/** Build, capturing library warnings (`log.ts` routes every one through `console.warn`). */
async function buildCapturingWarnings(buildFn) {
	const warnings = []
	const original = console.warn
	console.warn = (message) => warnings.push(String(message))
	try {
		const result = await build(buildFn)
		return { ...result, warnings }
	} finally {
		console.warn = original
	}
}

const SERIES = [{ name: 'S1', labels: ['A', 'B', 'C'], values: [1, 2, 3] }]
const BASE = { x: 1, y: 1, w: 6, h: 3 }

defineRegressionSuite('Chart definition', [
	{
		// `copyChartOptions` is what lets `addChart` treat the caller's bag as read-only: every nested
		// object a normalizer writes into is shallow-copied first. The cheap check is that the caller's
		// objects come back untouched after a build that would otherwise have defaulted them --
		// `chartArea.border` gains a color and width, `plotArea.fill` and the gridlines get scrubbed,
		// and `shadow` is normalized in place by `correctShadowOptions`.
		//
		// `plotArea` here has a `fill` but no `border`, which is also the only shape that reaches the
		// else of the border copy.
		name: 'every nested option bag is copied before normalization touches it',
		fn: async () => {
			const plotArea = { fill: { color: 'EEEEEE' } }
			const chartArea = { border: { color: 'FF0000' } }
			const dataBorder = { color: '00FF00' }
			const layout = { x: 0.1, y: 0.1, w: 0.8, h: 0.8 }
			const serGridLine = { color: 'CCCCCC', size: 1 }
			const shadow = { type: 'outer', angle: 45, opacity: 0.5, blur: 3 }
			const before = JSON.stringify({ plotArea, chartArea, dataBorder, layout, serGridLine, shadow })
			await build((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.bar3d,
					plotArea,
					chartArea,
					dataBorder,
					layout,
					serGridLine,
					shadow: /** @type {any} */ (shadow),
				})
			})
			assertEqual(
				JSON.stringify({ plotArea, chartArea, dataBorder, layout, serGridLine, shadow }),
				before,
				"the caller's nested option objects"
			)
		},
	},
	{
		// `<c:dLblPos>` is not valid on every plot type. Area, 3D bar, doughnut and radar have no
		// legal position at all, so the option is dropped outright rather than emitted and rejected.
		// The bar sibling shows the same option surviving where it IS legal.
		name: 'dataLabelPosition is dropped entirely for the plot types that have no legal position',
		fn: async () => {
			for (const type of [ChartType.area, ChartType.bar3d, ChartType.doughnut, ChartType.radar]) {
				const xml = await chartFrom(SERIES, { ...BASE, type, showValue: true, dataLabelPosition: 'ctr' })
				assertNotIncludes(xml, '<c:dLblPos', `dLblPos on a ${type} chart`)
			}
			const bar = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar,
				showValue: true,
				dataLabelPosition: 'ctr',
			})
			assertIncludes(bar, '<c:dLblPos val="ctr"/>', 'dLblPos on a bar chart')
		},
	},
	{
		// Pie accepts four positions; anything else (here a bar-only one) is dropped, and the emitter
		// then falls back to its own `ctr` default rather than leaving the element out -- so the
		// evidence that the drop happened is the substituted value, not an absence.
		name: 'a pie chart keeps its four legal label positions and drops the rest',
		fn: async () => {
			const kept = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.pie,
				showValue: true,
				dataLabelPosition: 'inEnd',
			})
			assertIncludes(kept, '<c:dLblPos val="inEnd"/>', 'a legal pie position')

			const dropped = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.pie,
				showValue: true,
				dataLabelPosition: 'inBase',
			})
			assertNotIncludes(dropped, 'inBase', 'a bar-only position on a pie chart')
			assertIncludes(dropped, '<c:dLblPos val="ctr"/>', "the emitter's own default takes over")
		},
	},
	{
		// Bar label positions are filtered by the grouping, and the two rules run in sequence rather
		// than as a switch -- which makes the outcome the opposite of the obvious reading. The first
		// rule applies to any grouping that is NOT stacked and allows only ctr/inBase/inEnd; the
		// second applies to any grouping that is NOT clustered and additionally allows outEnd. So
		// `outEnd`, whose name suggests a clustered bar's outside-end label, is dropped on clustered
		// bars by the first rule and kept on stacked ones by the second. A position outside both
		// lists only ever meets the second rule, since a stacked chart skips the first.
		name: 'bar label positions are filtered by the bar grouping, and outEnd inverts',
		fn: async () => {
			const barLabel = (barGrouping, dataLabelPosition) =>
				chartFrom(SERIES, { ...BASE, type: ChartType.bar, showValue: true, barGrouping, dataLabelPosition })

			assertIncludes(await barLabel('clustered', 'inEnd'), '<c:dLblPos val="inEnd"/>', 'inEnd on clustered bars')
			assertNotIncludes(await barLabel('clustered', 'outEnd'), '<c:dLblPos', 'outEnd on clustered bars')
			assertIncludes(await barLabel('stacked', 'outEnd'), '<c:dLblPos val="outEnd"/>', 'outEnd on stacked bars')
			assertNotIncludes(await barLabel('stacked', /** @type {any} */ ('b')), '<c:dLblPos', 'a line-only position')
		},
	},
	{
		// `<c:grouping>` is ST_Grouping, and which members it accepts depends on the plot type: area
		// takes no `clustered`, 3D bar does. A value already legal for the type passes through; an
		// illegal one is replaced by that type's default rather than emitted.
		name: 'barGrouping is corrected per plot type and left alone when already legal',
		fn: async () => {
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.area, barGrouping: 'percentStacked' }),
				'<c:grouping val="percentStacked"/>',
				'a legal area grouping'
			)
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.area, barGrouping: /** @type {any} */ ('clustered') }),
				'<c:grouping val="standard"/>',
				'an illegal area grouping falls back to standard'
			)
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.bar3d, barGrouping: 'clustered' }),
				'<c:grouping val="clustered"/>',
				'a legal 3D bar grouping'
			)
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.bar3d, barGrouping: /** @type {any} */ ('sideways') }),
				'<c:grouping val="standard"/>',
				'an illegal 3D bar grouping falls back to standard'
			)
		},
	},
	{
		// The three axis-line toggles default to "shown" and are only consulted with a `typeof`
		// check, so `false` has to be distinguishable from omitted. Turning all three off is the
		// only way to prove the check reads the value rather than its truthiness.
		name: 'axis line visibility is read as a value, so false suppresses the line',
		fn: async () => {
			const shown = await chartFrom(SERIES, { ...BASE, type: ChartType.bar3d })
			const hidden = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar3d,
				catAxisLineShow: false,
				valAxisLineShow: false,
				serAxisLineShow: false,
			})
			// The `<a:ln>` is emitted either way; what changes is its fill, so the marker is a
			// `<a:noFill/>` where the gridline color would be. Scoped to each axis element, because
			// the val axis also carries its major gridlines' own `<a:ln>` and both would match a
			// whole-document search.
			const axis = (xml, tag) => {
				const block = (xml.match(new RegExp(`<c:${tag}>[\\s\\S]*?</c:${tag}>`)) || [])[0]
				assert(block, `expected a <c:${tag}> in the chart part`)
				return block
			}
			for (const tag of ['catAx', 'valAx', 'serAx']) {
				assertNotIncludes(axis(shown, tag), 'cap="flat"><a:noFill/>', `${tag} is drawn by default`)
				assertIncludes(axis(hidden, tag), 'cap="flat"><a:noFill/>', `${tag} with its line turned off`)
			}
		},
	},
	{
		// The 3D view angles are bounded (`rotX` -90..90, `rotY` 0..360, `perspective` 0..240) and,
		// unlike the `clampChartPct` options, an out-of-range value is silently replaced by the
		// default rather than clamped to the nearest bound -- so 200 does not become 240.
		name: '3D view angles are kept when in range and replaced by the default when not',
		fn: async () => {
			const inRange = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar3d,
				v3DRotX: 45,
				v3DRotY: 340,
				v3DPerspective: 100,
			})
			assertIncludes(inRange, '<c:rotX val="45"/>', 'an in-range rotX')
			assertIncludes(inRange, '<c:rotY val="340"/>', 'an in-range rotY')
			assertIncludes(inRange, '<c:perspective val="100"/>', 'an in-range perspective')

			const outOfRange = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar3d,
				v3DRotX: -100,
				v3DRotY: 400,
				v3DPerspective: 999,
			})
			assertIncludes(outOfRange, '<c:rotX val="30"/>', 'an out-of-range rotX falls back to 30')
			assertIncludes(outOfRange, '<c:rotY val="30"/>', 'an out-of-range rotY falls back to 30')
			assertIncludes(outOfRange, '<c:perspective val="30"/>', 'an out-of-range perspective falls back to 30')
		},
	},
	{
		// `chartColorsOpacity` is only honoured when it is a usable number; it reaches the series fill
		// as an alpha in thousandths of a percent.
		name: 'chartColorsOpacity reaches the series fill as an alpha',
		fn: async () => {
			const xml = await chartFrom(SERIES, { ...BASE, type: ChartType.bar, chartColorsOpacity: 50 })
			assertIncludes(xml, '<a:alpha val="50000"/>', 'the series fill alpha')
			assertNotIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.bar }),
				'<a:alpha',
				'a chart with no opacity set'
			)
		},
	},
	{
		// plotArea and chartArea borders are defaulted independently: a plotArea border with no color
		// gets the shared default color, and a chartArea border is rebuilt from scratch so that a
		// partial one comes out complete. The two chartArea cases differ in which half was supplied.
		name: 'plotArea and chartArea borders are completed from partial input',
		fn: async () => {
			const plot = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar,
				plotArea: { border: /** @type {any} */ ({ width: 2 }) },
			})
			assertIncludes(plot, '<a:ln w="25400" cap="flat">', 'the supplied plotArea border width, in EMU')
			assertIncludes(plot, '<a:srgbClr val="363636"/>', 'the defaulted plotArea border color')

			// The chartArea border is rebuilt rather than patched, so each half is defaulted
			// independently: a color-only border gets the default 1pt width, a width-only border gets
			// the default color. The full `<a:ln>` pins both halves of each.
			const colorOnly = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar,
				chartArea: { border: /** @type {any} */ ({ color: 'FF0000' }) },
			})
			assertIncludes(
				colorOnly,
				'<a:ln w="12700" cap="flat"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
				'a color-only chartArea border'
			)

			const widthOnly = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar,
				chartArea: { border: /** @type {any} */ ({ width: 3 }) },
			})
			assertIncludes(
				widthOnly,
				'<a:ln w="38100" cap="flat"><a:solidFill><a:srgbClr val="363636"/></a:solidFill></a:ln>',
				'a width-only chartArea border'
			)
		},
	},
	{
		// `roundedCorners` defaults to true and is read with a `typeof` check, so an explicit false is
		// the only input that distinguishes "off" from "not set".
		name: 'chartArea roundedCorners: false is honoured',
		fn: async () => {
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.bar, chartArea: { roundedCorners: false } }),
				'<c:roundedCorners val="0"/>',
				'rounded corners off'
			)
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.bar }),
				'<c:roundedCorners val="1"/>',
				'the default'
			)
		},
	},
	{
		// A data border color has to be either a six-digit hex or a scheme color name; anything else
		// would reach `<a:srgbClr val>` verbatim and make PowerPoint report the part as damaged, so it
		// falls back to F9F9F9. Six characters that are not hex is the case the length test alone
		// would let through, which is why the regex exists.
		// A width is supplied on both so the `!width || isNaN(width)` guard actually reaches its second
		// arm -- with the width omitted, the first arm short-circuits and the isNaN check never runs.
		name: 'a dataBorder color that is neither hex nor scheme falls back',
		fn: async () => {
			const hex = await chartFrom(SERIES, { ...BASE, type: ChartType.bar, dataBorder: { color: 'FF0000', width: 2 } })
			assertIncludes(hex, '<a:srgbClr val="FF0000"/>', 'a valid six-digit hex')
			assertIncludes(hex, '<a:ln w="25400"', 'the supplied dataBorder width, in EMU')

			const notHex = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar,
				dataBorder: { color: /** @type {any} */ ('ZZZZZZ'), width: 2 },
			})
			assertIncludes(notHex, '<a:srgbClr val="F9F9F9"/>', 'six non-hex characters')
			assertNotIncludes(notHex, 'ZZZZZZ', 'the rejected color must not reach the part')

			// A scheme color name is the other accepted spelling, and passes through as a schemeClr.
			const scheme = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar,
				dataBorder: { color: /** @type {any} */ ('accent2'), width: 2 },
			})
			assertIncludes(scheme, '<a:schemeClr val="accent2"/>', 'a scheme color name')
		},
	},
	{
		// `<cx:subtotals>` holds zero-based category indices, so a fractional or negative entry would
		// make PowerPoint report the chartEx part as damaged. Bad entries are warned about and
		// skipped; if that leaves nothing, the whole element is omitted rather than emitted empty.
		name: 'waterfall subtotals drop invalid indices, and an all-invalid list is omitted',
		fn: async () => {
			const partial = await buildCapturingWarnings((p) => {
				p.addSlide().addChart(SERIES, {
					...BASE,
					type: ChartType.waterfall,
					subtotals: /** @type {any} */ ([1, -2, 1.5, 'x']),
				})
			})
			assertEqual(
				partial.warnings.filter((message) => /is not a non-negative integer/.test(message)).length,
				3,
				`expected one warning per skipped index; got ${JSON.stringify(partial.warnings)}`
			)
			assertIncludes(await chartExXml(partial.zip), '<cx:subtotals>', 'a list with one survivor')

			const none = await buildCapturingWarnings((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.waterfall, subtotals: /** @type {any} */ ([-1]) })
			})
			assertNotIncludes(await chartExXml(none.zip), '<cx:subtotals>', 'a list with no survivors')

			// Not an array at all: nothing to keep, nothing to warn about.
			const notAList = await buildCapturingWarnings((p) => {
				p.addSlide().addChart(SERIES, { ...BASE, type: ChartType.waterfall, subtotals: /** @type {any} */ (7) })
			})
			assertNotIncludes(await chartExXml(notAList.zip), '<cx:subtotals>', 'a non-array subtotals')
		},
	},
	{
		// A stacked bar chart wants a narrower gap than the clustered default, but only when the
		// caller did not ask for one. Two decks, same grouping, different `barGapWidthPct` provenance.
		name: 'a stacked bar chart takes the narrow default gap unless the caller set one',
		fn: async () => {
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.bar, barGrouping: 'stacked' }),
				'<c:gapWidth val="50"/>',
				'the stacked default gap'
			)
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.bar, barGrouping: 'stacked', barGapWidthPct: 300 }),
				'<c:gapWidth val="300"/>',
				"the caller's gap survives"
			)
		},
	},
	{
		// Three ST_ enumerations that are replaced wholesale when the value is not a member. Each is
		// checked twice: a legal value passes through, an illegal one becomes the default.
		name: 'bar3DShape, lineDataSymbol and displayBlanksAs pass through when legal',
		fn: async () => {
			const legal = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar3d,
				bar3DShape: 'cylinder',
				displayBlanksAs: 'zero',
			})
			assertIncludes(legal, '<c:shape val="cylinder"/>', 'a legal ST_Shape')
			assertIncludes(legal, '<c:dispBlanksAs val="zero"/>', 'a legal ST_DispBlanksAs')

			const illegal = await chartFrom(SERIES, {
				...BASE,
				type: ChartType.bar3d,
				bar3DShape: /** @type {any} */ ('sphere'),
				displayBlanksAs: /** @type {any} */ ('hide'),
			})
			assertIncludes(illegal, '<c:shape val="box"/>', 'an illegal ST_Shape falls back to box')
			assertIncludes(illegal, '<c:dispBlanksAs val="gap"/>', 'an illegal ST_DispBlanksAs falls back to gap')

			const symbol = await chartFrom(SERIES, { ...BASE, type: ChartType.line, lineDataSymbol: 'diamond' })
			assertIncludes(symbol, '<c:symbol val="diamond"/>', 'a legal ST_MarkerStyle')
		},
	},
	{
		// `lineDataSymbolLineSize` is given in points and emitted in EMU, so the conversion is the
		// observable. Omitted, it takes the same conversion applied to 0.75pt.
		name: 'lineDataSymbolLineSize converts points to EMU, defaulting to 0.75pt',
		fn: async () => {
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.line, lineDataSymbol: 'circle', lineDataSymbolLineSize: 4 }),
				'w="50800"',
				'4pt in EMU'
			)
			assertIncludes(
				await chartFrom(SERIES, { ...BASE, type: ChartType.line, lineDataSymbol: 'circle' }),
				'w="9525"',
				'the 0.75pt default in EMU'
			)
		},
	},
	{
		// Stock charts plot a fixed number of series per style (hlc: 3, ohlc: 4, vhlc: 4, vohlc: 5).
		// A mismatch is warned about rather than corrected -- there is no safe way to invent or drop a
		// series -- and an unrecognized style falls back to hlc first, so the count is judged against
		// the style that will actually be emitted.
		name: 'a stock chart warns when its series count does not match its style',
		fn: async () => {
			const three = [1, 2, 3].map((n) => ({ name: `S${n}`, labels: ['A', 'B'], values: [n, n + 1] }))
			const ok = await buildCapturingWarnings((p) => {
				p.addSlide().addChart(three, { ...BASE, type: ChartType.stock, stockStyle: 'hlc' })
			})
			assertEqual(
				ok.warnings.filter((message) => /expects \d+ data series/.test(message)).length,
				0,
				`a matching count must not warn; got ${JSON.stringify(ok.warnings)}`
			)
			const mismatch = await buildCapturingWarnings((p) => {
				p.addSlide().addChart(three, { ...BASE, type: ChartType.stock, stockStyle: 'vohlc' })
			})
			assertEqual(
				mismatch.warnings.filter((message) => /expects 5 data series \(got 3\)/.test(message)).length,
				1,
				`expected a series-count warning; got ${JSON.stringify(mismatch.warnings)}`
			)
		},
	},
	{
		// Combo subchart overrides are merged over the chart-level options at emit time, so they get
		// the same enum corrections and unit conversions -- keyed to the subchart's own type, not the
		// chart's. Without that pass an illegal `bar3DShape`/`lineDataSymbol` would reach the part
		// verbatim and a `lineDataSymbolLineSize` would be emitted in points where EMU are expected.
		name: 'combo subchart overrides go through the same corrections as chart-level options',
		fn: async () => {
			const xml = await chartFrom(
				/** @type {any} */ ([
					{
						type: ChartType.bar,
						data: [{ name: 'Bars', labels: ['A', 'B'], values: [1, 2] }],
						options: { barGrouping: 'stacked', bar3DShape: 'sphere' },
					},
					{
						type: ChartType.line,
						data: [{ name: 'Line', labels: ['A', 'B'], values: [3, 4] }],
						options: { lineDataSymbol: 'sphere', lineDataSymbolLineSize: 4 },
					},
				]),
				{ ...BASE }
			)
			assertNotIncludes(xml, 'sphere', 'no illegal enum value reaches the part')
			assertIncludes(xml, '<c:symbol val="circle"/>', 'the corrected marker symbol')
			assertIncludes(xml, 'w="50800"', "the subchart's 4pt marker line, in EMU")
		},
	},
	{
		// Only the FIRST entry has to look like a `ChartMulti` for the builder to route to the combo
		// path, so the later entries' guards are reachable from untyped JS: an entry with no `data`
		// contributes no series rather than throwing, and one with no `options` is treated as an empty
		// override bag. The sibling guard on non-array `data` a few lines down in the same function is
		// reachable the same way.
		name: 'combo entries missing data or options degrade instead of throwing',
		fn: async () => {
			const combo = await chartFrom(
				/** @type {any} */ ([
					{ type: ChartType.bar, data: [{ name: 'Bars', labels: ['A', 'B'], values: [1, 2] }], options: {} },
					{ type: ChartType.line },
				]),
				{ ...BASE }
			)
			assertEqual((combo.match(/<c:ser>/g) || []).length, 1, 'only the entry that had data contributes a series')

			const notAList = await chartFrom(/** @type {any} */ ({}), { ...BASE, type: ChartType.bar })
			assertEqual((notAList.match(/<c:ser>/g) || []).length, 0, 'non-array data plots nothing')
		},
	},
])
