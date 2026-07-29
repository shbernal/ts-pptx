/**
 * ts-pptx: Chart Definition
 *
 * `addChartDefinition` normalizes `addChart()` options onto the slide model and registers the
 * chart part rel; the `normalize*` / `clamp*` helpers apply the schema-valid defaults and range
 * clamps. The chart *XML* is emitted later by `gen/chart/chart-xml.ts`.
 */
import {
	asChartType,
	type CHART_NAME,
	ChartType,
	isChartExType,
	SchemeColor,
	SlideObjectType,
} from '../../core-enums.js'
import { BARCHART_COLORS, DEF_CHART_BORDER, PIECHART_COLORS } from '../../core-enums-internal.js'
import { warn } from '../../diagnostics.js'
import type { ChartMulti, ChartOpts, OptsChartData, OptsChartGridLine } from '../../core-interfaces.js'
import type { ChartOptsInternal, OptsChartDataInternal, PresSlideInternal, SlideObject } from '../../types/internal.js'
import { encodeXmlAttrValue, getNewRelId, validateObjectName } from '../../gen-utils.js'
import { correctShadowOptions } from '../drawingml/effect.js'
import { valToPts } from '../../units-internal.js'

/**
 * Copy one series into the internal shape the emitters read, without touching the caller's object.
 *
 * `labels` is widened to the nested `string[][]` form the multi-level category serializer wants and
 * `_dataIndex` records the series' position. Both used to be assigned back onto the caller's own
 * series object, so `data[0].labels` became `[['A','B','C']]` after `addChart` and any code that
 * reused the same array afterwards (a legend, a table, a second chart) silently saw one nested array
 * instead of three strings.
 *
 * The copy is one level deep on purpose: everything under `src/gen/chart/` is a pure string builder
 * that only reads `values`/`sizes`/`customLabels`/`pointStyles`/`errorBars`, so sharing those arrays
 * with the caller is safe.
 * @param item - caller-supplied series
 * @param index - series position, across all subcharts for a combo chart
 */
function normalizeChartSeries(item: OptsChartData, index: number): OptsChartDataInternal {
	const labels = item.labels
	return {
		...item,
		_dataIndex: index,
		labels: labels === undefined ? undefined : Array.isArray(labels[0]) ? (labels as string[][]) : [labels as string[]],
	}
}

/**
 * Copy a caller-supplied `ChartOpts` so the normalization below never writes back onto it.
 *
 * `addChartDefinition` fills in defaults, clamps out-of-range values and deletes invalid keys
 * (`layout.x`, `catGridLine.size`, `dataLabelPosition`, …). Doing that in place mutated the
 * caller's object, which is both surprising on its own and order-dependent when one options
 * object is shared across two charts.
 *
 * Nested objects are copied wherever a normalizer writes into them. `structuredClone` is
 * deliberately not used: it would deep-copy the `ChartMulti[]` on `_type` — breaking the series
 * identity the combo emit path depends on — and throws on a stray function in an untyped
 * caller's option bag.
 * @param opts - caller-supplied chart options
 */
function copyChartOptions(opts: ChartOpts | ChartOptsInternal): ChartOptsInternal {
	const copy: ChartOptsInternal = { ...opts }
	// `_type` is derived from the `type` argument below; never inherit a caller-stamped one.
	delete copy._type
	if (copy.plotArea) {
		copy.plotArea = { ...copy.plotArea }
		if (copy.plotArea.border) copy.plotArea.border = { ...copy.plotArea.border }
		if (copy.plotArea.fill) copy.plotArea.fill = { ...copy.plotArea.fill }
	}
	if (copy.chartArea) copy.chartArea = { ...copy.chartArea }
	if (copy.dataBorder) copy.dataBorder = { ...copy.dataBorder }
	if (copy.layout) copy.layout = { ...copy.layout }
	if (copy.catGridLine) copy.catGridLine = { ...copy.catGridLine }
	if (copy.valGridLine) copy.valGridLine = { ...copy.valGridLine }
	if (copy.serGridLine) copy.serGridLine = { ...copy.serGridLine }
	// `correctShadowOptions` normalizes its argument in place (angle rounding, `_alpha`).
	if (copy.shadow) copy.shadow = { ...copy.shadow }
	return copy
}

/**
 * Round and clamp an integer chart percentage/angle option into a schema-valid range.
 *
 * Several chart attributes are bounded integer types whose out-of-range values make
 * PowerPoint report the package as needing repair: `<c:overlap>` (ST_Overlap, -100..100),
 * `<c:gapWidth>`/`<c:gapDepth>` (ST_GapAmount, 0..500), `<c:holeSize>` (ST_HoleSize, 10..90)
 * and `<c:firstSliceAng>` (ST_FirstSliceAng, 0..360). Missing/non-numeric input returns
 * `undefined` so the caller can apply its own default; an out-of-range value is clamped
 * and a warning is emitted (per the library's warn-rather-than-degrade policy).
 * @param value - caller-supplied option value
 * @param min - inclusive lower bound
 * @param max - inclusive upper bound
 * @param name - option name, for the warning message
 */
function clampChartPct(value: number | undefined, min: number, max: number, name: string): number | undefined {
	if (typeof value !== 'number' || isNaN(value)) return undefined
	const clamped = Math.min(max, Math.max(min, Math.round(value)))
	if (clamped !== value)
		warn('chart/option-out-of-range', `${name} ${value} is outside the valid range ${min}-${max}; using ${clamped}.`)
	return clamped
}

/**
 * Drop `dataLabelPosition` values that are invalid for the chart type / bar grouping,
 * per the OOXML data-label placement rules, so PowerPoint does not flag the file.
 *
 * `chartType` is passed in rather than read off `options._type` so the combo path can run the
 * same rules per subchart: a combo chart's `_type` is a `ChartMulti[]`, which matches none of
 * the comparisons below (see {@link normalizeComboSubchartOptions}).
 * @param options - options bag to correct in place
 * @param chartType - the plot type these options are emitted for, if known
 */
function normalizeChartDataLabelPosition(options: ChartOptsInternal, chartType: ChartType | undefined): void {
	if (options.dataLabelPosition) {
		const dataLabelPosition = options.dataLabelPosition
		if (
			chartType === ChartType.area ||
			chartType === ChartType.bar3d ||
			chartType === ChartType.doughnut ||
			chartType === ChartType.radar
		) {
			delete options.dataLabelPosition
		}
		if (chartType === ChartType.pie) {
			if (!['bestFit', 'ctr', 'inEnd', 'outEnd'].includes(dataLabelPosition)) delete options.dataLabelPosition
		}
		if (
			chartType === ChartType.bubble ||
			chartType === ChartType.bubble3d ||
			chartType === ChartType.line ||
			chartType === ChartType.scatter
		) {
			if (!['b', 'ctr', 'l', 'r', 't'].includes(dataLabelPosition)) delete options.dataLabelPosition
		}
		if (chartType === ChartType.bar) {
			if (!['stacked', 'percentStacked'].includes(options.barGrouping || '')) {
				if (!['ctr', 'inBase', 'inEnd'].includes(dataLabelPosition)) delete options.dataLabelPosition
			}
			if (!['clustered'].includes(options.barGrouping || '')) {
				if (!['ctr', 'inBase', 'inEnd', 'outEnd'].includes(dataLabelPosition)) delete options.dataLabelPosition
			}
		}
	}
}

/**
 * Correct `barGrouping` to a value `<c:grouping>` (ST_Grouping) accepts for the given chart type.
 *
 * Split out of `addChartDefinition` for the same reason as
 * {@link normalizeChartDataLabelPosition}: the combo path needs it per subchart, keyed to that
 * subchart's own type.
 * @param options - options bag to correct in place
 * @param chartType - the plot type these options are emitted for, if known
 */
function normalizeChartBarGrouping(options: ChartOptsInternal, chartType: ChartType | undefined): void {
	// barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
	if (chartType === ChartType.area) {
		if (!['stacked', 'standard', 'percentStacked'].includes(options.barGrouping || '')) options.barGrouping = 'standard'
	}
	if (chartType === ChartType.bar) {
		if (!['clustered', 'stacked', 'percentStacked'].includes(options.barGrouping || ''))
			options.barGrouping = 'clustered'
	}
	if (chartType === ChartType.bar3d) {
		if (!['clustered', 'stacked', 'standard', 'percentStacked'].includes(options.barGrouping || ''))
			options.barGrouping = 'standard'
	}
}

/**
 * Apply plotArea option defaults: show* toggles, axis-line visibility, and the 3D view angles.
 */
function normalizeChartPlotAreaOptions(options: ChartOptsInternal): void {
	// The eleven `show*` toggles are deliberately NOT defaulted here. Each one used to carry
	// a statement of the form `x = x || !x ? x : <default>`, but `a || !a` is true for every
	// value of `a`, so the alternative never ran and the whole statement was an identity
	// assignment. Every consumer reads these as plain truthiness (`opts.showPercent ? 1 : 0`,
	// `!opts.showDataTableKeys ? 0 : 1`), which makes an absent option behave as `false` — and
	// that is what the public types document (`@default false` on `showDataTable`,
	// `showPercent`, `v3DRAngAx`). Removing the dead statements changes no emitted byte;
	// *applying* the defaults those ternaries appear to promise would, so do not "restore"
	// them without treating it as the behavior change it is.
	options.catAxisLineShow = typeof options.catAxisLineShow !== 'undefined' ? options.catAxisLineShow : true
	options.valAxisLineShow = typeof options.valAxisLineShow !== 'undefined' ? options.valAxisLineShow : true
	options.serAxisLineShow = typeof options.serAxisLineShow !== 'undefined' ? options.serAxisLineShow : true

	options.v3DRotX =
		typeof options.v3DRotX === 'number' && !isNaN(options.v3DRotX) && options.v3DRotX >= -90 && options.v3DRotX <= 90
			? options.v3DRotX
			: 30
	options.v3DRotY =
		typeof options.v3DRotY === 'number' && !isNaN(options.v3DRotY) && options.v3DRotY >= 0 && options.v3DRotY <= 360
			? options.v3DRotY
			: 30
	// v3DRAngAx: same dead-ternary shape as the show* block above, same reason for its absence.
	options.v3DPerspective =
		typeof options.v3DPerspective === 'number' &&
		!isNaN(options.v3DPerspective) &&
		options.v3DPerspective >= 0 &&
		options.v3DPerspective <= 240
			? options.v3DPerspective
			: 30
}

/**
 * Apply chart-level option defaults: gap/overlap/hole clamps, chart colors, plotArea/chartArea
 * borders and fills, data border, data-label format codes, line size and multi-level cat labels.
 */
function normalizeChartOptions(options: ChartOptsInternal): void {
	options.barGapWidthPct = clampChartPct(options.barGapWidthPct, 0, 500, 'barGapWidthPct') ?? 150
	options.barGapDepthPct = clampChartPct(options.barGapDepthPct, 0, 500, 'barGapDepthPct') ?? 150
	options.barOverlapPct = clampChartPct(options.barOverlapPct, -100, 100, 'barOverlapPct')
	// `<c:holeSize>` is ST_HoleSize (10..90); `<c:firstSliceAng>` is ST_FirstSliceAng (0..360).
	options.holeSize = clampChartPct(options.holeSize, 10, 90, 'holeSize')
	options.firstSliceAng = clampChartPct(options.firstSliceAng, 0, 360, 'firstSliceAng')

	options.chartColors = Array.isArray(options.chartColors)
		? options.chartColors
		: options._type === ChartType.pie || options._type === ChartType.doughnut
			? PIECHART_COLORS
			: BARCHART_COLORS
	options.chartColorsOpacity =
		options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : undefined
	options.plotArea = options.plotArea || {}
	options.plotArea.border =
		options.plotArea.border && typeof options.plotArea.border === 'object' ? options.plotArea.border : undefined
	if (options.plotArea.border && (!options.plotArea.border.width || isNaN(options.plotArea.border.width)))
		options.plotArea.border.width = DEF_CHART_BORDER.width
	if (
		options.plotArea.border &&
		(!options.plotArea.border.color || typeof options.plotArea.border.color !== 'string')
	) {
		options.plotArea.border.color = DEF_CHART_BORDER.color
	}
	options.plotArea.fill = options.plotArea.fill || {}
	options.chartArea = options.chartArea || {}
	options.chartArea.border =
		options.chartArea.border && typeof options.chartArea.border === 'object' ? options.chartArea.border : undefined
	if (options.chartArea.border) {
		options.chartArea.border = {
			color: options.chartArea.border.color || DEF_CHART_BORDER.color,
			width: options.chartArea.border.width || DEF_CHART_BORDER.width,
			transparency: options.chartArea.border.transparency,
		}
	}
	options.chartArea.roundedCorners =
		typeof options.chartArea.roundedCorners === 'boolean' ? options.chartArea.roundedCorners : true
	//
	options.dataBorder = options.dataBorder && typeof options.dataBorder === 'object' ? options.dataBorder : undefined
	if (options.dataBorder && (!options.dataBorder.width || isNaN(options.dataBorder.width)))
		options.dataBorder.width = 0.75
	if (options.dataBorder && options.dataBorder.color) {
		const isHexColor =
			typeof options.dataBorder.color === 'string' &&
			options.dataBorder.color.length === 6 &&
			/^[0-9A-Fa-f]{6}$/.test(options.dataBorder.color)
		const isSchemeColor = Object.values(SchemeColor).includes(options.dataBorder.color as SchemeColor)
		if (!isHexColor && !isSchemeColor) {
			options.dataBorder.color = 'F9F9F9' // Fallback if neither hex nor scheme color
		}
	}
	//
	if (!options.dataLabelFormatCode && options._type === ChartType.scatter) options.dataLabelFormatCode = 'General'
	if (!options.dataLabelFormatCode && (options._type === ChartType.pie || options._type === ChartType.doughnut)) {
		options.dataLabelFormatCode = options.showPercent ? '0%' : 'General'
	}
	options.dataLabelFormatCode =
		options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string'
			? options.dataLabelFormatCode
			: '#,##0'
	//
	// Set default format for Scatter chart labels to custom string if not defined
	if (!options.dataLabelFormatScatter && options._type === ChartType.scatter) options.dataLabelFormatScatter = 'custom'
	//
	options.lineSize = typeof options.lineSize === 'number' ? options.lineSize : 2
	options.valAxisMajorUnit = typeof options.valAxisMajorUnit === 'number' ? options.valAxisMajorUnit : undefined

	if (
		options._type === ChartType.area ||
		options._type === ChartType.bar ||
		options._type === ChartType.bar3d ||
		options._type === ChartType.line
	) {
		options.catAxisMultiLevelLabels = !!options.catAxisMultiLevelLabels
	} else {
		delete options.catAxisMultiLevelLabels
	}

	if (options._type === ChartType.waterfall && options.subtotals !== undefined) {
		// <cx:subtotals> holds zero-based category indices; drop non-integer / negative entries
		// (they would make PowerPoint report the chartEx part as needing repair). Warn per the
		// library's warn-rather-than-degrade policy.
		const clean = (Array.isArray(options.subtotals) ? options.subtotals : []).filter((idx) => {
			const ok = typeof idx === 'number' && Number.isInteger(idx) && idx >= 0
			if (!ok)
				warn(
					'chart/invalid-subtotal-index',
					`chart waterfall subtotal index "${String(idx)}" is not a non-negative integer; entry skipped.`
				)
			return ok
		})
		options.subtotals = clean.length > 0 ? clean : undefined
	}
}

/**
 * Options a combo subchart may override that land in a bounded or enumerated OOXML attribute,
 * i.e. the ones {@link normalizeComboSubchartOptions} is allowed to write back.
 */
const SUBCHART_VALIDATED_KEYS = [
	'barDir',
	'barGrouping',
	'barGapWidthPct',
	'barGapDepthPct',
	'barOverlapPct',
	'bar3DShape',
	'holeSize',
	'firstSliceAng',
	'lineDataSymbol',
	'lineDataSymbolSize',
	'lineDataSymbolLineSize',
	'dataLabelPosition',
] as const

/**
 * Clamp and correct one combo subchart's option overrides.
 *
 * `addChartDefinition` normalizes the chart-level options once, but a combo chart's per-subchart
 * `ChartMulti.options` are merged over them only at emit time (`gen/chart/chart-xml.ts`) — after
 * every clamp and enum correction has already run. Anything set there therefore reached the part
 * verbatim: `barOverlapPct: 250` emitted `<c:overlap val="250"/>` where ST_Overlap is -100..100,
 * `barGapWidthPct: 9999` blew past ST_GapAmount's 500, and `barGrouping: 'sideways'` failed the
 * ST_Grouping enumeration — three PowerPoint-repair prompts reachable only through the combo API.
 *
 * The gap runs the other way too: for a combo chart `options._type` is a `ChartMulti[]`, so the
 * *type-dependent* chart-level corrections (`barGrouping`, `dataLabelPosition`) match no branch
 * and never fire at all.
 *
 * Both are fixed by validating the value the emitter actually reads — `{...chartOptions,
 * ...subOptions}` — against this subchart's own type, then writing back only the keys a
 * correction changed so the subchart bag stays a sparse override of the chart-level options.
 * @param subOptions - caller-supplied `ChartMulti.options` (never written to)
 * @param chartOptions - the already-normalized chart-level options
 * @param subType - this subchart's own plot type
 * @param callerSetBarGapWidthPct - whether the caller supplied a chart-level `barGapWidthPct`
 */
function normalizeComboSubchartOptions(
	subOptions: ChartOpts | undefined,
	chartOptions: ChartOptsInternal,
	subType: ChartType,
	callerSetBarGapWidthPct: boolean
): ChartOpts {
	const sub: ChartOpts = subOptions && typeof subOptions === 'object' ? subOptions : {}
	// What the emitter reads for this subchart today, and the corrected copy to diff against it.
	const merged: ChartOptsInternal = { ...chartOptions, ...sub }
	const fixed: ChartOptsInternal = { ...merged }

	// Enumerations emitted verbatim: `<c:barDir>` (ST_BarDir), `<c:grouping>` (ST_Grouping),
	// `<c:shape>` (ST_Shape), `<c:symbol>` (ST_MarkerStyle).
	if (!['bar', 'col'].includes(fixed.barDir || '')) fixed.barDir = 'col'
	normalizeChartBarGrouping(fixed, subType)
	if (!['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'].includes(fixed.bar3DShape || ''))
		fixed.bar3DShape = 'box'
	if (!['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'].includes(fixed.lineDataSymbol || ''))
		fixed.lineDataSymbol = 'circle'
	// A stacked bar group takes the narrower default gap a chart-level stacked bar gets. The
	// merged bag already carries the clustered default, so only step in when neither the
	// chart-level nor the subchart caller asked for a specific width.
	if (fixed.barGrouping?.includes('tacked') && !callerSetBarGapWidthPct && sub.barGapWidthPct == null)
		fixed.barGapWidthPct = 50
	// Depends on the corrected grouping above, so it has to run after it.
	normalizeChartDataLabelPosition(fixed, subType)

	// Bounded integers. A non-numeric override falls back to the chart-level value, which
	// `normalizeChartOptions` has already put in range.
	fixed.barGapWidthPct = clampChartPct(fixed.barGapWidthPct, 0, 500, 'barGapWidthPct') ?? chartOptions.barGapWidthPct
	fixed.barGapDepthPct = clampChartPct(fixed.barGapDepthPct, 0, 500, 'barGapDepthPct') ?? chartOptions.barGapDepthPct
	fixed.barOverlapPct = clampChartPct(fixed.barOverlapPct, -100, 100, 'barOverlapPct')
	fixed.holeSize = clampChartPct(fixed.holeSize, 10, 90, 'holeSize')
	fixed.firstSliceAng = clampChartPct(fixed.firstSliceAng, 0, 360, 'firstSliceAng')
	// `<c:size val>` is ST_MarkerSize: an integer 2..72 points.
	if (fixed.lineDataSymbolSize != null && !isNaN(fixed.lineDataSymbolSize)) {
		const symbolSize = Math.min(72, Math.max(2, Math.round(fixed.lineDataSymbolSize)))
		if (symbolSize !== fixed.lineDataSymbolSize)
			warn(
				'chart/symbol-size-out-of-range',
				`lineDataSymbolSize ${fixed.lineDataSymbolSize} is outside the valid marker size range (integer 2-72); using ${symbolSize}.`
			)
		fixed.lineDataSymbolSize = symbolSize
	}
	// Points -> EMU, but only for a width this subchart supplied: the chart-level value has
	// already been through `valToPts` and converting it twice would emit a hairline.
	if (sub.lineDataSymbolLineSize != null && !isNaN(sub.lineDataSymbolLineSize))
		fixed.lineDataSymbolLineSize = valToPts(sub.lineDataSymbolLineSize)

	const result: ChartOptsInternal = { ...sub }
	for (const key of SUBCHART_VALIDATED_KEYS) {
		if (fixed[key] !== merged[key]) (result as Record<string, unknown>)[key] = fixed[key]
	}
	return result
}

/**
 * Generate the chart based on input data.
 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
 *
 * @param {CHART_NAME | ChartMulti[]} `type` should belong to: 'column', 'pie'
 * @param {[]} `data` a JSON object with follow the following format
 * @param {ChartOptsInternal} `opt` chart options
 * @param {PresSlideInternal} `target` slide object that the chart will be added to
 * @return {object} chart object
 * {
 *    title: 'eSurvey chart',
 *    data: [
 *        {
 *            name: 'Income',
 *            labels: ['2005', '2006', '2007', '2008', '2009'],
 *            values: [23.5, 26.2, 30.1, 29.5, 24.6]
 *        },
 *        {
 *            name: 'Expense',
 *            labels: ['2005', '2006', '2007', '2008', '2009'],
 *            values: [18.1, 22.8, 23.9, 25.1, 25]
 *        }
 *    ]
 * }
 */
export function addChartDefinition(
	target: PresSlideInternal,
	type: CHART_NAME | ChartMulti[],
	data: OptsChartData[] | ChartOpts,
	opt?: ChartOptsInternal
): object {
	function correctGridLineOptions(glOpts: OptsChartGridLine): void {
		if (!glOpts || glOpts.style === 'none') return
		if (glOpts.size !== undefined && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
			warn('chart/invalid-grid-line-size', 'chart.gridLine.size must be greater than 0.')
			delete glOpts.size // delete prop to used defaults
		}
		if (glOpts.style && !['solid', 'dash', 'dot'].includes(glOpts.style)) {
			warn('chart/invalid-grid-line-style', 'chart.gridLine.style options: `solid`, `dash`, `dot`.')
			delete glOpts.style
		}
		if (glOpts.cap && !['flat', 'square', 'round'].includes(glOpts.cap)) {
			warn('chart/invalid-grid-line-cap', 'chart.gridLine.cap options: `flat`, `square`, `round`.')
			delete glOpts.cap
		}
	}

	// Placeholder part identity, unique only within this target. The authoritative,
	// package-unique chart part filename is assigned at write time by a per-presentation
	// pass in `exportPresentation` (see backlog fork-chart-counter-nondeterminism): a
	// module-global counter here was never reset, so two identical decks built in one
	// process emitted different chart part filenames (same input, different bytes).
	const chartId = target._relsChart.length + 1
	const resultObject: SlideObject = {
		_type: SlideObjectType.chart,
	}
	// DESIGN: `type` can an object (ex: `ChartType.doughnut`) or an array of chart objects
	// EX: addChartDefinition([ { type:ChartType.bar, data:{name:'', labels:[], values[]} }, {<etc>} ])
	// Multi-Type Charts
	let tmpOpt: ChartOpts | ChartOptsInternal | undefined
	let tmpData: OptsChartDataInternal[] = []
	let tmpTypes: ChartMulti[] | undefined
	if (Array.isArray(type)) {
		// For multi-type charts there needs to be data for each type, as well as a single data
		// source for non-series operations. The series are indexed across subcharts to keep the
		// data in order when segmented into types.
		//
		// The whole `ChartMulti[]` is rebuilt around the normalized copies, not just flattened
		// into `tmpData`: the combo emit path plots each subchart from `opts._type[i].data`
		// (see gen/chart/chart-xml.ts), so both views have to reference the same series objects.
		let seriesIndex = 0
		tmpTypes = type.map((obj) => {
			const seriesData = (Array.isArray(obj.data) ? obj.data : []).map((item) =>
				normalizeChartSeries(item, seriesIndex++)
			)
			tmpData = tmpData.concat(seriesData)
			return { ...obj, data: seriesData }
		})
		tmpOpt = !Array.isArray(data) && data && typeof data === 'object' ? data : opt
	} else {
		tmpData = (Array.isArray(data) ? data : []).map(normalizeChartSeries)
		tmpOpt = opt
	}
	// Everything below normalizes onto this copy; the caller's options object is a read-only input.
	const options: ChartOptsInternal = copyChartOptions(tmpOpt && typeof tmpOpt === 'object' ? tmpOpt : {})
	// Captured before normalization fills in the default, so the combo pass below can tell an
	// explicit gap width from an inherited one.
	const callerSetBarGapWidthPct = typeof options.barGapWidthPct === 'number' && !isNaN(options.barGapWidthPct)

	// STEP 1: Set default options/decode user options
	// A: Core
	options._type = tmpTypes ?? asChartType(type as CHART_NAME)
	options.x = typeof options.x !== 'undefined' && options.x != null && !isNaN(Number(options.x)) ? options.x : 1
	options.y = typeof options.y !== 'undefined' && options.y != null && !isNaN(Number(options.y)) ? options.y : 1
	options.w = options.w || '50%'
	options.h = options.h || '50%'
	options.objectName = options.objectName
		? encodeXmlAttrValue(validateObjectName(options.objectName, 'chart'))
		: `Chart ${target._slideObjects.filter((obj) => obj._type === SlideObjectType.chart).length}`

	// B: Options: misc
	if (!['bar', 'col'].includes(options.barDir || '')) options.barDir = 'col'

	// barGrouping must be handled before data label validation as it can affect valid label positioning
	const chartLevelType = Array.isArray(options._type) ? undefined : options._type
	normalizeChartBarGrouping(options, chartLevelType)
	if (options.barGrouping?.includes('tacked')) {
		if (!options.barGapWidthPct) options.barGapWidthPct = 50
	}
	// Clean up and validate data label positions
	// REFERENCE: https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e2b1697c-7adc-463d-9081-3daef72f656f?redirectedfrom=MSDN
	normalizeChartDataLabelPosition(options, chartLevelType)
	// dataLabelBkgrdColors: same dead-ternary shape as the show* block in
	// normalizeChartPlotAreaOptions, same reason for its absence.
	if (!['b', 'l', 'r', 't', 'tr'].includes(options.legendPos || '')) options.legendPos = 'r'

	// 3D bar: ST_Shape
	if (!['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'].includes(options.bar3DShape || ''))
		options.bar3DShape = 'box'
	// lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
	// Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
	if (!['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'].includes(options.lineDataSymbol || ''))
		options.lineDataSymbol = 'circle'
	if (!['gap', 'span', 'zero'].includes(options.displayBlanksAs || '')) options.displayBlanksAs = 'gap'
	if (!['radar', 'markers', 'filled'].includes(options.radarStyle || '')) options.radarStyle = 'radar'
	// Marker size emits as `<c:size val>` (ST_MarkerSize): an integer in [2,72] points.
	// Out-of-range or non-integer values make PowerPoint report the file as needing
	// repair, so round and clamp into range and warn when the input is coerced.
	{
		const rawSymbolSize = options.lineDataSymbolSize
		const hasSymbolSize = rawSymbolSize != null && !isNaN(rawSymbolSize)
		const symbolSize = Math.min(72, Math.max(2, Math.round(hasSymbolSize ? rawSymbolSize : 6)))
		if (hasSymbolSize && symbolSize !== rawSymbolSize) {
			warn(
				'chart/symbol-size-out-of-range',
				`lineDataSymbolSize ${rawSymbolSize} is outside the valid marker size range (integer 2-72); using ${symbolSize}.`
			)
		}
		options.lineDataSymbolSize = symbolSize
	}
	options.lineDataSymbolLineSize =
		options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize)
			? valToPts(options.lineDataSymbolLineSize)
			: valToPts(0.75)
	// `layout` allows the override of PPT defaults to maximize space
	const chartLayout = options.layout
	if (chartLayout) {
		;(['x', 'y', 'w', 'h'] as const).forEach((key) => {
			const val = chartLayout[key]
			const numVal = Number(val)
			if (isNaN(numVal) || numVal < 0 || numVal > 1) {
				warn('chart/layout-out-of-range', 'chart.layout.' + key + ' can only be 0-1')
				delete chartLayout[key] // remove invalid value so that default will be used
			}
		})
	}

	// Set gridline defaults
	options.catGridLine =
		options.catGridLine || (options._type === ChartType.scatter ? { color: 'D9D9D9', size: 1 } : { style: 'none' })
	options.valGridLine = options.valGridLine || (options._type === ChartType.scatter ? { color: 'D9D9D9', size: 1 } : {})
	options.serGridLine =
		options.serGridLine || (options._type === ChartType.scatter ? { color: 'D9D9D9', size: 1 } : { style: 'none' })
	correctGridLineOptions(options.catGridLine)
	correctGridLineOptions(options.valGridLine)
	correctGridLineOptions(options.serGridLine)
	correctShadowOptions(options.shadow)

	// C: Options: plotArea
	normalizeChartPlotAreaOptions(options)

	// D: Options: chart
	// `<c:gapWidth>`/`<c:gapDepth>` are ST_GapAmount (integer 0..500); `<c:overlap>` is
	// ST_Overlap (integer -100..100). Out-of-range values trigger PowerPoint repair.
	normalizeChartOptions(options)

	// E: Options: combo subcharts
	// A `ChartMulti` entry's options override the chart-level ones at emit time, so they have to
	// go through the same clamps and enum corrections — keyed to that subchart's own plot type.
	if (Array.isArray(options._type)) {
		options._type.forEach((sub) => {
			// Safe to assign: these entries are the copies built above, not the caller's objects.
			sub.options = normalizeComboSubchartOptions(sub.options, options, asChartType(sub.type), callerSetBarGapWidthPct)
		})
	}

	// Stock charts require their series in a fixed order (see `stockStyle`); default to the
	// three-value High-Low-Close style and warn (rather than corrupt) when the number of data
	// series doesn't match the style, since PowerPoint expects an exact count per style.
	if (options._type === ChartType.stock) {
		const STOCK_SERIES_COUNT: Record<string, number> = { hlc: 3, ohlc: 4, vhlc: 4, vohlc: 5 }
		if (!Object.keys(STOCK_SERIES_COUNT).includes(options.stockStyle || '')) options.stockStyle = 'hlc'
		const expected = STOCK_SERIES_COUNT[options.stockStyle as string]
		if (tmpData.length !== expected) {
			warn(
				'chart/stock-series-count',
				`stock chart style "${options.stockStyle}" expects ${expected} data series (got ${tmpData.length}); the chart may not render as intended.`
			)
		}
	}

	// STEP 4: Set props
	resultObject._type = SlideObjectType.chart
	resultObject.options = options
	resultObject.chartRid = getNewRelId(target)

	// STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	// chartEx charts (waterfall, …) live alongside classic charts in `ppt/charts/` but use the
	// `chartEx{N}.xml` name, the `chartex+xml` content type and the MS chartEx rel type. The
	// authoritative, package-unique filename is (re)assigned at write time in `exportPresentation`;
	// this placeholder mirrors the same Ex-prefix rule so a single-chart deck is already correct.
	const isChartEx = isChartExType(options._type)
	const chartBase = isChartEx ? `chartEx${chartId}` : `chart${chartId}`
	target._relsChart.push({
		rId: getNewRelId(target),
		data: tmpData,
		opts: options,
		type: options._type,
		globalId: chartId,
		isChartEx,
		fileName: `${chartBase}.xml`,
		Target: `/ppt/charts/${chartBase}.xml`,
	})

	target._slideObjects.push(resultObject)
	return resultObject
}
