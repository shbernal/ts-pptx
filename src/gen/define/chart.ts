/**
 * PptxGenJS: Chart Definition
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
import { warn } from '../../log.js'
import type { ChartMulti, ChartOpts, OptsChartData, OptsChartGridLine } from '../../core-interfaces.js'
import type { ChartOptsInternal, OptsChartDataInternal, PresSlideInternal, SlideObject } from '../../types/internal.js'
import { encodeXmlEntities, getNewRelId, validateObjectName } from '../../gen-utils.js'
import { correctShadowOptions } from '../drawingml/effect.js'
import { valToPts } from '../../units-internal.js'

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
	if (clamped !== value) warn(`${name} ${value} is outside the valid range ${min}-${max}; using ${clamped}.`)
	return clamped
}

/**
 * Drop `dataLabelPosition` values that are invalid for the chart type / bar grouping,
 * per the OOXML data-label placement rules, so PowerPoint does not flag the file.
 */
function normalizeChartDataLabelPosition(options: ChartOptsInternal): void {
	if (options.dataLabelPosition) {
		const dataLabelPosition = options.dataLabelPosition
		if (
			options._type === ChartType.area ||
			options._type === ChartType.bar3d ||
			options._type === ChartType.doughnut ||
			options._type === ChartType.radar
		) {
			delete options.dataLabelPosition
		}
		if (options._type === ChartType.pie) {
			if (!['bestFit', 'ctr', 'inEnd', 'outEnd'].includes(dataLabelPosition)) delete options.dataLabelPosition
		}
		if (
			options._type === ChartType.bubble ||
			options._type === ChartType.bubble3d ||
			options._type === ChartType.line ||
			options._type === ChartType.scatter
		) {
			if (!['b', 'ctr', 'l', 'r', 't'].includes(dataLabelPosition)) delete options.dataLabelPosition
		}
		if (options._type === ChartType.bar) {
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
 * Apply plotArea option defaults: show* toggles, axis-line visibility, and the 3D view angles.
 */
function normalizeChartPlotAreaOptions(options: ChartOptsInternal): void {
	options.showDataTable = options.showDataTable || !options.showDataTable ? options.showDataTable : false
	options.showDataTableHorzBorder =
		options.showDataTableHorzBorder || !options.showDataTableHorzBorder ? options.showDataTableHorzBorder : true
	options.showDataTableVertBorder =
		options.showDataTableVertBorder || !options.showDataTableVertBorder ? options.showDataTableVertBorder : true
	options.showDataTableOutline =
		options.showDataTableOutline || !options.showDataTableOutline ? options.showDataTableOutline : true
	options.showDataTableKeys = options.showDataTableKeys || !options.showDataTableKeys ? options.showDataTableKeys : true
	options.showLabel = options.showLabel || !options.showLabel ? options.showLabel : false
	options.showLegend = options.showLegend || !options.showLegend ? options.showLegend : false
	options.showPercent = options.showPercent || !options.showPercent ? options.showPercent : true
	options.showTitle = options.showTitle || !options.showTitle ? options.showTitle : false
	options.showValue = options.showValue || !options.showValue ? options.showValue : false
	options.showLeaderLines = options.showLeaderLines || !options.showLeaderLines ? options.showLeaderLines : false
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
	options.v3DRAngAx = options.v3DRAngAx || !options.v3DRAngAx ? options.v3DRAngAx : true
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
			if (!ok) warn(`chart waterfall subtotal index "${String(idx)}" is not a non-negative integer; entry skipped.`)
			return ok
		})
		options.subtotals = clean.length > 0 ? clean : undefined
	}
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
			warn('chart.gridLine.size must be greater than 0.')
			delete glOpts.size // delete prop to used defaults
		}
		if (glOpts.style && !['solid', 'dash', 'dot'].includes(glOpts.style)) {
			warn('chart.gridLine.style options: `solid`, `dash`, `dot`.')
			delete glOpts.style
		}
		if (glOpts.cap && !['flat', 'square', 'round'].includes(glOpts.cap)) {
			warn('chart.gridLine.cap options: `flat`, `square`, `round`.')
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
	// DESIGN: `type` can an object (ex: `pptx.ChartType.doughnut`) or an array of chart objects
	// EX: addChartDefinition([ { type:pptx.ChartType.bar, data:{name:'', labels:[], values[]} }, {<etc>} ])
	// Multi-Type Charts
	let tmpOpt: ChartOpts | ChartOptsInternal | undefined
	let tmpData: OptsChartData[] = []
	if (Array.isArray(type)) {
		// For multi-type charts there needs to be data for each type,
		// as well as a single data source for non-series operations.
		// The data is indexed below to keep the data in order when segmented
		// into types.
		type.forEach((obj) => {
			tmpData = tmpData.concat(obj.data)
		})
		tmpOpt = !Array.isArray(data) && data && typeof data === 'object' ? data : opt
	} else {
		tmpData = Array.isArray(data) ? data : []
		tmpOpt = opt
	}
	tmpData.forEach((item, i) => {
		item._dataIndex = i

		// Converts the 'labels' array from string[] to string[][] (or the respective primitive type), if needed
		if (item.labels !== undefined && !Array.isArray(item.labels[0])) {
			item.labels = [item.labels as string[]]
		}
	})
	const options: ChartOptsInternal = tmpOpt && typeof tmpOpt === 'object' ? tmpOpt : {}

	// STEP 1: Set default options/decode user options
	// A: Core
	options._type = Array.isArray(type) ? type : asChartType(type)
	options.x = typeof options.x !== 'undefined' && options.x != null && !isNaN(Number(options.x)) ? options.x : 1
	options.y = typeof options.y !== 'undefined' && options.y != null && !isNaN(Number(options.y)) ? options.y : 1
	options.w = options.w || '50%'
	options.h = options.h || '50%'
	options.objectName = options.objectName
		? encodeXmlEntities(validateObjectName(options.objectName, 'chart'))
		: `Chart ${target._slideObjects.filter((obj) => obj._type === SlideObjectType.chart).length}`

	// B: Options: misc
	if (!['bar', 'col'].includes(options.barDir || '')) options.barDir = 'col'

	// barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
	// barGrouping must be handled before data label validation as it can affect valid label positioning
	if (options._type === ChartType.area) {
		if (!['stacked', 'standard', 'percentStacked'].includes(options.barGrouping || '')) options.barGrouping = 'standard'
	}
	if (options._type === ChartType.bar) {
		if (!['clustered', 'stacked', 'percentStacked'].includes(options.barGrouping || ''))
			options.barGrouping = 'clustered'
	}
	if (options._type === ChartType.bar3d) {
		if (!['clustered', 'stacked', 'standard', 'percentStacked'].includes(options.barGrouping || ''))
			options.barGrouping = 'standard'
	}
	if (options.barGrouping?.includes('tacked')) {
		if (!options.barGapWidthPct) options.barGapWidthPct = 50
	}
	// Clean up and validate data label positions
	// REFERENCE: https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e2b1697c-7adc-463d-9081-3daef72f656f?redirectedfrom=MSDN
	normalizeChartDataLabelPosition(options)
	options.dataLabelBkgrdColors =
		options.dataLabelBkgrdColors || !options.dataLabelBkgrdColors ? options.dataLabelBkgrdColors : false
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
				warn('chart.layout.' + key + ' can only be 0-1')
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
		data: tmpData as OptsChartDataInternal[],
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
