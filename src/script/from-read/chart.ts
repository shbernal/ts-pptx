/**
 * `Chart` → `addChart(OptsChartData[], ChartOpts)`.
 *
 * The open question this mapper settles is where the data comes from. A `.pptx` chart
 * carries its numbers twice: once in an embedded `.xlsx` workbook, and once as *cached*
 * values inside the chart XML itself (`c:tx`, `c:cat`, `c:val`) so a viewer can render
 * without opening the workbook. Probing the fixtures showed the cache is complete for
 * every chart the write API can express, so the workbook never needs cracking — which
 * removes a whole spreadsheet parser from this project's dependency surface.
 *
 * The one seam is blanks. The read model reports `(number | null)[]`, since a cached point
 * may legitimately be empty, while `OptsChartData.values` is `number[]`. A blank becomes
 * `0`, which is a visible change on a line chart — a gap turns into a dip to the axis — so
 * it is noted rather than quietly coerced.
 *
 * A rebuilt chart also gets a *regenerated* workbook rather than the source one, so
 * double-clicking it in PowerPoint opens data that matches the plot but has lost the source
 * sheet's formulas, extra columns, and formatting.
 */
import type { Chart, ChartSeries } from '../../read/api/chart.js'
import type { GraphicFrame } from '../../read/api/shapes.js'
import type { NoteScope } from '../fidelity.js'
import type { CallIr, IrValue } from '../ir.js'
import { compact, frameOf, identityOptions, literalColor, nameOf, orUndefined, positionOptions } from './values.js'

/**
 * Read chart-group token → `CHART_NAME`. The read model strips the `Chart` suffix from the
 * plot-area element name, so `c:barChart` arrives as `bar`. A 3-D group maps onto the
 * matching flat type where one exists; anything absent from this table has no write-API
 * chart type and is reported rather than guessed at.
 */
const CHART_TYPE: Record<string, string> = {
	area: 'area',
	area3D: 'area',
	bar: 'bar',
	bar3D: 'bar3D',
	bubble: 'bubble',
	bubble3D: 'bubble3D',
	doughnut: 'doughnut',
	line: 'line',
	line3D: 'line',
	pie: 'pie',
	pie3D: 'pie',
	radar: 'radar',
	scatter: 'scatter',
	stock: 'stock',
	surface: 'surface',
	surface3D: 'surface',
}

/**
 * The 3-D groups rebuilt flat, because the write API has no 3-D type for them. `bar3D` and
 * `bubble3D` keep their own, and a 3-D surface is what the write side's `surface` already draws.
 */
const FLATTENED_3D: ReadonlySet<string> = new Set(['area3D', 'line3D', 'pie3D'])

/**
 * The plot types whose `showLabel` writes the category name. A scatter's writes nothing unless a
 * `dataLabelFormatScatter` is also set, so a scatter is not among them.
 */
const CATEGORY_NAME_TYPES: ReadonlySet<string> = new Set(['doughnut', 'pie'])

/**
 * The plot types that pair every value with an X value of its own (`c:xVal` and `c:yVal`, plus
 * `c:bubbleSize` on a bubble) rather than with a category. `addChart` spells them as a leading row
 * of X values followed by one row per series.
 */
const XY_TYPES: ReadonlySet<string> = new Set(['bubble', 'bubble3D', 'scatter'])

/** `c:legendPos/@val` → the write API's `legendPos`. */
const LEGEND_POS: Record<string, string> = { r: 'r', l: 'l', t: 't', b: 'b', tr: 'tr' }

export function chartCall(frame: GraphicFrame, chart: Chart, notes: NoteScope): CallIr | null {
	const box = frameOf(frame, notes)
	if (!box) return null
	const sourceType = chart.chartType
	const type = sourceType === null ? undefined : CHART_TYPE[sourceType]
	if (!type) {
		notes.note(
			'chart.type',
			'dropped',
			'unwritable',
			`chart type "${sourceType ?? 'none'}" has no write-API counterpart, so the chart is omitted`
		)
		return null
	}

	if (sourceType !== null && FLATTENED_3D.has(sourceType)) {
		notes.note(
			'chart.type3D',
			'flattened',
			'unwritable',
			`a 3-D ${type} chart (c:${sourceType}Chart) is rebuilt as a flat ${type} chart: the write API has no 3-D ${type} type, so its depth, rotation and perspective are gone`
		)
	}

	if (chart.chartTypes.length > 1) {
		notes.note(
			'chart.combo',
			'flattened',
			'unsupported',
			`a combo chart (${chart.chartTypes.join(' + ')}) is emitted as a single ${type} chart; addChart's multi-type form needs a per-series type the read model does not report`
		)
	}

	const xy = XY_TYPES.has(type)
	const data = xy ? xyData(chart, type, notes) : seriesData(chart, notes)
	// Series that all read no points plot nothing either. `addChart` would draw an empty frame the
	// source never showed as a chart, and checking only for zero series let those through.
	if (data.length === 0 || chart.series.every((series) => (xy ? series.yValues : series.values).length === 0)) {
		notes.note('chart.data', 'dropped', 'unsupported', 'this chart caches no plottable series values')
		return null
	}

	// BELOW the guard above, deliberately. Recorded unconditionally, a chart with no cached
	// series emitted two notes — one saying it was dropped, one saying it was rebuilt — and this
	// one maps to `['*']`, so on a dropped chart it also excused every difference on that frame:
	// the widest exclusion in the table, applied to the case it least describes.
	notes.note(
		'chart.workbook',
		'approximated',
		'unsupported',
		"the chart is rebuilt from its cached plot values, so its embedded workbook is regenerated: the plotted numbers match, but the source sheet's formulas, extra columns and formatting are gone"
	)

	// `type` is an *option*, not a positional argument: the signature is
	// `addChart(data, options & { type })`. Passing it separately produces a call that
	// typechecks nowhere and throws "a chart `type` is required" at run time.
	const options = compact({
		type,
		// `c:surfaceChart` is the 2-D contour; the write side's `surface` draws the 3-D one unless told
		// otherwise, so a flat surface came back tilted.
		surface3D: sourceType === 'surface' ? false : undefined,
		...positionOptions(box),
		...identityOptions(frame),
		...titleOptions(chart),
		...legendOptions(chart),
		...labelOptions(chart, type, notes),
		...axisOptions(chart),
		chartColors: seriesColors(chart),
	})

	return {
		method: 'addChart',
		args: [data, options ?? { type }],
		...nameOf(frame),
	}
}

/**
 * Series as `OptsChartData[]`. Category labels ride on the first series, which is the shape
 * `addChart` expects; the read model reports them per-series but they are shared in
 * practice.
 */
function seriesData(chart: Chart, notes: NoteScope): IrValue[] {
	// Leaf level first, which is also how `OptsChartData.labels` takes several levels. An outer
	// level names each group once, and the writer spells the rest of the group as `''`.
	const levels = chart.categoryLevels.map((level) => level.map((label) => label ?? ''))
	const labels = levels.length > 1 ? levels : levels[0]
	const dense = blanksToZero(notes)

	return chart.series.map((series, index) => {
		const values = dense(series.values)
		return (
			compact({
				name: orUndefined(series.name),
				values,
				// Only the first series carries labels, matching how addChart reads them.
				labels: index === 0 && levels[0]?.length ? labels : undefined,
			}) ?? { values }
		)
	})
}

/**
 * A scatter or bubble chart as `OptsChartData[]`: the X row, then one row per series with its Y
 * values and, on a bubble, its sizes. The X row is not a series and has no name in the chart part;
 * its column header lives only in the workbook, which is regenerated anyway.
 *
 * `addChart` plots every series against that one X row, while a chart part gives each series a
 * `c:xVal` of its own. The first series' X values are taken, and a series whose own differ is noted.
 * @param type - the write-API chart type the chart is rebuilt as
 */
function xyData(chart: Chart, type: string, notes: NoteScope): IrValue[] {
	const series = chart.series
	const first = series[0]
	if (!first) return []
	const dense = blanksToZero(notes)
	const xValues = plottedX(first)
	const withSizes = type !== 'scatter'

	const rows: IrValue[] = [{ values: dense(xValues) }]
	let ownX = false
	for (const one of series) {
		const own = plottedX(one)
		if (own.length !== xValues.length || own.some((value, i) => value !== xValues[i])) ownX = true
		const values = dense(one.yValues)
		rows.push(
			compact({
				name: orUndefined(one.name),
				values,
				sizes: withSizes ? dense(one.bubbleSizes) : undefined,
			}) ?? { values }
		)
	}

	if (ownX) {
		notes.note(
			'chart.xValues',
			'approximated',
			'unwritable',
			`the series of this ${type} chart do not share one set of X values, and addChart plots every series against a single X row, so each is rebuilt against the first series' X values`
		)
	}
	if (series.some((one) => one.xLabels !== null)) {
		notes.note(
			'chart.xLabels',
			'dropped',
			'unwritable',
			`this ${type} chart is plotted against text X labels, which PowerPoint places at 1 to n in point order; it is rebuilt against those positions and plots the same, but addChart takes numbers for its X row, so the labels are gone`
		)
	}
	return rows
}

/**
 * The X coordinates a series is plotted at. A series plotted against text has none of its own, and
 * neither has a series with no `c:xVal`: PowerPoint places the points of both at 1 to n, and an X row
 * of those positions draws the same chart. Taking the X values instead turned every label into a
 * blank and the blank into a 0, and gave a series with no `c:xVal` an empty X row, against which
 * `addChart` plots no point at all.
 */
function plottedX(series: ChartSeries): (number | null)[] {
	const labels = series.xLabels
	if (labels !== null) return positions(labels.length)
	const xValues = series.xValues
	return xValues.length === 0 ? positions(series.yValues.length) : xValues
}

/** The positions 1 to `count`. */
function positions(count: number): number[] {
	return Array.from({ length: count }, (_unused, i) => i + 1)
}

/**
 * Densify cached points for `number[]`, noting once per chart when a blank had to become 0.
 * The returned function is called once per point list, and the note is recorded on the first blank.
 */
function blanksToZero(notes: NoteScope): (points: (number | null)[]) => number[] {
	let noted = false
	return (points) =>
		points.map((value) => {
			if (value !== null) return value
			if (!noted) {
				noted = true
				notes.note(
					'chart.blanks',
					'approximated',
					'unwritable',
					'a blank cached data point becomes 0, because OptsChartData.values is number[] and has no spelling for a gap; on a line chart this draws a dip to the axis where the source showed a break'
				)
			}
			return 0
		})
}

function titleOptions(chart: Chart): Record<string, IrValue | undefined> {
	const title = chart.title
	return title === null ? {} : { showTitle: true, title }
}

function legendOptions(chart: Chart): Record<string, IrValue | undefined> {
	const legend = chart.legend
	if (!legend) return { showLegend: false }
	return {
		showLegend: true,
		legendPos: legend.position === null ? undefined : LEGEND_POS[legend.position],
	}
}

/**
 * Data-label flags as `ChartOpts`.
 *
 * Two of the read model's flags have no key of their own on the write side, and they were emitted
 * as `showCatName` and `showLegendKey` anyway. Neither is a `ChartOpts` key, so the writer ignored
 * both and nothing said so. `ChartOpts` spells the category name `showLabel`, which a pie and a
 * doughnut write as one; nothing spells the legend-key swatch.
 *
 * A pie and a doughnut are read from the series' own block. PowerPoint writes the flags the user set
 * there and an all-off block for the group, so reading the group block rebuilt a labelled pie
 * without its labels.
 * @param type - the write-API chart type the chart is rebuilt as
 */
function labelOptions(chart: Chart, type: string, notes: NoteScope): Record<string, IrValue | undefined> {
	const carriesCategoryName = CATEGORY_NAME_TYPES.has(type)
	const labels = (carriesCategoryName ? chart.series[0]?.dataLabels : null) ?? chart.dataLabels
	if (!labels) return {}

	if (labels.showLegendKey === true) {
		notes.note(
			'chart.labels',
			'dropped',
			'unwritable',
			'the data labels show the legend key swatch (c:showLegendKey), which no ChartOpts option spells, so the labels come back without it'
		)
	}
	if (labels.showCategoryName === true && !carriesCategoryName) {
		notes.note(
			'chart.labels',
			'dropped',
			'unwritable',
			`the data labels show the category name (c:showCatName); ChartOpts spells that showLabel, which only a pie and a doughnut write as the category name, so on this ${type} chart the labels come back without it`
		)
	}

	return {
		showValue: labels.showValue ?? undefined,
		showSerName: labels.showSeriesName ?? undefined,
		showLabel: carriesCategoryName ? (labels.showCategoryName ?? undefined) : undefined,
		showPercent: labels.showPercent ?? undefined,
		showLeaderLines: labels.showLeaderLines ?? undefined,
		// A pie's labels with no `c:dLblPos` sit at best fit (PowerPoint reports
		// `xlLabelPositionBestFit` for them), and the writer spells an unset pie position `ctr`.
		dataLabelPosition: orUndefined(labels.position) ?? (type === 'pie' ? 'bestFit' : undefined),
		dataLabelFormatCode: labels.numberFormat?.formatCode ?? undefined,
	}
}

/**
 * Axis bounds, titles and gridlines. The write API spells these as a flat `catAxis*` /
 * `valAxis*` prefix set rather than an axis object, so each side is unpacked by hand.
 */
function axisOptions(chart: Chart): Record<string, IrValue | undefined> {
	const cat = chart.categoryAxis
	const val = chart.valueAxis
	return {
		catAxisHidden: cat?.hidden ? true : undefined,
		catAxisTitle: cat?.title ?? undefined,
		valAxisHidden: val?.hidden ? true : undefined,
		valAxisTitle: val?.title ?? undefined,
		valAxisMinVal: val?.min ?? undefined,
		valAxisMaxVal: val?.max ?? undefined,
		valAxisMajorUnit: val?.majorUnit ?? undefined,
		valAxisLabelFormatCode: val?.numberFormat?.formatCode ?? undefined,
	}
}

/**
 * Per-series fill colours as `chartColors`, which is positional — entry `n` colours series
 * `n`. A series with no explicit fill takes the theme's accent for its position, which the
 * read model does not resolve, so the whole option is dropped unless every series names a
 * colour; a partial list would silently recolour the unnamed ones.
 */
function seriesColors(chart: Chart): IrValue | undefined {
	const series = chart.series
	if (series.length === 0) return undefined
	const colors: string[] = []
	for (const one of series) {
		const fill = one.fill
		const hex = fill?.colorRef.srgb
		if (!hex) return undefined
		colors.push(literalColor(hex))
	}
	return colors
}
