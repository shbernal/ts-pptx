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
import type { Chart } from '../../read/api/chart.js'
import type { GraphicFrame } from '../../read/api/shapes.js'
import type { NoteScope } from '../fidelity.js'
import type { CallIr, IrValue } from '../ir.js'
import { compact, emu, literalColor, orUndefined } from './values.js'

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

/** `c:legendPos/@val` → the write API's `legendPos`. */
const LEGEND_POS: Record<string, string> = { r: 'r', l: 'l', t: 't', b: 'b', tr: 'tr' }

export function chartCall(frame: GraphicFrame, chart: Chart, notes: NoteScope): CallIr | null {
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

	if (chart.chartTypes.length > 1) {
		notes.note(
			'chart.combo',
			'flattened',
			'unsupported',
			`a combo chart (${chart.chartTypes.join(' + ')}) is emitted as a single ${type} chart; addChart's multi-type form needs a per-series type the read model does not report`
		)
	}

	notes.note(
		'chart.workbook',
		'approximated',
		'unsupported',
		"the chart is rebuilt from its cached plot values, so its embedded workbook is regenerated: the plotted numbers match, but the source sheet's formulas, extra columns and formatting are gone"
	)

	const data = seriesData(chart, notes)
	if (data.length === 0) {
		notes.note('chart.data', 'dropped', 'unsupported', 'this chart caches no plottable series values')
		return null
	}

	// `type` is an *option*, not a positional argument: the signature is
	// `addChart(data, options & { type })`. Passing it separately produces a call that
	// typechecks nowhere and throws "a chart `type` is required" at run time.
	const options = compact({
		type,
		...positionOfFrame(frame),
		objectName: frame.name || undefined,
		...titleOptions(chart),
		...legendOptions(chart),
		...labelOptions(chart),
		...axisOptions(chart),
		chartColors: seriesColors(chart),
	})

	return {
		method: 'addChart',
		args: [data, options ?? { type }],
		...(frame.name ? { sourceName: frame.name } : {}),
	}
}

/** A graphic frame's position; frames never rotate or flip. */
function positionOfFrame(frame: GraphicFrame): Record<string, IrValue> {
	const box = frame.absoluteFrame
	if (!box) return {}
	return { x: emu(box.left), y: emu(box.top), w: emu(box.width), h: emu(box.height) }
}

/**
 * Series as `OptsChartData[]`. Category labels ride on the first series, which is the shape
 * `addChart` expects; the read model reports them per-series but they are shared in
 * practice.
 */
function seriesData(chart: Chart, notes: NoteScope): IrValue[] {
	const categories = chart.categories.map((label) => label ?? '')
	let sawBlank = false

	const data = chart.series.map((series, index) => {
		const values = series.values.map((value) => {
			if (value === null) sawBlank = true
			return value ?? 0
		})
		return (
			compact({
				name: orUndefined(series.name),
				values,
				// Only the first series carries labels, matching how addChart reads them.
				labels: index === 0 && categories.length > 0 ? categories : undefined,
			}) ?? { values }
		)
	})

	if (sawBlank) {
		notes.note(
			'chart.blanks',
			'approximated',
			'unwritable',
			'a blank cached data point becomes 0, because OptsChartData.values is number[] and has no spelling for a gap; on a line chart this draws a dip to the axis where the source showed a break'
		)
	}
	return data
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

function labelOptions(chart: Chart): Record<string, IrValue | undefined> {
	const labels = chart.dataLabels
	if (!labels) return {}
	return {
		showValue: labels.showValue ?? undefined,
		showSerName: labels.showSeriesName ?? undefined,
		showCatName: labels.showCategoryName ?? undefined,
		showPercent: labels.showPercent ?? undefined,
		showLegendKey: labels.showLegendKey ?? undefined,
		showLeaderLines: labels.showLeaderLines ?? undefined,
		dataLabelPosition: orUndefined(labels.position),
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
		const hex = fill?.color
		if (!hex) return undefined
		colors.push(literalColor(hex))
	}
	return colors
}
