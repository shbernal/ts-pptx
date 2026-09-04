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
	isChartType,
	SchemeColor,
	SlideObjectType,
} from '../../enums.js'
import { DEF_CHART_BORDER } from '../../constants-internal.js'
import { isHexColor } from '../../hex-color.js'
import { checkEnumOrWarn } from '../../ooxml/check-enum.js'
import {
	BAR_3D_SHAPES,
	BAR_DIRECTIONS,
	BAR_GROUPINGS,
	DISPLAY_BLANKS_AS,
	GROUPINGS,
	LEGEND_POSITIONS,
	LINE_DATA_SYMBOLS,
	RADAR_STYLE_ALIAS_NAMES,
	RADAR_STYLE_ALIASES,
	RADAR_STYLES,
	type RadarStyleAlias,
} from '../../ooxml/st-enums.js'
import { defaultChartPalette } from '../chart/chart-parts.js'
import { warn } from '../../diagnostics.js'
import type { DiagnosticCode } from '../../codes.js'
import { InvalidOptionError } from '../../errors.js'
import type {
	BorderProps,
	ChartMulti,
	ChartOpts,
	ChartSeriesOpts,
	OptsChartData,
	OptsChartGridLine,
} from '../../types/index.js'
import { scrubGridLine } from '../chart/chart-stroke.js'
import type {
	ChartMultiInternal,
	ChartOptsInternal,
	ChartOptsOverrides,
	OptsChartDataInternal,
	PresSlideInternal,
	SlideObject,
} from '../../types/internal.js'
import { getNewRelId } from '../utils.js'
import { resolveObjectName } from './object-name.js'
import { setOrClear } from '../../options-internal.js'
import { normalizeShadowOptions } from '../drawingml/effect.js'
import { clampRangedInput, lineWidthToEmu, ptsToEmuLenient } from '../../units-internal.js'
import { isBubbleChart, STOCK_STYLE_SPEC, type StockStyle } from '../chart/chart-kind.js'

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
	// `labels` is destructured out rather than overwritten so a series with none carries no
	// `labels` key at all: absent is this model's one spelling of "no categories", and the
	// emitters read it with `series.labels?.length`, which cannot tell the two apart anyway.
	const { labels, ...rest } = item
	const series: OptsChartDataInternal = { ...rest, _dataIndex: index }
	if (labels !== undefined) series.labels = Array.isArray(labels[0]) ? (labels as string[][]) : [labels as string[]]
	return series
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
		// `typeof === 'object'`, not just truthy: `fill` takes a bare colour string as the
		// solid-fill shorthand, and spreading a string produces `{0:'F',1:'F',…}` — an object
		// that names neither `color` nor `type`, so `isStatedFill` rejected it and the area
		// came out `<a:noFill/>`. A string is immutable and needs no defensive copy anyway.
		if (typeof copy.plotArea.fill === 'object') copy.plotArea.fill = { ...copy.plotArea.fill }
	}
	if (copy.chartArea) copy.chartArea = { ...copy.chartArea }
	if (copy.dataBorder) copy.dataBorder = { ...copy.dataBorder }
	if (copy.layout) copy.layout = { ...copy.layout }
	if (copy.catGridLine) copy.catGridLine = { ...copy.catGridLine }
	if (copy.valGridLine) copy.valGridLine = { ...copy.valGridLine }
	if (copy.serGridLine) copy.serGridLine = { ...copy.serGridLine }
	return copy
}

/**
 * A border width the emitter can use: a positive, finite number of points.
 *
 * `0` counts as "not stated" here and takes the default, which is what the truthiness test
 * this replaces already did. What it adds is `Infinity` and a negative — neither is a width,
 * and a negative one reached `a:ln/@w` as a negative attribute.
 */
function isUsableBorderWidth(width: number | undefined): boolean {
	return typeof width === 'number' && Number.isFinite(width) && width > 0
}

/**
 * Round and clamp an integer chart option into the range its schema type allows.
 *
 * Several chart attributes are bounded integer types whose out-of-range values make PowerPoint
 * report the package as needing repair: `<c:overlap>` (ST_Overlap, -100..100),
 * `<c:gapWidth>`/`<c:gapDepth>` (ST_GapAmount, 0..500), `<c:holeSize>` (ST_HoleSize, 10..90),
 * `<c:firstSliceAng>` (ST_FirstSliceAng, 0..360) and `<c:size>` (ST_MarkerSize, 2..72).
 *
 * The out-of-range policy is {@link clampRangedInput}'s, which is the library's one policy and
 * says so in its own docblock: a finite value has a nearest legal neighbour, so it clamps and
 * warns; a value that is not a number at all has none, so the request is discarded and that
 * throws. This helper used to answer `undefined` for the second case -- discarding the request
 * and reporting nothing -- so `holeSize: NaN` silently took the default while `holeSize: 200`
 * warned. Same option, same class of mistake, two behaviours.
 *
 * `undefined` in still means `undefined` out: that is the genuine "unset" case, and the callers
 * spell their own default with `??` or `setOrClear`.
 *
 * @param value - caller-supplied option value
 * @param min - inclusive lower bound
 * @param max - inclusive upper bound
 * @param name - option name, as the caller spells it, for the diagnostic
 * @param code - the out-of-range diagnostic this option raises
 */
function clampChartInt(
	value: number | undefined | null,
	min: number,
	max: number,
	name: string,
	code: DiagnosticCode = 'chart/option-out-of-range'
): number | undefined {
	if (value === undefined || value === null) return undefined
	const clamped = clampRangedInput(value, min, max, code, name, 'chart/option-non-finite')
	const rounded = Math.round(clamped)
	// A fractional in-range value is still coerced, and the caller is still told: these are
	// integer types, so `holeSize: 42.5` is as much a correction as `holeSize: 200`.
	if (rounded !== clamped) warn(code, `${name} ${String(value)} must be a whole number; using ${rounded}.`)
	return rounded
}

/** {@link clampChartInt} for the percentage and angle options, which share one diagnostic. */
function clampChartPct(value: number | undefined, min: number, max: number, name: string): number | undefined {
	return clampChartInt(value, min, max, name)
}

/**
 * `<c:size val>` is ST_MarkerSize: an integer 2..72 points. Two copies of this clamp, with the
 * same bounds and the same message, sat 170 lines apart -- one for the chart-level option and
 * one for a combo subchart's override.
 */
function clampSymbolSize(value: number | undefined | null): number | undefined {
	return clampChartInt(value, 2, 72, 'lineDataSymbolSize', 'chart/symbol-size-out-of-range')
}

/**
 * Both vocabularies `radarStyle` accepts: `ST_RadarStyle`'s own members and this library's two
 * PowerPoint-UI aliases for them. Validated against the union and normalized to the wire member
 * by {@link normalizeRadarStyle}, so the emitter reads an `ST_RadarStyle` value and nothing else.
 */
const RADAR_STYLE_INPUTS = [...RADAR_STYLES, ...RADAR_STYLE_ALIAS_NAMES] as const

/**
 * Resolve `radarStyle` to its wire spelling, defaulting an absent or rejected one to `standard`
 * (the PowerPoint UI's plain "Radar"). The alias map is applied after the check rather than
 * before it, so a diagnostic names the value the caller actually wrote.
 */
function normalizeRadarStyle(options: ChartOptsOverrides): void {
	const stated = chartEnum(options.radarStyle, RADAR_STYLE_INPUTS, 'radarStyle')
	options.radarStyle = stated ? (RADAR_STYLE_ALIASES[stated as RadarStyleAlias] ?? stated) : 'standard'
}

/**
 * `ST_BarGrouping` minus `standard`, which PowerPoint does not offer on a 2-D bar plot.
 */
const BAR_GROUPINGS_2D = BAR_GROUPINGS.filter((g) => g !== 'standard')

/**
 * The `ST_DLblPos` members PowerPoint actually offers per plot type. The attribute's own value
 * space is `DATA_LABEL_POSITIONS`; these are the subsets of it a given plot accepts, and a value
 * outside its own plot's subset makes PowerPoint flag the file.
 */
const DATA_LABEL_POSITIONS_PIE = ['bestFit', 'ctr', 'inEnd', 'outEnd'] as const
/** Bubble, line and scatter: the four sides and the centre. */
const DATA_LABEL_POSITIONS_POINT = ['b', 'ctr', 'l', 'r', 't'] as const
/** A stacked or percent-stacked bar has no outside, so `outEnd` is not among its choices. */
const DATA_LABEL_POSITIONS_BAR_STACKED = ['ctr', 'inBase', 'inEnd'] as const
/** A clustered bar adds `outEnd` to the stacked set. */
const DATA_LABEL_POSITIONS_BAR_CLUSTERED = ['ctr', 'inBase', 'inEnd', 'outEnd'] as const

/**
 * Check one chart option against the values the library accepts for it, reporting and dropping
 * an unrecognized one so the caller's own default can stand.
 *
 * The definer hand-rolled about twenty of these as bare `Array.includes` tests with the list
 * written inline -- three of the lists twice, verbatim -- and every one was silent, while
 * `correctGridLineOptions` a few lines away reported the same class of mistake under its own
 * code. One wrapper puts them all on the reporting policy `check-enum.ts` describes.
 * @param value - the caller's value, if any
 * @param valid - the values this option accepts
 * @param option - option name as the caller spells it
 * @returns the value when accepted, else `null`
 */
function chartEnum<T extends string>(value: string | undefined, valid: readonly T[], option: string): T | null {
	return checkEnumOrWarn(value, valid, 'chart/invalid-option-value', `chart: ${option}`)
}

/**
 * Drop `dataLabelPosition` values that are invalid for the chart type / bar grouping,/**
 * Drop `dataLabelPosition` values that are invalid for the chart type / bar grouping,
 * per the OOXML data-label placement rules, so PowerPoint does not flag the file.
 *
 * `chartType` is passed in rather than read off `options._type` so the combo path can run the
 * same rules per subchart: a combo chart's `_type` is a `ChartMulti[]`, which matches none of
 * the comparisons below (see {@link normalizeComboSubchartOptions}).
 * @param options - options bag to correct in place
 * @param chartType - the plot type these options are emitted for, if known
 */
function normalizeChartDataLabelPosition(options: ChartOptsOverrides, chartType: ChartType | undefined): void {
	if (!options.dataLabelPosition) return
	const allowed = dataLabelPositionsFor(chartType, options.barGrouping)
	if (allowed === null) return
	if (allowed.length === 0 || !chartEnum(options.dataLabelPosition, allowed, 'dataLabelPosition')) {
		delete options.dataLabelPosition
	}
}

/**
 * The `ST_DLblPos` values one plot accepts, or `null` where this library states no rule.
 *
 * An empty list means the plot takes no position at all: `<c:dLblPos>` on an area, 3-D bar,
 * doughnut or radar chart is a PowerPoint repair prompt whatever its value.
 *
 * The bar row is the one that was wrong, and wrong in the direction that hurts: the list
 * *containing* `outEnd` was applied when the grouping was NOT clustered and the list without it
 * when the grouping was NOT stacked, so `{ barGrouping: 'clustered', dataLabelPosition:
 * 'outEnd' }` -- the most ordinary combination there is -- was silently deleted, while the
 * stacked form that PowerPoint itself refuses was kept. Settled against PowerPoint over COM:
 * setting `DataLabels.Position` to `xlLabelPositionOutsideEnd` on a clustered column chart is
 * accepted and reads back, and on a stacked or 100%-stacked one it raises 0x80004005, while
 * inside-end, inside-base and centre are accepted on all three.
 *
 * `barGrouping` is already defaulted by the time this runs -- {@link normalizeChartBarGrouping}
 * gives a 2-D bar `clustered` -- so there is no third arm for "no grouping stated".
 */
function dataLabelPositionsFor(
	chartType: ChartType | undefined,
	barGrouping: string | undefined
): readonly string[] | null {
	switch (chartType) {
		case ChartType.area:
		case ChartType.bar3d:
		case ChartType.doughnut:
		case ChartType.radar:
			return []
		case ChartType.pie:
			return DATA_LABEL_POSITIONS_PIE
		case ChartType.line:
		case ChartType.scatter:
			return DATA_LABEL_POSITIONS_POINT
		case ChartType.bar:
			return barGrouping === 'stacked' || barGrouping === 'percentStacked'
				? DATA_LABEL_POSITIONS_BAR_STACKED
				: DATA_LABEL_POSITIONS_BAR_CLUSTERED
		default:
			return isBubbleChart(chartType) ? DATA_LABEL_POSITIONS_POINT : null
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
function normalizeChartBarGrouping(options: ChartOptsOverrides, chartType: ChartType | undefined): void {
	// An area plot writes `<c:grouping>` (ST_Grouping), which has no clustered form; the two bar
	// plots write the `<c:grouping>` inside `c:barChart`/`c:bar3DChart`, which is ST_BarGrouping.
	// A 2-D bar additionally has no `standard`, PowerPoint offering clustered in its place.
	if (chartType === ChartType.area) {
		if (!chartEnum(options.barGrouping, GROUPINGS, 'barGrouping')) options.barGrouping = 'standard'
	}
	if (chartType === ChartType.bar) {
		if (!chartEnum(options.barGrouping, BAR_GROUPINGS_2D, 'barGrouping')) options.barGrouping = 'clustered'
	}
	if (chartType === ChartType.bar3d) {
		if (!chartEnum(options.barGrouping, BAR_GROUPINGS, 'barGrouping')) options.barGrouping = 'standard'
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
	// The three `*AxisLineShow` defaults that used to sit here are gone: they wrote `true` over
	// an absent flag, and the axis emitter now folds the flag into a stroke where only an
	// explicit `false` says anything (`type: 'none'`). Defaulting it was the last thing keeping
	// "the caller said nothing" and "the caller said yes" distinguishable, and nothing read the
	// distinction.

	options.v3DRotX =
		typeof options.v3DRotX === 'number' &&
		Number.isFinite(options.v3DRotX) &&
		options.v3DRotX >= -90 &&
		options.v3DRotX <= 90
			? options.v3DRotX
			: 30
	options.v3DRotY =
		typeof options.v3DRotY === 'number' &&
		Number.isFinite(options.v3DRotY) &&
		options.v3DRotY >= 0 &&
		options.v3DRotY <= 360
			? options.v3DRotY
			: 30
	// v3DRAngAx: same dead-ternary shape as the show* block above, same reason for its absence.
	options.v3DPerspective =
		typeof options.v3DPerspective === 'number' &&
		Number.isFinite(options.v3DPerspective) &&
		options.v3DPerspective >= 0 &&
		options.v3DPerspective <= 240
			? options.v3DPerspective
			: 30
}

/**
 * Apply chart-level option defaults: gap/overlap/hole clamps, chart colors, plotArea/chartArea
 * borders and fills, data border, data-label format codes, line size and multi-level cat labels.
 */
/** The five data-label font fields, which every per-series `<c:dLbls>` builder reads. */
const SERIES_LABEL_FONT_FIELDS = [
	'dataLabelColor',
	'dataLabelFontBold',
	'dataLabelFontFace',
	'dataLabelFontItalic',
	'dataLabelFontSize',
] as const

/**
 * Which {@link ChartSeriesOpts} fields each plot builder actually reads.
 *
 * The check this drives is per FIELD, not per chart type, because "the caller said it and nothing
 * happened" does not stop at the type boundary: `lineSize` reaches a stroke only where a series
 * draws one, so it was accepted and dropped on `bar`/`bar3d`/`area` even though those types read
 * every other field. A stock chart is the extreme of the same shape -- its price series draw no
 * line by design and its `<c:dLbls>` is a constant, so `color` is the only field with a referent.
 *
 * A type absent from this table supports nothing: a pie colours *points* rather than series and a
 * surface colours bands, so a per-series override has no referent on either even in principle.
 *
 * Keep each row in step with the builder named beside it. The rows are the only statement of what
 * is wired; nothing derives them from the emitters, so a field threaded through a builder without
 * being added here goes on warning that it does nothing.
 */
const SERIES_OPTION_FIELDS: ReadonlyMap<ChartType, ReadonlySet<keyof ChartSeriesOpts>> = new Map([
	// `makeCatAxisPlot`: `serShapeProps` + `serDataLabels`. `lineSize` reaches `seriesStroke` only
	// for the two line-like types; a bar or an area takes its outline from `dataBorder` instead.
	[ChartType.area, new Set(['color', ...SERIES_LABEL_FONT_FIELDS, 'dataLabelFormatCode'] as const)],
	[ChartType.bar, new Set(['color', ...SERIES_LABEL_FONT_FIELDS, 'dataLabelFormatCode'] as const)],
	[ChartType.bar3d, new Set(['color', ...SERIES_LABEL_FONT_FIELDS, 'dataLabelFormatCode'] as const)],
	[ChartType.line, new Set(['color', 'lineSize', ...SERIES_LABEL_FONT_FIELDS, 'dataLabelFormatCode'] as const)],
	[ChartType.radar, new Set(['color', 'lineSize', ...SERIES_LABEL_FONT_FIELDS, 'dataLabelFormatCode'] as const)],
	// `makeScatterPlot`: `scatterSerShapeProps`, plus `labelFontAttrs`/`labelFontChildren` inside
	// both label builders. Neither emits a `<c:numFmt>`, so `dataLabelFormatCode` has no referent.
	[ChartType.scatter, new Set(['color', 'lineSize', ...SERIES_LABEL_FONT_FIELDS] as const)],
	// `makeBubblePlot`: `bubbleSerShapeProps`. Its `<c:dLbls>` is one chart-level block, not one
	// per series, so no data-label field can be resolved per series.
	[ChartType.bubble, new Set(['color', 'lineSize'] as const)],
	[ChartType.bubble3d, new Set(['color', 'lineSize'] as const)],
	// `makeStockPlot`: the volume bar's fill and the close-series marker.
	[ChartType.stock, new Set(['color'] as const)],
])

function normalizeChartOptions(options: ChartOptsInternal): void {
	options.barGapWidthPct = clampChartPct(options.barGapWidthPct, 0, 500, 'barGapWidthPct') ?? 150
	options.barGapDepthPct = clampChartPct(options.barGapDepthPct, 0, 500, 'barGapDepthPct') ?? 150
	// These three have no default, so a value the clamp rejects has to *leave*, not become an
	// explicit `undefined`: this bag is spread over other bags (`gen/chart/chart-xml.ts` merges a
	// subchart's overrides onto it, and the axis builders merge `catAxes[n]`/`valAxes[n]` onto it),
	// where a present-but-undefined key overrides and an absent one inherits. `delete` is already
	// how the rest of this module spells "the caller's value did not survive" — see
	// `normalizeChartDataLabelPosition`.
	setOrClear(options, 'barOverlapPct', clampChartPct(options.barOverlapPct, -100, 100, 'barOverlapPct'))
	// `<c:holeSize>` is ST_HoleSize (10..90); `<c:firstSliceAng>` is ST_FirstSliceAng (0..360).
	setOrClear(options, 'holeSize', clampChartPct(options.holeSize, 10, 90, 'holeSize'))
	setOrClear(options, 'firstSliceAng', clampChartPct(options.firstSliceAng, 0, 360, 'firstSliceAng'))

	// An empty array is not a palette, so it means what saying nothing means: the built-in
	// default for this chart type. It used to survive this pass (`Array.isArray([])` is true)
	// and meet the plot builders' fallback instead, which was the *bar* palette on every type —
	// so `{ chartColors: [] }` on a pie was neither the caller's colours nor the pie default.
	options.chartColors = options.chartColors?.length ? options.chartColors : defaultChartPalette(options._type)
	// NaN is falsy, so this only ever has to answer for a value the caller actually set.
	// An out-of-range one reaches `percentToFixedPercent` at the emitter, which clamps and says so.
	if (!options.chartColorsOpacity) delete options.chartColorsOpacity
	options.plotArea = options.plotArea || {}
	if (!options.plotArea.border || typeof options.plotArea.border !== 'object') delete options.plotArea.border
	if (options.plotArea.border && !isUsableBorderWidth(options.plotArea.border.width))
		options.plotArea.border.width = DEF_CHART_BORDER.width
	if (
		options.plotArea.border &&
		(!options.plotArea.border.color || typeof options.plotArea.border.color !== 'string')
	) {
		options.plotArea.border.color = DEF_CHART_BORDER.color
	}
	options.plotArea.fill = options.plotArea.fill || {}
	options.chartArea = options.chartArea || {}
	if (!options.chartArea.border || typeof options.chartArea.border !== 'object') delete options.chartArea.border
	const chartAreaBorder = options.chartArea.border
	if (chartAreaBorder) {
		// Rebuilt rather than patched, and only the three keys the chart-area emitter reads survive
		// — `type` and `dashType` are dropped, as they always have been here. `transparency` is
		// copied only when the caller set one, so the rebuilt border spells "no transparency" the
		// same absent way the caller's own bag did.
		const border: BorderProps = {
			color: chartAreaBorder.color || DEF_CHART_BORDER.color,
			width: chartAreaBorder.width || DEF_CHART_BORDER.width,
		}
		if (chartAreaBorder.transparency !== undefined) border.transparency = chartAreaBorder.transparency
		options.chartArea.border = border
	}
	options.chartArea.roundedCorners =
		typeof options.chartArea.roundedCorners === 'boolean' ? options.chartArea.roundedCorners : true
	//
	if (!options.dataBorder || typeof options.dataBorder !== 'object') delete options.dataBorder
	if (options.dataBorder && !isUsableBorderWidth(options.dataBorder.width)) options.dataBorder.width = 0.75
	if (options.dataBorder && options.dataBorder.color) {
		const isHex = typeof options.dataBorder.color === 'string' && isHexColor(options.dataBorder.color)
		const isSchemeColor = Object.values(SchemeColor).includes(options.dataBorder.color as SchemeColor)
		if (!isHex && !isSchemeColor) {
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
	if (typeof options.valAxisMajorUnit !== 'number') delete options.valAxisMajorUnit
	if (typeof options.valAxisMinorUnit !== 'number') delete options.valAxisMinorUnit

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
		if (clean.length > 0) options.subtotals = clean
		else delete options.subtotals
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
	subOptions: ChartOptsOverrides | undefined,
	chartOptions: ChartOptsInternal,
	subType: ChartType,
	callerSetBarGapWidthPct: boolean
): ChartOptsOverrides {
	const sub: ChartOptsOverrides = subOptions && typeof subOptions === 'object' ? subOptions : {}
	// What the emitter reads for this subchart today, and the corrected copy to diff against it.
	// Both are `ChartOptsOverrides` because a correction that *rejects* a value records that as a
	// present `undefined`, which is the state the write-back below has to be able to carry.
	const merged: ChartOptsOverrides = { ...chartOptions, ...sub }
	const fixed: ChartOptsOverrides = { ...merged }

	// Enumerations emitted verbatim: `<c:barDir>` (ST_BarDir), `<c:grouping>` (ST_Grouping),
	// `<c:shape>` (ST_Shape), `<c:symbol>` (ST_MarkerStyle).
	if (!chartEnum(fixed.barDir, BAR_DIRECTIONS, 'barDir')) fixed.barDir = 'col'
	normalizeChartBarGrouping(fixed, subType)
	if (!chartEnum(fixed.bar3DShape, BAR_3D_SHAPES, 'bar3DShape')) fixed.bar3DShape = 'box'
	if (!chartEnum(fixed.lineDataSymbol, LINE_DATA_SYMBOLS, 'lineDataSymbol')) fixed.lineDataSymbol = 'circle'
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
	setOrClear(fixed, 'lineDataSymbolSize', clampSymbolSize(fixed.lineDataSymbolSize))
	// Points -> EMU, but only for a width this subchart supplied: the chart-level value has
	// already been converted and doing it twice would emit a hairline.
	if (sub.lineDataSymbolLineSize != null) fixed.lineDataSymbolLineSize = lineWidthToEmu(sub.lineDataSymbolLineSize)

	const result: ChartOptsOverrides = { ...sub }
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
	// Placeholder part identity, unique only within this target. The authoritative,
	// package-unique chart part filename is assigned at write time by a per-presentation
	// pass in `buildPackageParts` (`package/assemble.ts`, STEP 2). A module-global counter
	// here was never reset, so two identical decks built in one process emitted different
	// chart part filenames — same input, different bytes.
	const chartId = target._relsChart.length + 1
	const resultObject: SlideObject = {
		_type: SlideObjectType.chart,
	}
	// DESIGN: `type` can an object (ex: `ChartType.doughnut`) or an array of chart objects
	// EX: addChartDefinition([ { type:ChartType.bar, data:{name:'', labels:[], values[]} }, {<etc>} ])
	// Multi-Type Charts
	let tmpOpt: ChartOpts | ChartOptsInternal | undefined
	let tmpData: OptsChartDataInternal[] = []
	let tmpTypes: ChartMultiInternal[] | undefined
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
	const callerSetBarGapWidthPct = typeof options.barGapWidthPct === 'number' && !Number.isNaN(options.barGapWidthPct)

	// STEP 1: Set default options/decode user options
	// A: Core
	// The `type` is checked against the catalog here, at the boundary, because neither emitter can:
	// `makeChartType` and `chartExLayoutId` each own half of `ChartType` and treat the other half as
	// not theirs, so an off-catalog string matches no arm anywhere and would otherwise reach the deck
	// as a chart frame with no plot inside it.
	for (const name of tmpTypes ? tmpTypes.map((sub) => sub.type) : [type as CHART_NAME]) {
		if (!isChartType(name)) {
			throw new InvalidOptionError(
				'chart/unknown-type',
				`addChart: "${String(name)}" is not a chart type. Valid types are: ${Object.values(ChartType).join(', ')}.`,
				{ detail: { type: name } }
			)
		}
	}
	options._type = tmpTypes ?? asChartType(type as CHART_NAME)
	// Default only what the caller omitted. The guard used to be `!isNaN(Number(x))`, which is
	// false for every `Coord` that is not a bare number — so `x: '50%'` and `x: '2in'` were
	// thrown away and replaced by 1 inch, and a `NaN` was too, silently. `getSmartParseNumber`
	// is what vets a coordinate, and it reports a bad one instead of guessing.
	options.x = options.x ?? 1
	options.y = options.y ?? 1
	options.w = options.w || '50%'
	options.h = options.h || '50%'
	// Was the one definer still counting `_slideObjects` for its default name, which numbers a
	// chart by how many charts are *currently* on the slide rather than by how many have been
	// added — the difference `nextObjectNameIdx` exists for.
	options.objectName = resolveObjectName(target, SlideObjectType.chart, {
		label: 'Chart',
		kind: 'chart',
		supplied: options.objectName,
	})

	// B: Options: misc
	if (!chartEnum(options.barDir, BAR_DIRECTIONS, 'barDir')) options.barDir = 'col'

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
	if (!chartEnum(options.legendPos, LEGEND_POSITIONS, 'legendPos')) options.legendPos = 'r'

	if (!chartEnum(options.bar3DShape, BAR_3D_SHAPES, 'bar3DShape')) options.bar3DShape = 'box'
	if (!chartEnum(options.lineDataSymbol, LINE_DATA_SYMBOLS, 'lineDataSymbol')) options.lineDataSymbol = 'circle'
	if (!chartEnum(options.displayBlanksAs, DISPLAY_BLANKS_AS, 'displayBlanksAs')) options.displayBlanksAs = 'gap'
	normalizeRadarStyle(options)
	// Marker size emits as `<c:size val>` (ST_MarkerSize): an integer in [2,72] points.
	// Out-of-range or non-integer values make PowerPoint report the file as needing
	// repair, so round and clamp into range and warn when the input is coerced.
	options.lineDataSymbolSize = clampSymbolSize(options.lineDataSymbolSize) ?? 6
	// `lineWidthToEmu` rather than `ptsToEmuLenient`: this is an `a:ln/@w`, so an out-of-range
	// width is a repair prompt, and collapsing one to zero would be a silent hairline instead.
	options.lineDataSymbolLineSize = options.lineDataSymbolLineSize
		? lineWidthToEmu(options.lineDataSymbolLineSize)
		: ptsToEmuLenient(0.75)
	// `layout` allows the override of PPT defaults to maximize space
	const chartLayout = options.layout
	if (chartLayout) {
		;(['x', 'y', 'w', 'h'] as const).forEach((key) => {
			const val = chartLayout[key]
			const numVal = Number(val)
			if (!Number.isFinite(numVal) || numVal < 0 || numVal > 1) {
				warn('chart/layout-out-of-range', 'chart.layout.' + key + ' can only be 0-1')
				delete chartLayout[key] // remove invalid value so that default will be used
			}
		})
	}

	// Set gridline defaults
	const scatterGrid: OptsChartGridLine = { color: 'D9D9D9', width: 1 }
	options.catGridLine =
		options.catGridLine || (options._type === ChartType.scatter ? { ...scatterGrid } : { type: 'none' })
	options.valGridLine = options.valGridLine || (options._type === ChartType.scatter ? { ...scatterGrid } : {})
	options.serGridLine =
		options.serGridLine || (options._type === ChartType.scatter ? { ...scatterGrid } : { type: 'none' })
	scrubGridLine(options.catGridLine)
	scrubGridLine(options.valGridLine)
	scrubGridLine(options.serGridLine)
	setOrClear(options, 'shadow', normalizeShadowOptions(options.shadow))

	// C: Options: plotArea
	normalizeChartPlotAreaOptions(options)

	// D: Options: chart
	// `<c:gapWidth>`/`<c:gapDepth>` are ST_GapAmount (integer 0..500); `<c:overlap>` is
	// ST_Overlap (integer -100..100). Out-of-range values trigger PowerPoint repair.
	normalizeChartOptions(options)

	// D.1: A stated `seriesOptions` field that no plot builder reads is the third state the option
	// rules forbid -- "the caller said it and nothing happened" -- so each one is named against the
	// type that will drop it. The check is per field because the gaps are per field: a `lineSize` on
	// a bar series and a `dataLabelFormatCode` on a scatter series are both accepted by the type and
	// reach nothing in the part.
	if (options.seriesOptions?.length) {
		const types = (Array.isArray(options._type) ? options._type.map((sub) => sub.type) : [options._type]).map((name) =>
			asChartType(name)
		)
		const stated = new Set(
			options.seriesOptions.flatMap((entry) =>
				entry ? (Object.keys(entry) as Array<keyof ChartSeriesOpts>).filter((key) => entry[key] !== undefined) : []
			)
		)
		for (const type of new Set(types)) {
			const supported = SERIES_OPTION_FIELDS.get(type)
			const dropped = [...stated].filter((field) => !supported?.has(field)).sort()
			if (dropped.length > 0) {
				warn(
					'chart/option-not-supported',
					`"seriesOptions" ${dropped.map((field) => `\`${field}\``).join(', ')} ${dropped.length === 1 ? 'has' : 'have'} no effect on ${type}; ${supported ? `that plot reads ${[...supported].sort().join('/')} only` : 'that plot colours points or bands rather than series'}. Style them through the chart-level options.`
				)
			}
		}
	}

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
		// The same table the plot builder lays the series out from — a second copy here is how the
		// warning and the emitted chart came to disagree about what a style expects.
		if (!Object.keys(STOCK_STYLE_SPEC).includes(options.stockStyle || '')) options.stockStyle = 'hlc'
		const expected = STOCK_STYLE_SPEC[options.stockStyle as StockStyle].seriesCount
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
	// authoritative, package-unique filename is (re)assigned at write time in
	// `buildPackageParts` (`package/assemble.ts`); this placeholder mirrors the same
	// Ex-prefix rule so a single-chart deck is already correct.
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
