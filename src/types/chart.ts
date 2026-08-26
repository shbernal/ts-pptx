/**
 * Chart types: series data (`OptsChartData`), per-axis and per-chart-family options, and the
 * combined `ChartOpts` surface.
 *
 * Re-exported by `./index.js`, which is the import site for the rest of `src/`.
 */
import type { CHART_NAME } from '../enums.js'
import type { Color, HexColor, PatternFillProps, PositionProps } from './core.js'
import type { ObjectNameProps } from './object.js'
import type { BorderProps, ShadowProps, ShapeFillProps } from './style.js'
import type { TextBaseProps } from './text.js'

export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
/**
 * Line end cap style. Maps to the OOXML `cap` attribute on `<a:ln>` (`flat`/`sq`/`rnd`).
 */
export type LineCap = 'flat' | 'round' | 'square'
export type ChartLineDash =
	| 'dash'
	| 'dashDot'
	| 'lgDash'
	| 'lgDashDot'
	| 'lgDashDotDot'
	| 'solid'
	| 'sysDash'
	| 'sysDot'

export interface OptsChartData {
	/**
	 * category labels
	 * @example ['Year 2000', 'Year 2010', 'Year 2020'] // single-level category axes labels
	 * @example [['Year 2000', 'Year 2010', 'Year 2020'], ['Decades', '', '']] // multi-level category axes labels
	 */
	labels?: string[] | string[][]
	/**
	 * series name
	 * @example 'Locations'
	 */
	name?: string
	/**
	 * bubble sizes
	 * @example [5, 1, 5, 1]
	 */
	sizes?: number[]
	/**
	 * category values
	 * @example [2000, 2010, 2020]
	 */
	values?: number[]
	/**
	 * Custom text label per data point, replacing the auto-generated value label.
	 * Index aligns with `values[]`. Empty string or missing entries fall back to the chart-level label settings.
	 * Supported for BAR, LINE, AREA, RADAR, PIE, and DOUGHNUT chart types.
	 * @example ['Low', '', 'High']  // only points 0 and 2 get custom labels
	 */
	customLabels?: string[]
	/**
	 * Per-data-point visual overrides (border / fill), index-aligned with `values[]`.
	 * Empty (`{}`) or missing entries fall back to series/chart styling.
	 * Supported for BAR, LINE, AREA, SCATTER, PIE, and DOUGHNUT chart types.
	 * @example
	 * pointStyles: [
	 *   { border: { width: 2, color: 'FF0000' } }, // point 0: red 2pt border
	 *   {},                                     // point 1: default
	 *   { fill: '00B050', border: { type: 'dash', color: '404040' } }, // point 2
	 * ]
	 */
	pointStyles?: ChartDataPointStyle[]
	/**
	 * Error bars for this series (`<c:errBars>`).
	 * - Supported for BAR, BAR3D, LINE, AREA, and SCATTER chart types (RADAR has no error bars in the schema).
	 * - Pass a single config, or an array to draw both X and Y error bars (SCATTER/AREA only; BAR/LINE use the first entry).
	 * @example { valueType: 'percentage', value: 5 } // ±5% error bars
	 * @example { valueType: 'fixedVal', value: 2, barType: 'plus', noEndCap: true }
	 * @example { valueType: 'cust', plusValues: [1, 2, 1], minusValues: [0.5, 1, 0.5] }
	 */
	errorBars?: ChartErrorBarOptions | ChartErrorBarOptions[]
}
/**
 * Per-data-point style override for a chart series.
 * Each entry applies to the data point at the same index in `values[]`.
 * Unset fields fall back to the series/chart-level styling.
 */
export interface ChartDataPointStyle {
	/**
	 * Data-point border (line). Reuses {@link BorderProps}.
	 * - `type: 'none'` hides the border; `'dash'` draws a dashed border.
	 * @example { width: 2, color: 'FF0000' }
	 */
	border?: BorderProps
	/**
	 * Data-point fill color (hex), overriding `chartColors[idx]`.
	 * Most meaningful on fill-based charts (BAR, AREA, PIE, DOUGHNUT).
	 * @example '00B050'
	 */
	fill?: HexColor
	/**
	 * Data-point pattern fill (`<a:pattFill>`), e.g. diagonal hatching, for the
	 * BAR/BAR3D and SCATTER charts that emit per-point `c:dPt`. Takes precedence
	 * over `fill` (OOXML allows only one fill per data point).
	 *
	 * When `pattern.fgColor` is omitted it defaults to this point's resolved fill
	 * color (`fill` or the varied `chartColors[idx]`), giving a hatched version of
	 * the bar color; if no point color is resolvable it falls back to black.
	 * `pattern.bgColor` defaults to white.
	 * @example { preset: 'ltUpDiag' }
	 * @example { preset: 'diagCross', fgColor: 'C00000', bgColor: 'FFFFFF' }
	 */
	pattern?: PatternFillProps
}
/**
 * Error-bar configuration for a chart series (`<c:errBars>`).
 * Maps onto OOXML `CT_ErrBars` (errDir / errBarType / errValType / noEndCap / plus / minus / val).
 */
export interface ChartErrorBarOptions {
	/**
	 * Axis the error bars measure along.
	 * - `'y'` (the value axis) for BAR/BAR3D/LINE/AREA; SCATTER may also use `'x'`.
	 * @default 'y'
	 */
	direction?: 'x' | 'y'
	/**
	 * Which sides of each marker draw a bar.
	 * @default 'both'
	 */
	barType?: 'both' | 'minus' | 'plus'
	/**
	 * How `value` (or `plusValues`/`minusValues`) is interpreted.
	 * - `'fixedVal'` — fixed amount in axis units
	 * - `'percentage'` — percent of each value (e.g. `value: 5` → ±5%)
	 * - `'stdDev'` — `value` standard deviations
	 * - `'stdErr'` — standard error (ignores `value`)
	 * - `'cust'` — explicit per-point amounts via `plusValues`/`minusValues`
	 * @default 'fixedVal'
	 */
	valueType?: 'cust' | 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr'
	/**
	 * Magnitude for `'fixedVal'`, `'percentage'`, or `'stdDev'`. Ignored for `'stdErr'` and `'cust'`.
	 * @default 1
	 */
	value?: number
	/** Per-point positive magnitudes; required when `valueType === 'cust'` (unless `barType: 'minus'`). Index-aligned with `values[]`. */
	plusValues?: number[]
	/** Per-point negative magnitudes; required when `valueType === 'cust'` (unless `barType: 'plus'`). Index-aligned with `values[]`. */
	minusValues?: number[]
	/**
	 * Hide the perpendicular end caps.
	 * @default false
	 */
	noEndCap?: boolean
	/** Error-bar line color (hex, e.g. `'FF0000'`). */
	color?: HexColor
	/** Error-bar line width (points). */
	size?: number
}
export interface OptsChartGridLine {
	/**
	 * MS-PPT > Chart format > Format Major Gridlines > Line > Cap type
	 * - line cap type
	 * @default flat
	 */
	cap?: LineCap
	/**
	 * Gridline color
	 * - `HexColor` or `ThemeColor`
	 * @example 'FF3399'
	 * @example SchemeColor.accent1 // theme color Accent1
	 */
	color?: Color
	/**
	 * Gridline size (points)
	 */
	size?: number
	/**
	 * Gridline style
	 */
	style?: 'solid' | 'dash' | 'dot' | 'none'
}
export interface ChartMulti {
	type: CHART_NAME
	data: OptsChartData[]
	options: ChartOpts
}
export interface ChartPropsFillLine {
	/**
	 * PowerPoint: Format Chart Area/Plot > Border ["Line"]
	 * @example border: {color: 'FF0000', width: 1} // hex RGB color, 1 pt line
	 */
	border?: BorderProps
	/**
	 * PowerPoint: Format Chart Area/Plot Area > Fill
	 *
	 * Omitting `fill` leaves the area transparent (`<a:noFill/>`), as does a `fill` that
	 * states neither a `color` nor a `type` — `transparency` alone is not a fill, since
	 * there is no colour for the alpha to apply to.
	 *
	 * `type: 'image'` is not supported here and falls back to no fill with a warning: a
	 * blip fill needs a media relationship on the chart part, which nothing registers.
	 * @example fill: {color: '696969'} // hex RGB color value
	 * @example fill: {color: SchemeColor.background2} // Theme color value
	 * @example fill: {color: '696969', transparency: 50} // 50% transparent grey
	 * @example fill: {type: 'gradient', gradient: {kind: 'linear', angle: 90, stops: [{position: 0, color: '0088CC'}, {position: 100, color: 'FFFFFF'}]}}
	 * @example fill: {type: 'inherit'} // no fill child at all -- take the chart style's fill
	 */
	fill?: ShapeFillProps
}
export interface ChartAreaProps extends ChartPropsFillLine {
	/**
	 * Whether the chart area has rounded corners
	 * - only applies when either `fill` or `border` is used
	 * @default true
	 */
	roundedCorners?: boolean
}
export interface ChartPropsBase {
	/**
	 * Axis position
	 */
	axisPos?: 'b' | 'l' | 'r' | 't'
	/**
	 * Series colours, cycled: series (or, on a pie/doughnut, data point) `n` takes
	 * `chartColors[n % chartColors.length]`, so a shorter list repeats rather than running out.
	 *
	 * Omitting it, or passing an empty array, both mean the same thing — the built-in default
	 * palette for the chart type. Two spellings of "I named no colours"; there is no third state
	 * that suppresses the fills.
	 */
	chartColors?: HexColor[]
	/**
	 * opacity (0 - 100)
	 * @example 50 // 50% opaque
	 */
	chartColorsOpacity?: number
	dataBorder?: BorderProps
	displayBlanksAs?: 'gap' | 'span' | 'zero'
	invertedColors?: HexColor[]
	lang?: string
	layout?: PositionProps
	shadow?: ShadowProps
	/**
	 * Show each bubble's size value as a data label (bubble / bubble3D charts only).
	 * Has no effect on other chart types.
	 * @default false
	 */
	showBubbleSize?: boolean
	/**
	 * Enable data labels, with a meaning that is **overloaded per plot type** — it is not a
	 * plain duplicate of {@link ChartPropsBase.showValue}:
	 * - pie / doughnut: shows the **category name** (`c:showCatName`);
	 * - scatter: shows **both** the value and the category name;
	 * - chartEx (waterfall/funnel/…): OR-combined with `showValue` to show the **value** label.
	 *
	 * For the common "show the numeric value" case on standard plots, prefer `showValue`.
	 * @default false
	 */
	showLabel?: boolean
	showLeaderLines?: boolean
	/**
	 * Leader line color (pie/doughnut data labels). Requires `showLeaderLines: true`.
	 * When omitted, PowerPoint applies its automatic leader-line color.
	 * @example 'FF0000' // red leader lines
	 */
	leaderLineColor?: HexColor
	/**
	 * Leader line width, in points (pie/doughnut data labels). Requires `showLeaderLines: true`.
	 * @default 0.75
	 * @example 1.5
	 */
	leaderLineSize?: number
	/**
	 * @default false
	 */
	showLegend?: boolean
	/**
	 * @default false
	 */
	showPercent?: boolean
	/**
	 * @default false
	 */
	showSerName?: boolean
	/**
	 * @default false
	 */
	showTitle?: boolean
	/**
	 * @default false
	 */
	showValue?: boolean
	/**
	 * 3D Perspecitve
	 * - range: 0-120
	 * @default 30
	 */
	v3DPerspective?: number
	/**
	 * Right Angle Axes
	 * - Shows chart from first-person perspective
	 * - Overrides `v3DPerspective` when true
	 * - PowerPoint: Chart Options > 3-D Rotation
	 * @default false
	 */
	v3DRAngAx?: boolean
	/**
	 * X Rotation
	 * - PowerPoint: Chart Options > 3-D Rotation
	 * - range: 0-359.9
	 * @default 30
	 */
	v3DRotX?: number
	/**
	 * Y Rotation
	 * - range: 0-359.9
	 * @default 30
	 */
	v3DRotY?: number

	/**
	 * PowerPoint: Format Chart Area (Fill & Border/Line)
	 */
	chartArea?: ChartAreaProps
	/**
	 * PowerPoint: Format Plot Area (Fill & Border/Line)
	 */
	plotArea?: ChartPropsFillLine

	/**
	 * Per-series style overrides.
	 * Element at index N applies to the series at data[N].
	 * Missing indices or unset fields fall back to the chart-level option.
	 */
	seriesOptions?: ChartSeriesOpts[]
}
export interface ChartPropsAxisCat {
	/**
	 * Multi-Chart prop: array of cat axes
	 */
	catAxes?: ChartPropsAxisCat[]
	catAxisBaseTimeUnit?: string
	catAxisCrossesAt?: number | 'autoZero'
	catAxisHidden?: boolean
	catAxisLabelColor?: string
	catAxisLabelFontBold?: boolean
	catAxisLabelFontFace?: string
	catAxisLabelFontItalic?: boolean
	catAxisLabelFontSize?: number
	/**
	 * Number format code for the category (X) axis labels on scatter and bubble charts.
	 * Falls back to `valAxisLabelFormatCode` when not set.
	 * - Example: `'0.00'`, `'#,##0'`, `'mmm yyyy'`
	 * - PowerPoint: Format Axis > Number > Format Code
	 */
	catAxisLabelFormatCode?: string
	catAxisLabelFrequency?: string
	catAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
	catAxisLabelRotate?: number
	catAxisLineColor?: string
	catAxisLineShow?: boolean
	catAxisLineSize?: number
	catAxisLineStyle?: 'solid' | 'dash' | 'dot'
	catAxisMajorTickMark?: ChartAxisTickMark
	catAxisMajorTimeUnit?: string
	catAxisMajorUnit?: number
	catAxisMaxVal?: number
	catAxisMinorTickMark?: ChartAxisTickMark
	catAxisMinorTimeUnit?: string
	catAxisMinorUnit?: number
	catAxisMinVal?: number
	catAxisMultiLevelLabels?: boolean
	catAxisOrientation?: 'minMax' | 'maxMin'
	catAxisTitle?: string
	catAxisTitleColor?: string
	catAxisTitleFontFace?: string
	catAxisTitleFontSize?: number
	catAxisTitleRotate?: number
	catGridLine?: OptsChartGridLine
	catLabelFormatCode?: string
	/**
	 * Whether data should use secondary category axis (instead of primary)
	 * @default false
	 */
	secondaryCatAxis?: boolean
	showCatAxisTitle?: boolean
}
export interface ChartPropsAxisSer {
	serAxisBaseTimeUnit?: string
	serAxisHidden?: boolean
	serAxisLabelColor?: string
	serAxisLabelFontBold?: boolean
	serAxisLabelFontFace?: string
	serAxisLabelFontItalic?: boolean
	serAxisLabelFontSize?: number
	serAxisLabelFrequency?: string
	serAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
	serAxisLineColor?: string
	serAxisLineShow?: boolean
	serAxisMajorTimeUnit?: string
	serAxisMajorUnit?: number
	serAxisMinorTimeUnit?: string
	serAxisMinorUnit?: number
	serAxisOrientation?: string
	serAxisTitle?: string
	serAxisTitleColor?: string
	serAxisTitleFontFace?: string
	serAxisTitleFontSize?: number
	serAxisTitleRotate?: number
	serGridLine?: OptsChartGridLine
	serLabelFormatCode?: string
	showSerAxisTitle?: boolean
}
export interface ChartPropsAxisVal {
	/**
	 * Whether data should use secondary value axis (instead of primary)
	 * @default false
	 */
	secondaryValAxis?: boolean
	showValAxisTitle?: boolean
	/**
	 * Multi-Chart prop: array of val axes
	 */
	valAxes?: ChartPropsAxisVal[]
	valAxisCrossesAt?: number | 'autoZero'
	/**
	 * Controls where axis values are plotted relative to tick marks
	 * - `'between'` = values plotted between tick marks (default for bar/column/line)
	 * - `'midCat'` = values plotted on tick marks (default for scatter/area)
	 * - PowerPoint: Format Axis > Axis Options > Axis crosses > On tick marks / Between tick marks
	 */
	valAxisCrossBetween?: 'between' | 'midCat'
	valAxisDisplayUnit?:
		| 'billions'
		| 'hundredMillions'
		| 'hundreds'
		| 'hundredThousands'
		| 'millions'
		| 'tenMillions'
		| 'tenThousands'
		| 'thousands'
		| 'trillions'
	valAxisDisplayUnitLabel?: boolean
	valAxisHidden?: boolean
	valAxisLabelColor?: string
	valAxisLabelFontBold?: boolean
	valAxisLabelFontFace?: string
	valAxisLabelFontItalic?: boolean
	valAxisLabelFontSize?: number
	valAxisLabelFormatCode?: string
	valAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
	valAxisLabelRotate?: number
	valAxisLineColor?: string
	valAxisLineShow?: boolean
	valAxisLineSize?: number
	valAxisLineStyle?: 'solid' | 'dash' | 'dot'
	/**
	 * PowerPoint: Format Axis > Axis Options > Logarithmic scale - Base
	 * - range: 2-99
	 */
	valAxisLogScaleBase?: number
	valAxisMajorTickMark?: ChartAxisTickMark
	valAxisMajorUnit?: number
	valAxisMaxVal?: number
	valAxisMinorTickMark?: ChartAxisTickMark
	valAxisMinVal?: number
	valAxisOrientation?: 'minMax' | 'maxMin'
	valAxisTitle?: string
	valAxisTitleColor?: string
	valAxisTitleFontFace?: string
	valAxisTitleFontSize?: number
	valAxisTitleRotate?: number
	valGridLine?: OptsChartGridLine
	/**
	 * Value label format code
	 * - this also directs Data Table formatting
	 * @example '#%' // round percent
	 * @example '0.00%' // shows values as '0.00%'
	 * @example '$0.00' // shows values as '$0.00'
	 */
	valLabelFormatCode?: string
}
export interface ChartPropsChartBar {
	bar3DShape?: string
	barDir?: string
	barGapDepthPct?: number
	/**
	 * MS-PPT > Format chart > Format Data Point > Series Options >  "Gap Width"
	 * - width (percent)
	 * - range: `0`-`500`
	 * @default 150
	 */
	barGapWidthPct?: number
	barGrouping?: string
	/**
	 * MS-PPT > Format chart > Format Data Point > Series Options >  "Series Overlap"
	 * - overlap (percent)
	 * - range: `-100`-`100`
	 * @default 0
	 */
	barOverlapPct?: number
	/**
	 * Draw connector lines between data points across stacked bar/column series
	 * ("Series Lines" in PowerPoint). Emits `<c:serLines>` in the bar chart.
	 *
	 * - `true` uses PowerPoint's automatic line styling.
	 * - An {@link OptsChartGridLine} object customizes color/size/style/cap.
	 * - Omit (or pass an object with `style: 'none'`) to disable.
	 *
	 * Bar (`bar`) charts only; ignored for 3D bar charts.
	 * @default undefined
	 * @example true
	 * @example { color: '777777', size: 1, style: 'dash' }
	 */
	barSeriesLine?: boolean | OptsChartGridLine
}
export interface ChartPropsChartDoughnut {
	dataNoEffects?: boolean
	holeSize?: number
}
export interface ChartPropsChartLine {
	/**
	 * MS-PPT > Chart format > Format Data Series > Line > Cap type
	 * - line cap type
	 * @default flat
	 */
	lineCap?: LineCap
	/**
	 * MS-PPT > Chart format > Format Data Series > Line > Dash type (chart-level default)
	 * - applies to every series that has no entry in `lineDashValues`
	 * @default solid
	 */
	lineDash?: ChartLineDash
	/**
	 * Per-series dash type overrides; index matches the series order in the `data` array.
	 * - entries shorter than the series count fall back to `lineDash`
	 * - example: `['solid', 'dash', 'lgDash']` gives each series its own dash pattern
	 */
	lineDashValues?: ChartLineDash[]
	/**
	 * MS-PPT > Chart format > Format Data Series > Marker Options > Built-in > Type
	 * - marker type
	 * @default circle
	 */
	lineDataSymbol?: 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle'
	/**
	 * MS-PPT > Chart format > Format Data Series > [Marker Options] > Border > Color
	 * - border color
	 * @default circle
	 */
	lineDataSymbolLineColor?: string
	/**
	 * MS-PPT > Chart format > Format Data Series > [Marker Options] > Border > Width
	 * - border width (points)
	 * @default 0.75
	 */
	lineDataSymbolLineSize?: number
	/**
	 * MS-PPT > Chart format > Format Data Series > Marker Options > Built-in > Size
	 * - marker size
	 * - range: 2-72
	 * @default 6
	 */
	lineDataSymbolSize?: number
	/**
	 * MS-PPT > Chart format > Format Data Series > Line > Width
	 * - line width (points)
	 * - range: 0-1584
	 * @default 2
	 */
	lineSize?: number
	/**
	 * MS-PPT > Chart format > Format Data Series > Line > Smoothed line
	 * - "Smoothed line"
	 * @default false
	 */
	lineSmooth?: boolean
}
export interface ChartPropsChartPie {
	dataNoEffects?: boolean
	/**
	 * MS-PPT > Format chart > Format Data Series > Series Options >  "Angle of first slice"
	 * - angle (degrees)
	 * - range: 0-359
	 * @default 0
	 */
	firstSliceAng?: number
}
export interface ChartPropsChartRadar {
	/**
	 * Radar chart sub-type, named to match the PowerPoint UI ("Radar", "Radar with
	 * Markers", "Filled Radar").
	 * @default radar
	 */
	radarStyle?: 'radar' | 'markers' | 'filled'
}
/**
 * Options for the modern (Office 2016) chartEx chart types.
 *
 * NOTE: chartEx charts (currently `waterfall`) render only in PowerPoint 2016+/Microsoft 365.
 * In older Office, Google Slides, Keynote and LibreOffice a fallback placeholder shows instead.
 */
export interface ChartPropsChartWaterfall {
	/**
	 * Zero-based category indices to render as **subtotal/total** columns in a `waterfall` chart
	 * (columns anchored to the axis baseline rather than floating from the running cumulative
	 * value). Out-of-range or non-integer entries are dropped with a console warning.
	 * @example [0, 4] // first and fifth categories are totals
	 */
	subtotals?: number[]
}
/**
 * Bin configuration for a `histogram` chart. A histogram is fed **raw observations** (one series of
 * `values`, no `labels`); PowerPoint groups them into bins automatically. The only control wired
 * today is which side of each bin interval is closed.
 *
 * NOTE: explicit bin geometry (fixed bin width / bin count / overflow / underflow) is not yet
 * supported — PowerPoint desktop rejects the decks those settings produce, so they are deferred.
 */
export interface ChartExBinning {
	/** Which side of each bin interval is closed: `'r'` (right, the default) or `'l'` (left). */
	intervalClosed?: 'l' | 'r'
}
export interface ChartPropsChartHistogram {
	/**
	 * Bin configuration for a `histogram` chart. Omit for PowerPoint's automatic binning.
	 * @example { intervalClosed: 'l' }
	 */
	binning?: ChartExBinning
}
/**
 * Statistics configuration for a `boxWhisker` (box-and-whisker) chart. Each value series is
 * summarized into a box (quartiles + median) with whiskers; these flags mirror the toggles
 * PowerPoint exposes in the *Format Data Series* pane. All fields are optional and default to
 * PowerPoint's own defaults (inclusive of the outlier points and the mean marker).
 */
export interface ChartExStatistics {
	/**
	 * How the first/third quartiles are computed: `'exclusive'` (PowerPoint's default — the median
	 * is excluded when splitting the halves) or `'inclusive'`.
	 */
	quartileMethod?: 'inclusive' | 'exclusive'
	/** Draw a line connecting the mean of each box across categories. Default `false`. */
	meanLine?: boolean
	/** Mark the mean of each box with a marker. Default `true`. */
	meanMarker?: boolean
	/** Plot outlier points (values beyond the whiskers) individually. Default `true`. */
	outliers?: boolean
	/** Plot every non-outlier point alongside the box, not just the box itself. Default `false`. */
	nonoutliers?: boolean
}
export interface ChartPropsChartBoxWhisker {
	/**
	 * Statistics configuration for a `boxWhisker` chart. Omit for PowerPoint's defaults
	 * (exclusive quartiles, mean marker shown, outliers shown).
	 * @example { quartileMethod: 'inclusive', meanLine: true }
	 */
	statistics?: ChartExStatistics
}
/**
 * Geography configuration for a `regionMap` (filled-map) chart. A region map colors each geographic
 * area by its value; the category labels are region names (countries, states, provinces, …) that
 * PowerPoint resolves against its Bing-backed geography database when the deck is opened.
 *
 * These fields tell PowerPoint how to interpret ambiguous region names (e.g. "Georgia" the country
 * vs. the US state). They do NOT change what ts-pptx writes beyond the `<cx:geography>` hint —
 * the actual map is resolved and rendered by PowerPoint 2016+/Microsoft 365 at open time, so the
 * region names must match PowerPoint's geography database or the map renders blank.
 */
export interface ChartExGeography {
	/** BCP-47 language used to interpret the region names, e.g. `'en-US'` (the default). */
	cultureLanguage?: string
	/** ISO region the names are scoped to, e.g. `'US'` (the default) or `'FR'`. */
	cultureRegion?: string
}
export interface ChartPropsChartRegionMap {
	/**
	 * Geography configuration for a `regionMap` chart. Omit for PowerPoint's defaults
	 * (`cultureLanguage: 'en-US'`, `cultureRegion: 'US'`).
	 * @example { cultureLanguage: 'en-US', cultureRegion: 'FR' }
	 */
	geography?: ChartExGeography
}
export interface ChartPropsChartStock {
	/**
	 * Stock (high-low-close) chart style — selects how many value series the chart expects and how
	 * they are drawn. The data series must be supplied in the exact order implied by the style:
	 * - `'hlc'` — High, Low, Close (3 series). The close is marked with a dot.
	 * - `'ohlc'` — Open, High, Low, Close (4 series). Open→close is drawn as up/down bars.
	 * - `'vhlc'` — Volume, High, Low, Close (4 series). Volume is a column on its own axis.
	 * - `'vohlc'` — Volume, Open, High, Low, Close (5 series). Volume column + up/down bars.
	 *
	 * `<c:hiLowLines>` connect the high/low of each category in every style.
	 * @default 'hlc'
	 */
	stockStyle?: 'hlc' | 'ohlc' | 'vhlc' | 'vohlc'
}
export interface ChartPropsChartSurface {
	/**
	 * Draw a 3-D surface (`<c:surface3DChart>`) rather than a 2-D contour / top view
	 * (`<c:surfaceChart>`). The 3-D surface tilts the scene; the contour looks straight down.
	 * @default true
	 */
	surface3D?: boolean
	/**
	 * Draw the surface as a wireframe mesh (`<c:wireframe val="1"/>`) instead of a filled sheet.
	 * Applies to both the 3-D surface and the 2-D contour.
	 * @default false
	 */
	surfaceWireframe?: boolean
}
/**
 * Per-series style overrides for a chart.
 * Each entry applies to the series at the same index in the data array.
 * Unset fields fall back to the chart-level option.
 */
export interface ChartSeriesOpts {
	/** Series fill / line color (hex, e.g. `'FF0000'`) */
	color?: HexColor
	/** Data-label font color */
	dataLabelColor?: string
	/** Data-label font bold */
	dataLabelFontBold?: boolean
	/** Data-label typeface */
	dataLabelFontFace?: string
	/** Data-label font italic */
	dataLabelFontItalic?: boolean
	/** Data-label font size (points) */
	dataLabelFontSize?: number
	/**
	 * Data-label number format code for this series.
	 * Overrides the chart-level `dataLabelFormatCode` for this series only.
	 * @example '#,##0' // thousands separator
	 * @example '0.00%' // shows values as '0.00%'
	 * @example '$0.00' // shows values as '$0.00'
	 */
	dataLabelFormatCode?: string
	/**
	 * Line/radar series line width (points).
	 * Pass `0` to hide the line.
	 */
	lineSize?: number
}

export interface ChartPropsDataLabel {
	dataLabelBkgrdColors?: boolean
	dataLabelColor?: string
	dataLabelFontBold?: boolean
	dataLabelFontFace?: string
	dataLabelFontItalic?: boolean
	dataLabelFontSize?: number
	/**
	 * Data label format code
	 * @example '#%' // round percent
	 * @example '0.00%' // shows values as '0.00%'
	 * @example '$0.00' // shows values as '$0.00'
	 */
	dataLabelFormatCode?: string
	dataLabelFormatScatter?: 'custom' | 'customXY' | 'XY'
	dataLabelPosition?: 'b' | 'bestFit' | 'ctr' | 'l' | 'r' | 't' | 'inEnd' | 'outEnd'
}
export interface ChartPropsDataTable {
	dataTableFontSize?: number
	/**
	 * Data table format code
	 * @example '#%' // round percent
	 * @example '0.00%' // shows values as '0.00%'
	 * @example '$0.00' // shows values as '$0.00'
	 */
	dataTableFormatCode?: string
	/**
	 * Whether to show a data table adjacent to the chart
	 * @default false
	 */
	showDataTable?: boolean
	showDataTableHorzBorder?: boolean
	showDataTableKeys?: boolean
	showDataTableOutline?: boolean
	showDataTableVertBorder?: boolean
}
export interface ChartPropsLegend {
	legendColor?: string
	legendFontFace?: string
	legendFontSize?: number
	/**
	 * Manual legend placement within the chart area.
	 *
	 * Each of `x`/`y`/`w`/`h` is a fraction (0-1) of the chart's width/height.
	 * `x`/`y` position the legend's top-left corner relative to the chart edge;
	 * `w`/`h` set its size. Each axis is independent: provide only `x` to move the
	 * legend horizontally while leaving vertical placement and size automatic.
	 * Setting this overrides the automatic placement implied by `legendPos`.
	 *
	 * Has no effect unless `showLegend` is `true`.
	 *
	 * @example { x: 0.7, y: 0.3, w: 0.25, h: 0.4 }
	 */
	legendLayout?: PositionProps
	legendPos?: 'b' | 'l' | 'r' | 't' | 'tr'
}
export interface ChartPropsTitle extends TextBaseProps {
	title?: string
	titleAlign?: string
	titleBold?: boolean
	titleColor?: string
	titleFontFace?: string
	titleFontSize?: number
	titleItalic?: boolean
	titleUnderline?: boolean
	/**
	 * Manual title position (inches), relative to the chart.
	 * Each axis is independent: omit `x` to keep automatic horizontal centering,
	 * or omit `y` to keep automatic vertical placement. Provide at least one.
	 */
	titlePos?: { x?: number; y?: number }
	titleRotate?: number
}
export interface ChartOpts
	extends
		ChartPropsAxisCat,
		ChartPropsAxisSer,
		ChartPropsAxisVal,
		ChartPropsBase,
		ChartPropsChartBar,
		ChartPropsChartDoughnut,
		ChartPropsChartLine,
		ChartPropsChartPie,
		ChartPropsChartRadar,
		ChartPropsChartWaterfall,
		ChartPropsChartHistogram,
		ChartPropsChartBoxWhisker,
		ChartPropsChartRegionMap,
		ChartPropsChartStock,
		ChartPropsChartSurface,
		ChartPropsDataLabel,
		ChartPropsDataTable,
		ChartPropsLegend,
		ChartPropsTitle,
		ObjectNameProps,
		OptsChartGridLine,
		PositionProps {
	/**
	 * Chart type — required when using the canonical `addChart(data, options)` signature.
	 * - Omit only for multi-type (combo) charts, where each `ChartMulti` entry carries its own `type`.
	 */
	type?: CHART_NAME
	/**
	 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
	 * - PowerPoint: [right-click on a chart] > "Edit Alt Text..."
	 */
	altText?: string
	/**
	 * Custom chart-level metadata, emitted as a schema-valid extension on the chart space
	 * (`c:chartSpace/c:extLst`) under a stable ts-pptx vendor GUID.
	 * - Use for round-trippable, machine-readable annotations a consumer wants to travel with the
	 *   chart (e.g. a source-data id, a generator tag, a semantic role). PowerPoint preserves the
	 *   extension untouched and ignores it for rendering.
	 * - Keys must be non-empty strings; values must be strings. Invalid entries are dropped with a
	 *   console warning rather than emitting degenerate XML.
	 * @example { sourceId: 'q3-revenue', generator: 'my-deck-tool' }
	 */
	metadata?: Record<string, string>
}
