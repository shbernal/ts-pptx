/**
 * ts-pptx: Category-Axis Plot Assembly
 *
 * Emits the `<c:areaChart>` / `<c:barChart>` / `<c:bar3DChart>` / `<c:lineChart>` /
 * `<c:radarChart>` plot elements. These five chart types share one builder because they
 * share a plot shape -- grouping, per-series `<c:ser>` with a category reference and a
 * value cache, then the axis-id pair -- and differ only in a handful of type-gated
 * children. Reached through {@link ./chart-xml}'s `makeChartType` dispatch.
 */

import { ChartType } from '../../enums.js'
import { AXIS_ID_SERIES_PRIMARY, BARCHART_COLORS } from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { createLineCap } from '../drawingml/line.js'
import {
	categoryRange,
	dataLabels,
	dataValues,
	firstLabelGroup,
	seriesColumn,
	sheetCellRef,
	type SheetLayout,
	sheetRangeRef,
} from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { xsdBool } from '../../ooxml/xsd-boolean.js'
import { RADAR_STYLE_ALIASES, type RadarStyleAlias } from '../../ooxml/st-enums.js'
import {
	catRefBlock,
	chartDataLabels,
	createDataBorderLine,
	createSerLinesElement,
	dataLabelDefRPr,
	dLblNumFmt,
	dLblsBlock,
	dLblShowFlags,
	labelTextProps,
	makeChartErrorBarsXml,
	makeCustomDLblXml,
	makeSeriesDataPointsXml,
	numRefBlock,
	paletteColor,
	type PlotBuilder,
	resolveChartPalette,
	seriesShapeProps,
	seriesStroke,
	serMarker,
	strRefBlock,
} from './chart-parts.js'

/** The chart types drawn as a line (a marker per point, an optional smooth) rather than as a body. */
const isLineLike = (chartType: ChartType): boolean => chartType === ChartType.line || chartType === ChartType.radar

/** The two bar families, which share `<c:barDir>`, `<c:invertIfNegative>` and a gap width. */
const isBarLike = (chartType: ChartType): boolean => chartType === ChartType.bar || chartType === ChartType.bar3d

/** The type-gated children that open a plot element, before `<c:varyColors>`. */
function plotGrouping(chartType: ChartType, opts: ChartOptsInternal): string {
	if (chartType === ChartType.area || chartType === ChartType.line) {
		// CT_Line/AreaChart take ST_Grouping, which has no `clustered`; anything not stacked is standard.
		const lineGrouping =
			opts.barGrouping === 'stacked' || opts.barGrouping === 'percentStacked' ? opts.barGrouping : 'standard'
		return voidEl('c:grouping', { val: lineGrouping })
	}
	if (isBarLike(chartType)) {
		return voidEl('c:barDir', { val: opts.barDir }) + voidEl('c:grouping', { val: opts.barGrouping || 'clustered' })
	}
	if (chartType === ChartType.radar) {
		// `normalizeChartOptions` has already resolved this to an `ST_RadarStyle` member; the alias
		// map is applied again here only because it is the one that types the value, and it is
		// idempotent (no wire member is also an alias key).
		const radarStyle = opts.radarStyle ?? 'standard'
		return voidEl('c:radarStyle', { val: RADAR_STYLE_ALIASES[radarStyle as RadarStyleAlias] ?? radarStyle })
	}
	return ''
}

/** Fill, outline and shadow for one series, from its resolved colour and the line/border options. */
function serShapeProps(
	chartType: ChartType,
	opts: ChartOptsInternal,
	seriesColor: string,
	lineSize: number | undefined,
	serIndex: number
): string {
	let line = ''
	if (isLineLike(chartType)) {
		line = seriesStroke(opts, seriesColor, serIndex, lineSize)
	} else if (opts.dataBorder) {
		line = createDataBorderLine(opts.dataBorder, createLineCap(opts.lineCap))
	}
	return seriesShapeProps(opts, seriesColor, line)
}

/** Per-series `<c:dLbls>`: number format, an optional label background, text style, show flags. */
function serDataLabels(obj: OptsChartDataInternal, opts: ChartOptsInternal, seriesColor: string): string {
	const over = opts.seriesOptions?.[obj._dataIndex]
	const defRPr = dataLabelDefRPr(opts, over)
	const txPr = labelTextProps(defRPr)
	const lblFmtCode = over?.dataLabelFormatCode ?? opts.dataLabelFormatCode
	return dLblsBlock({
		// Per-point custom labels precede the aggregate settings (CT_DLbls order: dLbl* then Group_DLbls).
		lead: (obj.customLabels ?? []).map((lbl, idx) => (lbl ? makeCustomDLblXml(idx, lbl, opts) : '')).join(''),
		numFmt: dLblNumFmt(lblFmtCode),
		spPr: opts.dataLabelBkgrdColors ? el('c:spPr', null, raw(genXmlColorSelection(seriesColor))) : undefined,
		txPr,
		dLblPos: opts.dataLabelPosition ? voidEl('c:dLblPos', { val: opts.dataLabelPosition }) : undefined,
		flags: dLblShowFlags({ val: xsdBool(opts.showValue), serName: xsdBool(opts.showSerName) }),
		showLeaderLines: voidEl('c:showLeaderLines', { val: xsdBool(opts.showLeaderLines) }),
	})
}

/**
 * The `<c:cat>` category reference. Numeric (date) categories and single-level text ones are
 * {@link catRefBlock}; multi-level labels nest a `<c:lvl>` per group, which is a different shape
 * rather than a different spelling of the same one.
 */
function serCategories(obj: OptsChartDataInternal, opts: ChartOptsInternal): string {
	const groups = dataLabels(obj)
	const cats = firstLabelGroup(obj)
	// A series with no labels of its own states no categories. It used to fall through to the
	// multi-level arm with an empty group, which wrote `Sheet1!$A$2:$$1` -- a reference that
	// resolves nowhere. `<c:cat>` is optional in `CT_Ser`; the same arm `plot-pie.ts` takes.
	if (groups.length === 0 || cats.length === 0) return ''
	const catRef = categoryRange(cats.length)
	if (opts.catLabelFormatCode) {
		// A `catLabelFormatCode` implies numbers, so the cache is a numRef carrying that format.
		return el('c:cat', null, raw(catRefBlock('num', catRef, cats, opts.catLabelFormatCode || 'General')))
	}
	if (groups.length === 1) return el('c:cat', null, raw(catRefBlock('str', catRef, cats)))
	const lvls = groups
		.map((labelsGroup) =>
			el('c:lvl', null, raw(labelsGroup.map((label, idx) => el('c:pt', { idx }, raw(el('c:v', null, label)))).join('')))
		)
		.join('')
	const cache = el('c:multiLvlStrCache', null, [raw(voidEl('c:ptCount', { val: cats.length })), raw(lvls)])
	const ref = el('c:multiLvlStrRef', null, [
		raw(el('c:f', null, sheetRangeRef(1, 2, groups.length, cats.length + 1))),
		raw(cache),
	])
	return el('c:cat', null, raw(ref))
}

/** The `<c:val>` numeric cache: the series' own sheet column, one point per category. */
function serValues(obj: OptsChartDataInternal, valFmtCode: string, sheet: SheetLayout): string {
	const valCol = seriesColumn(obj, sheet)
	// The sheet's row count, not this series' own label count: the workbook writes one row per
	// category of the FIRST series and fills every series column across it, so a series with no
	// labels of its own would otherwise take a range that runs backwards.
	const rows = sheet.rowCount
	return numRefBlock('c:val', sheetRangeRef(valCol, 2, valCol, rows + 1), valFmtCode, dataValues(obj), rows)
}

/**
 * Plot a category-axis chart family (area / bar / bar3d / line / radar) into a
 * `<c:xxxChart>` element. These share the grouping / series / cat+val axis structure.
 */
export const makeCatAxisPlot: PlotBuilder = (chartType, data, opts, valAxisId, catAxisId, valFmtCode, sheet) => {
	/* EX1:
				data: [
				 {
				   name: 'Region 1',
				   labels: [['April', 'May', 'June', 'July']],
				   values: [17, 26, 53, 96]
				 },
				 {
				   name: 'Region 2',
				   labels: [['April', 'May', 'June', 'July']],
				   values: [55, 43, 70, 58]
				 }
				]
            */
	/* EX2:
				data: [
				 {
				   name: 'Region 1',
				   labels: [
					   ['April', 'May', 'June', 'April', 'May', 'June'],
					   ['2020',     '',     '', '2021',     '',     '']
				   ],
				   values: [17, 26, 53, 96, 40, 33]
				 },
				 {
				   name: 'Region 2',
				   labels: [
					   ['April', 'May', 'June', 'April', 'May', 'June'],
					   ['2020',     '',     '', '2021',     '',     '']
				   ],
				   values: [55, 43, 70, 58, 78, 63]
				 }
				]
             */
	// `chartColors` is always populated by addChartDefinition() (defaulting to BARCHART_COLORS); the
	// fallback in `paletteColor` only satisfies the optional type and keeps a colour a non-null string.
	const chartColors = resolveChartPalette(opts)
	// A single bar series varies its own colours when the caller asked for a palette other than the
	// default one, or supplied `invertedColors`; every other shape colours by series.
	const barVaryColors =
		isBarLike(chartType) &&
		data.length === 1 &&
		((opts.chartColors && opts.chartColors !== BARCHART_COLORS && opts.chartColors.length > 1) ||
			opts.invertedColors?.length)
			? opts.chartColors || BARCHART_COLORS
			: null

	// One `<c:ser>` per data row.
	const sers = data
		.map((obj) => {
			// EVERY per-series lookup here keys on `_dataIndex`, the series' position across the whole
			// chart -- the same number written to `<c:idx>`/`<c:order>` just below. For a single-type
			// chart that is also its position within this subchart, so the two only part company on a
			// combo, where each subchart's own loop restarts at 0.
			//
			// The palette used to key on that per-subchart position, so a bar(2) + line(1) combo painted
			// the line with palette entry 0 -- the first bar's colour. PowerPoint does not restart the
			// cycle per plot group: a three-series clustered-column chart whose third series is switched
			// to a line keeps that series' third colour and merely moves it from the fill to the stroke
			// (read back over COM; `test/regression/chart/combo-charts.test.js` carries the case).
			// `lineDashValues` moved with it, for the same reason and by the same reading of "the series
			// order in the `data` array".
			const seriesOverride = opts.seriesOptions?.[obj._dataIndex]
			const seriesColor = seriesOverride?.color ?? paletteColor(chartColors, obj._dataIndex)
			return el('c:ser', null, [
				raw(voidEl('c:idx', { val: obj._dataIndex })),
				raw(voidEl('c:order', { val: obj._dataIndex })),
				raw(strRefBlock(sheetCellRef(seriesColumn(obj, sheet), 1), obj.name ?? '')),
				raw(serShapeProps(chartType, opts, seriesColor, seriesOverride?.lineSize, obj._dataIndex)),
				// `invertIfNegative` is bar-only in the schema (CT_BarSer); area/line/radar must omit it.
				isBarLike(chartType) ? raw(voidEl('c:invertIfNegative', { val: 0 })) : null,
				// The marker takes the series' own colour, not a second palette lookup — the two used to
				// read different entries, so bar(2) + line(1) drew a `C0504D` line with `9BBB59` dots and
				// a `seriesOptions.color` override moved the body and not the dots.
				isLineLike(chartType) ? raw(serMarker(opts, seriesColor, seriesColor)) : null,
				// Per-point data points (`c:dPt`) MUST precede `c:dLbls` in CT_*Ser schema order.
				raw(makeSeriesDataPointsXml(chartType, obj, opts, barVaryColors)),
				// NOTE: [20190117] Adding data labels to a RADAR chart causes unrecoverable corruption,
				// and CT_RadarSer has no error bars either.
				chartType === ChartType.radar ? null : raw(serDataLabels(obj, opts, seriesColor)),
				chartType === ChartType.radar ? null : raw(makeChartErrorBarsXml(chartType, obj.errorBars, obj)),
				raw(serCategories(obj, opts)),
				raw(serValues(obj, valFmtCode, sheet)),
				chartType === ChartType.line ? raw(voidEl('c:smooth', { val: xsdBool(opts.lineSmooth) })) : null,
			])
		})
		.join('')

	// Per-type plot options. Schema order (CT_BarChart): gapWidth → overlap → serLines → axId.
	let plotOptions = ''
	if (chartType === ChartType.bar) {
		plotOptions =
			voidEl('c:gapWidth', { val: opts.barGapWidthPct }) +
			voidEl('c:overlap', { val: opts.barOverlapPct ?? ((opts.barGrouping || '').includes('tacked') ? 100 : 0) }) +
			// `<c:serLines>` connects data points across stacked bar/column series.
			createSerLinesElement(opts.barSeriesLine)
	} else if (chartType === ChartType.bar3d) {
		plotOptions =
			voidEl('c:gapWidth', { val: opts.barGapWidthPct }) +
			voidEl('c:gapDepth', { val: opts.barGapDepthPct }) +
			voidEl('c:shape', { val: opts.bar3DShape })
	} else if (chartType === ChartType.line) {
		plotOptions = voidEl('c:marker', { val: 1 })
	}

	return el(`c:${chartType}Chart`, null, [
		raw(plotGrouping(chartType, opts)),
		raw(voidEl('c:varyColors', { val: 0 })),
		raw(sers),
		raw(chartDataLabels(opts, true)),
		raw(plotOptions),
		// Axis id order matters: category comes first. Only 3-D charts get a series axis — emitting a
		// SERIES_PRIMARY axId for a 2-D chart left a dangling reference, and every axId in
		// `<c:plotArea>` must resolve to a defined catAx/valAx.
		raw(voidEl('c:axId', { val: catAxisId })),
		raw(voidEl('c:axId', { val: valAxisId })),
		chartType === ChartType.bar3d ? raw(voidEl('c:axId', { val: AXIS_ID_SERIES_PRIMARY })) : null,
	])
}
