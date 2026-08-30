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
import {
	AXIS_ID_SERIES_PRIMARY,
	BARCHART_COLORS,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SHAPE_SHADOW,
} from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { createLineCap } from '../drawingml/line.js'
import { ptsToEmuLenient } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { dataLabels, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	catRefBlock,
	chartColorLineFill,
	chartDataLabels,
	createChartTextFonts,
	createDataBorderLine,
	createSerLinesElement,
	dLblShowFlags,
	makeChartErrorBarsXml,
	makeCustomDLblXml,
	makeSeriesDataPointsXml,
	numRefBlock,
	paletteColor,
	resolveChartPalette,
	seriesFill,
	strRefBlock,
	type PlotBuilder,
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
		// Map the public PowerPoint-UI names to ST_RadarStyle wire values.
		const radarStyleWire =
			{ radar: 'standard', markers: 'marker', filled: 'filled' }[opts.radarStyle || 'radar'] ?? 'standard'
		return voidEl('c:radarStyle', { val: radarStyleWire })
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
	const fill = seriesFill(opts, seriesColor)
	let line = ''
	if (isLineLike(chartType)) {
		const effectiveLineSize = lineSize ?? opts.lineSize ?? 2
		line =
			effectiveLineSize === 0
				? el('a:ln', null, raw(voidEl('a:noFill')))
				: el('a:ln', { w: ptsToEmuLenient(effectiveLineSize), cap: createLineCap(opts.lineCap) }, [
						raw(chartColorLineFill(seriesColor)),
						raw(voidEl('a:prstDash', { val: opts.lineDashValues?.[serIndex] ?? opts.lineDash ?? 'solid' })),
						raw(voidEl('a:round')),
					])
	} else if (opts.dataBorder) {
		line = createDataBorderLine(opts.dataBorder, createLineCap(opts.lineCap))
	}
	return el('c:spPr', null, [raw(fill), raw(line), raw(createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW))])
}

/**
 * The point marker for a line/radar series. `<c:marker>` must precede `<c:dLbls>` in CT_LineSer
 * (schema order: spPr → marker → dPt → dLbls).
 */
function serMarker(opts: ChartOptsInternal, markerColor: string, seriesColor: string): string {
	const spPr = el('c:spPr', null, [
		raw(markerColor === 'transparent' ? voidEl('a:noFill') : genXmlColorSelection(markerColor)),
		raw(
			el('a:ln', { w: opts.lineDataSymbolLineSize, cap: 'flat' }, [
				raw(chartColorLineFill(opts.lineDataSymbolLineColor || seriesColor)),
				raw(voidEl('a:prstDash', { val: 'solid' })),
				raw(voidEl('a:round')),
			])
		),
		raw(voidEl('a:effectLst', null)),
	])
	return el('c:marker', null, [
		raw(voidEl('c:symbol', { val: opts.lineDataSymbol })),
		// Defaults to "auto" otherwise (but this is usually too small, so there is a default).
		opts.lineDataSymbolSize ? raw(voidEl('c:size', { val: opts.lineDataSymbolSize })) : null,
		raw(spPr),
	])
}

/** Per-series `<c:dLbls>`: number format, an optional label background, text style, show flags. */
function serDataLabels(obj: OptsChartDataInternal, opts: ChartOptsInternal, seriesColor: string): string {
	const over = opts.seriesOptions?.[obj._dataIndex]
	const defRPr = el(
		'a:defRPr',
		{
			b: (over?.dataLabelFontBold ?? opts.dataLabelFontBold ?? false) ? 1 : 0,
			i: (over?.dataLabelFontItalic ?? opts.dataLabelFontItalic ?? false) ? 1 : 0,
			strike: 'noStrike',
			sz: ptToHundredths(over?.dataLabelFontSize ?? opts.dataLabelFontSize ?? DEF_FONT_SIZE),
			u: 'none',
		},
		[
			raw(genXmlColorSelection(over?.dataLabelColor ?? opts.dataLabelColor ?? DEF_FONT_COLOR)),
			raw(createChartTextFonts(over?.dataLabelFontFace ?? opts.dataLabelFontFace ?? 'Arial')),
		]
	)
	const txPr = el('c:txPr', null, [
		raw(voidEl('a:bodyPr')),
		raw(voidEl('a:lstStyle')),
		raw(el('a:p', null, raw(el('a:pPr', null, raw(defRPr))))),
	])
	const lblFmtCode = over?.dataLabelFormatCode ?? opts.dataLabelFormatCode
	return el('c:dLbls', null, [
		// Per-point custom labels precede the aggregate settings (CT_DLbls order: dLbl* then Group_DLbls).
		raw((obj.customLabels ?? []).map((lbl, idx) => (lbl ? makeCustomDLblXml(idx, lbl, opts) : '')).join('')),
		raw(voidEl('c:numFmt', { formatCode: (lblFmtCode ?? '') || 'General', sourceLinked: 0 })),
		opts.dataLabelBkgrdColors ? raw(el('c:spPr', null, raw(genXmlColorSelection(seriesColor)))) : null,
		raw(txPr),
		opts.dataLabelPosition ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition })) : null,
		...dLblShowFlags({ val: opts.showValue ? 1 : 0, serName: opts.showSerName ? 1 : 0 }),
		raw(voidEl('c:showLeaderLines', { val: opts.showLeaderLines ? 1 : 0 })),
	])
}

/**
 * The `<c:cat>` category reference. Numeric (date) categories and single-level text ones are
 * {@link catRefBlock}; multi-level labels nest a `<c:lvl>` per group, which is a different shape
 * rather than a different spelling of the same one.
 */
function serCategories(obj: OptsChartDataInternal, opts: ChartOptsInternal): string {
	const groups = dataLabels(obj)
	const cats = firstLabelGroup(obj)
	const catRef = `Sheet1!$A$2:$A$${cats.length + 1}`
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
function serValues(obj: OptsChartDataInternal, valFmtCode: string): string {
	const valCol = obj._dataIndex + dataLabels(obj).length + 1
	const catCount = firstLabelGroup(obj).length
	return numRefBlock('c:val', sheetRangeRef(valCol, 2, valCol, catCount + 1), valFmtCode, dataValues(obj), catCount)
}

/**
 * Plot a category-axis chart family (area / bar / bar3d / line / radar) into a
 * `<c:xxxChart>` element. These share the grouping / series / cat+val axis structure.
 */
export const makeCatAxisPlot: PlotBuilder = (chartType, data, opts, valAxisId, catAxisId, valFmtCode) => {
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
		.map((obj, serIndex) => {
			const seriesOverride = opts.seriesOptions?.[obj._dataIndex]
			const seriesColor = seriesOverride?.color ?? paletteColor(chartColors, serIndex)
			return el('c:ser', null, [
				raw(voidEl('c:idx', { val: obj._dataIndex })),
				raw(voidEl('c:order', { val: obj._dataIndex })),
				raw(strRefBlock(sheetCellRef(obj._dataIndex + dataLabels(obj).length + 1, 1), obj.name ?? '')),
				raw(serShapeProps(chartType, opts, seriesColor, seriesOverride?.lineSize, serIndex)),
				// `invertIfNegative` is bar-only in the schema (CT_BarSer); area/line/radar must omit it.
				isBarLike(chartType) ? raw(voidEl('c:invertIfNegative', { val: 0 })) : null,
				isLineLike(chartType) ? raw(serMarker(opts, paletteColor(chartColors, obj._dataIndex), seriesColor)) : null,
				// Per-point data points (`c:dPt`) MUST precede `c:dLbls` in CT_*Ser schema order.
				raw(makeSeriesDataPointsXml(chartType, obj, opts, barVaryColors)),
				// NOTE: [20190117] Adding data labels to a RADAR chart causes unrecoverable corruption,
				// and CT_RadarSer has no error bars either.
				chartType === ChartType.radar ? null : raw(serDataLabels(obj, opts, seriesColor)),
				chartType === ChartType.radar ? null : raw(makeChartErrorBarsXml(chartType, obj.errorBars, obj)),
				raw(serCategories(obj, opts)),
				raw(serValues(obj, valFmtCode)),
				chartType === ChartType.line ? raw(voidEl('c:smooth', { val: opts.lineSmooth ? 1 : 0 })) : null,
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
