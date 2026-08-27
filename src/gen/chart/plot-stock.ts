/**
 * ts-pptx: Stock (High-Low-Close) Plot Assembly
 *
 * Emits the classic `<c:stockChart>` element (plus, for the volume styles, a leading
 * `<c:barChart>` volume series on a secondary axis pair). A stock chart is a fixed-order
 * multi-series line chart drawn with invisible lines: `<c:hiLowLines>` connect the
 * high/low of each category, and the open-high-low-close styles add `<c:upDownBars>`.
 * Reached through {@link ./chart-xml}'s `makeChartType` dispatch; its axes are built by
 * the stock branch of `makeChartAxesXml` (which reuses {@link ./chart-axes}).
 */

import { ChartType } from '../../enums.js'
import {
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
} from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { dataLabels, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, raw } from '../oxml/el.js'
import { catRefBlock, numCachePt, paletteColor, resolveChartPalette } from './chart-parts.js'

type StockStyle = 'hlc' | 'ohlc' | 'vhlc' | 'vohlc'

/**
 * Per-style stock chart geometry: how many value series the style expects, whether the first
 * series is a Volume column drawn as a bar (on its own axis pair), and whether the open-close
 * `<c:upDownBars>` are drawn. HLC/VHLC are three-value (no open) and instead mark the close
 * with a dot; OHLC/VOHLC are four-value and use up/down bars for the open-close body.
 */
const STOCK_STYLE_SPEC: Record<StockStyle, { seriesCount: number; volume: boolean; upDownBars: boolean }> = {
	hlc: { seriesCount: 3, volume: false, upDownBars: false },
	ohlc: { seriesCount: 4, volume: false, upDownBars: true },
	vhlc: { seriesCount: 4, volume: true, upDownBars: false },
	vohlc: { seriesCount: 5, volume: true, upDownBars: true },
}

/** True when the given (already-normalized) stock style leads with a Volume bar series. */
export const isVolumeStockStyle = (style: StockStyle | undefined): boolean => !!style && STOCK_STYLE_SPEC[style].volume

/** Minimal `<c:dLbls>` block (all labels off) shared by the stock and volume-bar subcharts. */
const STOCK_DLBLS =
	'<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'

/** Emit the shared `<c:cat>` + `<c:val>` refs for a stock/volume series (single-level categories). */
function stockCatVal(obj: OptsChartDataInternal, opts: ChartOptsInternal, valFmtCode: string): string {
	const cats = firstLabelGroup(obj)
	const valColRow = obj._dataIndex + dataLabels(obj).length + 1
	const catRef = `Sheet1!$A$2:$A$${cats.length + 1}`
	// Numeric categories (dates) take a `numRef` carrying the source format, so PowerPoint renders
	// them as dates rather than serial numbers; text ones take a plain `strRef`.
	let strXml = el(
		'c:cat',
		null,
		raw(
			opts.catLabelFormatCode
				? catRefBlock('num', catRef, cats, opts.catLabelFormatCode || 'General')
				: catRefBlock('str', catRef, cats)
		)
	)

	strXml += '<c:val><c:numRef>'
	strXml += `<c:f>${sheetRangeRef(valColRow, 2, valColRow, cats.length + 1)}</c:f>`
	strXml += '<c:numCache>'
	strXml += `<c:formatCode>${valFmtCode}</c:formatCode>`
	strXml += `<c:ptCount val="${cats.length}"/>`
	dataValues(obj).forEach((value, idx) => (strXml += numCachePt(idx, value)))
	strXml += '</c:numCache></c:numRef></c:val>'
	return strXml
}

/** Emit the `<c:tx>` series-name reference for a stock/volume series. */
function stockSeriesName(obj: OptsChartDataInternal): string {
	const nameCol = obj._dataIndex + dataLabels(obj).length + 1
	return (
		'<c:tx><c:strRef>' +
		`<c:f>${sheetCellRef(nameCol, 1)}</c:f>` +
		'<c:strCache><c:ptCount val="1"/><c:pt idx="0">' +
		el('c:v', null, obj.name ?? '') +
		'</c:pt></c:strCache></c:strRef></c:tx>'
	)
}

/** Emit a single stock (line) series: invisible line, optional close-marker, cat/val refs. */
function makeStockLineSer(
	obj: OptsChartDataInternal,
	opts: ChartOptsInternal,
	valFmtCode: string,
	markCloseColor: string | null
): string {
	let strXml = '<c:ser>'
	strXml += `<c:idx val="${obj._dataIndex}"/><c:order val="${obj._dataIndex}"/>`
	strXml += stockSeriesName(obj)
	// Stock series draw no line themselves (the hi-low lines / up-down bars carry the visual).
	strXml += '<c:spPr><a:ln w="19050" cap="rnd"><a:noFill/><a:round/></a:ln><a:effectLst/></c:spPr>'
	if (markCloseColor) {
		strXml +=
			'<c:marker><c:symbol val="dot"/><c:size val="5"/><c:spPr>' +
			genXmlColorSelection(markCloseColor) +
			`<a:ln w="9525">${genXmlColorSelection(markCloseColor)}</a:ln><a:effectLst/></c:spPr></c:marker>`
	} else {
		strXml += '<c:marker><c:symbol val="none"/></c:marker>'
	}
	strXml += stockCatVal(obj, opts, valFmtCode)
	strXml += '<c:smooth val="0"/>'
	strXml += '</c:ser>'
	return strXml
}

/**
 * Plot a stock chart. For `vhlc`/`vohlc` the first data series is a Volume column drawn as a
 * `<c:barChart>` on the primary axis pair, with the remaining high/low/close(/open) series in a
 * `<c:stockChart>` on the secondary axis pair; the non-volume styles draw only the stock chart on
 * the primary pair. `valAxisId`/`catAxisId` are the primary ids passed by the dispatch.
 */
export function makeStockPlot(
	_chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valAxisId: string,
	catAxisId: string,
	valFmtCode: string
): string {
	const spec = STOCK_STYLE_SPEC[(opts.stockStyle as StockStyle) || 'hlc']
	const chartColors = resolveChartPalette(opts)
	const volumeSeries = spec.volume ? data[0] : null
	const stockSeries = spec.volume ? data.slice(1) : data
	let strXml = ''

	// VOLUME: a bar series on the PRIMARY axis pair (drawn behind the price series).
	if (volumeSeries) {
		strXml += '<c:barChart>'
		strXml += '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>'
		strXml += '<c:ser>'
		strXml += `<c:idx val="${volumeSeries._dataIndex}"/><c:order val="${volumeSeries._dataIndex}"/>`
		strXml += stockSeriesName(volumeSeries)
		strXml += `<c:spPr>${genXmlColorSelection(chartColors[0] ?? '4472C4')}</c:spPr>`
		strXml += '<c:invertIfNegative val="0"/>'
		strXml += stockCatVal(volumeSeries, opts, valFmtCode)
		strXml += '</c:ser>'
		strXml += STOCK_DLBLS
		strXml += '<c:gapWidth val="150"/>'
		strXml += `<c:axId val="${AXIS_ID_CATEGORY_PRIMARY}"/><c:axId val="${AXIS_ID_VALUE_PRIMARY}"/>`
		strXml += '</c:barChart>'
	}

	// STOCK: the price series drawn with invisible lines + hi-low lines (+ up-down bars for OHLC).
	strXml += '<c:stockChart>'
	stockSeries.forEach((obj, idx) => {
		// HLC/VHLC (no up-down bars) mark the final "close" series with a dot so it reads on the chart.
		const isClose = !spec.upDownBars && idx === stockSeries.length - 1
		const markColor = isClose ? paletteColor(chartColors, obj._dataIndex, 'ED7D31') : null
		strXml += makeStockLineSer(obj, opts, valFmtCode, markColor)
	})
	strXml += STOCK_DLBLS
	strXml +=
		'<c:hiLowLines><c:spPr><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:hiLowLines>'
	if (spec.upDownBars) {
		strXml +=
			'<c:upDownBars><c:gapWidth val="150"/>' +
			'<c:upBars><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:upBars>' +
			'<c:downBars><c:spPr><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr></c:downBars>' +
			'</c:upDownBars>'
	}
	// Volume styles put the price series on the SECONDARY axis pair (the bar owns the primary pair).
	const stockCatId = spec.volume ? AXIS_ID_CATEGORY_SECONDARY : catAxisId
	const stockValId = spec.volume ? AXIS_ID_VALUE_SECONDARY : valAxisId
	strXml += `<c:axId val="${stockCatId}"/><c:axId val="${stockValId}"/>`
	strXml += '</c:stockChart>'

	return strXml
}
