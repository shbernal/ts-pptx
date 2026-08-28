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
import { el, raw, voidEl } from '../oxml/el.js'
import {
	catRefBlock,
	dimmedTextFill,
	dimmedTextLine,
	numCachePt,
	paletteColor,
	resolveChartPalette,
	strRefBlock,
} from './chart-parts.js'

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
const STOCK_DLBLS = el('c:dLbls', null, [
	raw(voidEl('c:showLegendKey', { val: 0 })),
	raw(voidEl('c:showVal', { val: 0 })),
	raw(voidEl('c:showCatName', { val: 0 })),
	raw(voidEl('c:showSerName', { val: 0 })),
	raw(voidEl('c:showPercent', { val: 0 })),
	raw(voidEl('c:showBubbleSize', { val: 0 })),
])

/** Emit the shared `<c:cat>` + `<c:val>` refs for a stock/volume series (single-level categories). */
function stockCatVal(obj: OptsChartDataInternal, opts: ChartOptsInternal, valFmtCode: string): string {
	const cats = firstLabelGroup(obj)
	const valColRow = obj._dataIndex + dataLabels(obj).length + 1
	const catRef = `Sheet1!$A$2:$A$${cats.length + 1}`
	// Numeric categories (dates) take a `numRef` carrying the source format, so PowerPoint renders
	// them as dates rather than serial numbers; text ones take a plain `strRef`.
	const cat = el(
		'c:cat',
		null,
		raw(
			opts.catLabelFormatCode
				? catRefBlock('num', catRef, cats, opts.catLabelFormatCode || 'General')
				: catRefBlock('str', catRef, cats)
		)
	)
	const numCache = el('c:numCache', null, [
		raw(el('c:formatCode', null, valFmtCode)),
		raw(voidEl('c:ptCount', { val: cats.length })),
		raw(
			dataValues(obj)
				.map((value, idx) => numCachePt(idx, value))
				.join('')
		),
	])
	const numRef = el('c:numRef', null, [
		raw(el('c:f', null, sheetRangeRef(valColRow, 2, valColRow, cats.length + 1))),
		raw(numCache),
	])
	return cat + el('c:val', null, raw(numRef))
}

/** Emit the `<c:tx>` series-name reference for a stock/volume series. */
function stockSeriesName(obj: OptsChartDataInternal): string {
	const nameCol = obj._dataIndex + dataLabels(obj).length + 1
	return strRefBlock(sheetCellRef(nameCol, 1), obj.name ?? '', 'compact')
}

/** Emit a single stock (line) series: invisible line, optional close-marker, cat/val refs. */
function makeStockLineSer(
	obj: OptsChartDataInternal,
	opts: ChartOptsInternal,
	valFmtCode: string,
	markCloseColor: string | null
): string {
	// Stock series draw no line themselves (the hi-low lines / up-down bars carry the visual).
	const spPr = el('c:spPr', null, [
		raw(el('a:ln', { w: 19050, cap: 'rnd' }, [raw(voidEl('a:noFill')), raw(voidEl('a:round'))])),
		raw(voidEl('a:effectLst')),
	])
	const marker = markCloseColor
		? el('c:marker', null, [
				raw(voidEl('c:symbol', { val: 'dot' })),
				raw(voidEl('c:size', { val: 5 })),
				raw(
					el('c:spPr', null, [
						raw(genXmlColorSelection(markCloseColor)),
						raw(el('a:ln', { w: 9525 }, raw(genXmlColorSelection(markCloseColor)))),
						raw(voidEl('a:effectLst')),
					])
				),
			])
		: el('c:marker', null, raw(voidEl('c:symbol', { val: 'none' })))
	return el('c:ser', null, [
		raw(voidEl('c:idx', { val: obj._dataIndex })),
		raw(voidEl('c:order', { val: obj._dataIndex })),
		raw(stockSeriesName(obj)),
		raw(spPr),
		raw(marker),
		raw(stockCatVal(obj, opts, valFmtCode)),
		raw(voidEl('c:smooth', { val: 0 })),
	])
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
		const volumeSer = el('c:ser', null, [
			raw(voidEl('c:idx', { val: volumeSeries._dataIndex })),
			raw(voidEl('c:order', { val: volumeSeries._dataIndex })),
			raw(stockSeriesName(volumeSeries)),
			raw(el('c:spPr', null, raw(genXmlColorSelection(chartColors[0] ?? '4472C4')))),
			raw(voidEl('c:invertIfNegative', { val: 0 })),
			raw(stockCatVal(volumeSeries, opts, valFmtCode)),
		])
		strXml += el('c:barChart', null, [
			raw(voidEl('c:barDir', { val: 'col' })),
			raw(voidEl('c:grouping', { val: 'clustered' })),
			raw(voidEl('c:varyColors', { val: 0 })),
			raw(volumeSer),
			raw(STOCK_DLBLS),
			raw(voidEl('c:gapWidth', { val: 150 })),
			raw(voidEl('c:axId', { val: AXIS_ID_CATEGORY_PRIMARY })),
			raw(voidEl('c:axId', { val: AXIS_ID_VALUE_PRIMARY })),
		])
	}

	// STOCK: the price series drawn with invisible lines + hi-low lines (+ up-down bars for OHLC).
	const sers = stockSeries
		.map((obj, idx) => {
			// HLC/VHLC (no up-down bars) mark the final "close" series with a dot so it reads on the chart.
			const isClose = !spec.upDownBars && idx === stockSeries.length - 1
			const markColor = isClose ? paletteColor(chartColors, obj._dataIndex, 'ED7D31') : null
			return makeStockLineSer(obj, opts, valFmtCode, markColor)
		})
		.join('')
	const hiLowLines = el(
		'c:hiLowLines',
		null,
		raw(el('c:spPr', null, [raw(dimmedTextLine(75000, 25000)), raw(voidEl('a:effectLst'))]))
	)
	// The up bar is a hollow body (background fill), the down bar a filled one; both are outlined in
	// the same grey, which is how a rising and a falling category tell themselves apart.
	const upDownBars = spec.upDownBars
		? el('c:upDownBars', null, [
				raw(voidEl('c:gapWidth', { val: 150 })),
				raw(
					el(
						'c:upBars',
						null,
						raw(
							el('c:spPr', null, [
								raw(el('a:solidFill', null, raw(voidEl('a:schemeClr', { val: 'bg1' })))),
								raw(dimmedTextLine(65000, 35000)),
								raw(voidEl('a:effectLst')),
							])
						)
					)
				),
				raw(
					el(
						'c:downBars',
						null,
						raw(
							el('c:spPr', null, [
								raw(dimmedTextFill(65000, 35000)),
								raw(dimmedTextLine(65000, 35000)),
								raw(voidEl('a:effectLst')),
							])
						)
					)
				),
			])
		: ''
	// Volume styles put the price series on the SECONDARY axis pair (the bar owns the primary pair).
	const stockCatId = spec.volume ? AXIS_ID_CATEGORY_SECONDARY : catAxisId
	const stockValId = spec.volume ? AXIS_ID_VALUE_SECONDARY : valAxisId
	strXml += el('c:stockChart', null, [
		raw(sers),
		raw(STOCK_DLBLS),
		raw(hiLowLines),
		raw(upDownBars),
		raw(voidEl('c:axId', { val: stockCatId })),
		raw(voidEl('c:axId', { val: stockValId })),
	])

	return strXml
}
