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

import { namedColorOr } from '../drawingml/color.js'
import {
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
} from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import type { WorksheetLayout } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	catValRefs,
	dLblShowFlags,
	dimmedTextFill,
	dimmedTextLine,
	paletteColor,
	resolveChartPalette,
	seriesNameRef,
	type PlotBuilder,
} from './chart-parts.js'
import { seriesIdx, STOCK_STYLE_SPEC, type StockStyle } from './chart-kind.js'

/** True when the given (already-normalized) stock style leads with a Volume bar series. */
export const isVolumeStockStyle = (style: StockStyle | undefined): boolean => !!style && STOCK_STYLE_SPEC[style].volume

/** Minimal `<c:dLbls>` block (all labels off) shared by the stock and volume-bar subcharts. */
const STOCK_DLBLS = el('c:dLbls', null, [...dLblShowFlags({})])

/** Emit a single stock (line) series: invisible line, optional close-marker, cat/val refs. */
function makeStockLineSer(
	idx: number,
	obj: OptsChartDataInternal,
	opts: ChartOptsInternal,
	valFmtCode: string,
	markCloseColor: string | null,
	sheet: WorksheetLayout
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
		raw(voidEl('c:idx', { val: idx })),
		raw(voidEl('c:order', { val: idx })),
		raw(seriesNameRef(obj, sheet)),
		raw(spPr),
		raw(marker),
		raw(catValRefs(obj, opts, valFmtCode, sheet, { multiLevel: false })),
		raw(voidEl('c:smooth', { val: 0 })),
	])
}

/**
 * Plot a stock chart. For `vhlc`/`vohlc` the first data series is a Volume column drawn as a
 * `<c:barChart>` on the primary axis pair, with the remaining high/low/close(/open) series in a
 * `<c:stockChart>` on the secondary axis pair; the non-volume styles draw only the stock chart on
 * the primary pair. `valAxisId`/`catAxisId` are the primary ids passed by the dispatch.
 */
export const makeStockPlot: PlotBuilder = (chartType, data, opts, valAxisId, catAxisId, valFmtCode, sheet) => {
	const spec = STOCK_STYLE_SPEC[(opts.stockStyle as StockStyle) || 'hlc']
	const chartColors = resolveChartPalette(opts)
	const volumeSeries = spec.volume ? data[0] : null
	const stockSeries = spec.volume ? data.slice(1) : data
	let strXml = ''

	// VOLUME: a bar series on the PRIMARY axis pair (drawn behind the price series).
	if (volumeSeries) {
		const volumeIdx = seriesIdx(volumeSeries, chartType)
		const volumeSer = el('c:ser', null, [
			raw(voidEl('c:idx', { val: volumeIdx })),
			raw(voidEl('c:order', { val: volumeIdx })),
			raw(seriesNameRef(volumeSeries, sheet)),
			raw(
				el(
					'c:spPr',
					null,
					raw(
						genXmlColorSelection(
							namedColorOr(opts.seriesOptions?.[volumeIdx]?.color, chartColors[0] ?? '4472C4', 'seriesOptions color')
						)
					)
				)
			),
			raw(voidEl('c:invertIfNegative', { val: 0 })),
			raw(catValRefs(volumeSeries, opts, valFmtCode, sheet, { multiLevel: false })),
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
		.map((obj, pos) => {
			const idx = seriesIdx(obj, chartType)
			// HLC/VHLC (no up-down bars) mark the final "close" series with a dot so it reads on the chart.
			const isClose = !spec.upDownBars && pos === stockSeries.length - 1
			// The price series draw no line of their own -- the hi-low lines and up-down bars carry the
			// visual -- so a `seriesOptions.color` has exactly two referents on a stock chart: the
			// volume bar above, and the dot that marks the close series where there are no up-down bars.
			const markColor = isClose
				? namedColorOr(
						opts.seriesOptions?.[idx]?.color,
						paletteColor(chartColors, idx, 'ED7D31'),
						'seriesOptions color'
					)
				: null
			return makeStockLineSer(idx, obj, opts, valFmtCode, markColor, sheet)
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
