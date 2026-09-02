/**
 * ts-pptx: Bubble Plot Assembly
 *
 * Emits the `<c:bubbleChart>` plot element for `bubble` and `bubble3D`. Like scatter it
 * takes X values from the first data row, and adds a third `<c:bubbleSize>` cache per
 * series. Reached through {@link ./chart-xml}'s `makeChartType` dispatch.
 */

import { ChartType } from '../../enums.js'
import type { ChartOptsInternal } from '../../types/internal.js'
import { createLineCap } from '../drawingml/line.js'
import { ptsToEmuLenient } from '../../units-internal.js'
import { categoryRange, dataSizes, dataValues, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	chartColorLineFill,
	createDataBorderLine,
	dataLabelDefRPr,
	dLblShowFlags,
	labelTextProps,
	numRefBlock,
	paletteColor,
	resolveChartPalette,
	seriesShapeProps,
	strRefBlock,
	type PlotBuilder,
} from './chart-parts.js'

/**
 * Fill + outline + shadow for one bubble series, from the palette colour and the line options.
 *
 * Both the stroke fill and the cap read the same way as in `plot-scatter.ts` and
 * `plot-cat-axis.ts`, which this used to differ from on both counts with nothing saying why.
 * The colour went through `genXmlColorSelection` directly, so a `'transparent'` palette entry
 * reached colour validation, warned, and painted the bubble outline black — the exact hole
 * `chartColorLineFill` exists to close. The cap was hardcoded `flat`, so `lineCap` was accepted
 * and silently dropped for bubble charts alone; it is visible here, because a bubble outline
 * carries `lineDash` and a cap shapes the end of every dash.
 */
function bubbleSerShapeProps(opts: ChartOptsInternal, serColor: string, serIndex: number): string {
	const line =
		opts.lineSize === 0
			? el('a:ln', null, raw(voidEl('a:noFill')))
			: opts.dataBorder
				? createDataBorderLine(opts.dataBorder, createLineCap(opts.lineCap))
				: el('a:ln', { w: ptsToEmuLenient(opts.lineSize ?? 2), cap: createLineCap(opts.lineCap) }, [
						raw(chartColorLineFill(serColor)),
						raw(voidEl('a:prstDash', { val: opts.lineDashValues?.[serIndex] ?? opts.lineDash ?? 'solid' })),
						raw(voidEl('a:round')),
					])
	return seriesShapeProps(opts, serColor, line)
}

/** The shared `<c:dLbls>` block: number format, label text style, and which parts are shown. */
function bubbleDataLabels(opts: ChartOptsInternal): string {
	const defRPr = dataLabelDefRPr(opts)
	const txPr = labelTextProps(defRPr)
	// The 2012 chart extension carrying the leader-line toggle; it has no c: equivalent.
	const extLst = el(
		'c:extLst',
		null,
		raw(
			el(
				'c:ext',
				{
					uri: '{CE6537A1-D6FC-4f65-9D91-7224C49458BB}',
					'xmlns:c15': 'http://schemas.microsoft.com/office/drawing/2012/chart',
				},
				raw(voidEl('c15:showLeaderLines', { val: opts.showLeaderLines ? 1 : 0 }))
			)
		)
	)
	return el('c:dLbls', null, [
		raw(voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })),
		raw(txPr),
		opts.dataLabelPosition ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition })) : null,
		...dLblShowFlags({
			val: opts.showValue ? 1 : 0,
			serName: opts.showSerName ? 1 : 0,
			bubbleSize: opts.showBubbleSize ? 1 : 0,
		}),
		raw(extLst),
	])
}

/**
 * Plot a bubble / bubble3d chart into `<c:bubbleChart>` (X/Y plus per-point size).
 */
export const makeBubblePlot: PlotBuilder = (chartType, data, opts, valAxisId, catAxisId, valFmtCode) => {
	/*
				`data` = [
					{ name:'X-Axis',     values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Values 1', values:[13, 20, 21, 25], sizes:[10, 5, 20, 15] },
					{ name:'Y-Values 2', values:[ 1,  2,  5,  9], sizes:[ 5, 3,  9,  3] }
				];
            */
	const chartColors = resolveChartPalette(opts)
	// X values come from the first row; each later row is one series, and consumes two sheet
	// columns — its Y values and its sizes.
	const xValues = dataValues(data[0])
	let idxColLtr = 1

	// One series per Y-Axis row.
	const sers = data
		.slice(1)
		.map((obj, idx) => {
			const name = strRefBlock(sheetCellRef(idxColLtr + 1, 1), obj.name ?? '')
			const spPr = bubbleSerShapeProps(opts, paletteColor(chartColors, idx), idx)
			// The Y series is cached against the X series' length, so a caller who supplied fewer Y
			// values than X leaves gaps rather than a short cache.
			const yValues = dataValues(obj)
			const xVal = numRefBlock('c:xVal', categoryRange(xValues.length), valFmtCode, xValues)
			const yVal = numRefBlock(
				'c:yVal',
				sheetRangeRef(idxColLtr + 1, 2, idxColLtr + 1, xValues.length + 1),
				valFmtCode,
				xValues.map((_value, i) => yValues[i])
			)
			idxColLtr++
			// The sizes carry a constant `General` format code: no option spells a size number format.
			const sizes = dataSizes(obj)
			const sizeVal = numRefBlock(
				'c:bubbleSize',
				sheetRangeRef(idxColLtr + 1, 2, idxColLtr + 1, sizes.length + 1),
				'General',
				sizes
			)
			idxColLtr++
			return el('c:ser', null, [
				raw(voidEl('c:idx', { val: idx })),
				raw(voidEl('c:order', { val: idx })),
				raw(name),
				raw(spPr),
				// No `<c:dLbls>` per series — the chart-level block below carries the labels.
				raw(xVal),
				raw(yVal),
				raw(sizeVal),
				raw(voidEl('c:bubble3D', { val: chartType === ChartType.bubble3d ? 1 : 0 })),
			])
		})
		.join('')

	// `<c:bubbleScale>` / `<c:showNegBubbles>` are intentionally omitted so PowerPoint applies its
	// own defaults; no library option exposes them yet.
	return el('c:bubbleChart', null, [
		raw(voidEl('c:varyColors', { val: 0 })),
		raw(sers),
		raw(bubbleDataLabels(opts)),
		// Axis id order matters: category comes first.
		raw(voidEl('c:axId', { val: catAxisId })),
		raw(voidEl('c:axId', { val: valAxisId })),
	])
}
