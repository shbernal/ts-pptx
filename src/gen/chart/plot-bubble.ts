/**
 * ts-pptx: Bubble Plot Assembly
 *
 * Emits the `<c:bubbleChart>` plot element for `bubble` and `bubble3D`. Like scatter it
 * takes X values from the first data row, and adds a third `<c:bubbleSize>` cache per
 * series. Reached through {@link ./chart-xml}'s `makeChartType` dispatch.
 */

import { ChartType } from '../../enums.js'
import { DEF_FONT_COLOR, DEF_FONT_SIZE, DEF_SHAPE_SHADOW } from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { resolveBorderWidth } from '../drawingml/line.js'
import { lineWidthToEmu, percentToFixedPercent, ptsToEmuLenient } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { dataSizes, dataValues, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	createChartTextFonts,
	numCachePt,
	numRefBlock,
	paletteColor,
	resolveChartPalette,
	strRefBlock,
} from './chart-parts.js'

/** Fill + outline + shadow for one bubble series, from the palette colour and the line options. */
function bubbleSerShapeProps(opts: ChartOptsInternal, serColor: string, serIndex: number): string {
	const fill =
		serColor === 'transparent'
			? voidEl('a:noFill')
			: opts.chartColorsOpacity
				? el(
						'a:solidFill',
						null,
						raw(
							createColorElement(
								serColor,
								voidEl('a:alpha', {
									val: percentToFixedPercent(
										opts.chartColorsOpacity,
										'chart/option-out-of-range',
										'chartColorsOpacity'
									),
								})
							)
						)
					)
				: genXmlColorSelection(serColor)
	const line =
		opts.lineSize === 0
			? el('a:ln', null, raw(voidEl('a:noFill')))
			: opts.dataBorder
				? el('a:ln', { w: lineWidthToEmu(resolveBorderWidth(opts.dataBorder, 0.75)), cap: 'flat' }, [
						raw(
							genXmlColorSelection({
								color: opts.dataBorder.color ?? '363636',
								transparency: opts.dataBorder.transparency,
							})
						),
						raw(voidEl('a:prstDash', { val: 'solid' })),
						raw(voidEl('a:round')),
					])
				: el('a:ln', { w: ptsToEmuLenient(opts.lineSize ?? 2), cap: 'flat' }, [
						raw(genXmlColorSelection(serColor)),
						raw(voidEl('a:prstDash', { val: opts.lineDashValues?.[serIndex] ?? opts.lineDash ?? 'solid' })),
						raw(voidEl('a:round')),
					])
	return el('c:spPr', null, [raw(fill), raw(line), raw(createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW))])
}

/**
 * The `<c:bubbleSize>` cache: the per-point sizes that scale each bubble. Its own block rather
 * than {@link numRefBlock} because the wrapping tag differs and so does the indentation — the
 * `<c:ptCount>` here sits three spaces further in than anything around it.
 */
function bubbleSizeBlock(obj: OptsChartDataInternal, ref: string): string {
	const sizes = dataSizes(obj)
	const numCache = el(
		'c:numCache',
		null,
		[
			raw(el('c:formatCode', null, 'General', { openPrefix: '        ' })),
			raw(voidEl('c:ptCount', { val: sizes.length }, { openPrefix: '           ' })),
			raw(sizes.map((value, idx) => numCachePt(idx, value)).join('')),
		],
		{ openPrefix: '      ', closePrefix: '      ' }
	)
	const numRef = el('c:numRef', null, [raw(el('c:f', null, ref)), raw(numCache)], {
		openPrefix: '    ',
		closePrefix: '    ',
	})
	return el('c:bubbleSize', null, raw(numRef), { openPrefix: '  ', closePrefix: '  ' })
}

/** The shared `<c:dLbls>` block: number format, label text style, and which parts are shown. */
function bubbleDataLabels(opts: ChartOptsInternal): string {
	const defRPr = el(
		'a:defRPr',
		{
			b: opts.dataLabelFontBold ? 1 : 0,
			i: opts.dataLabelFontItalic ? 1 : 0,
			strike: 'noStrike',
			sz: ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE),
			u: 'none',
		},
		[
			raw(genXmlColorSelection(opts.dataLabelColor || DEF_FONT_COLOR)),
			raw(createChartTextFonts(opts.dataLabelFontFace || 'Arial')),
		]
	)
	const txPr = el('c:txPr', null, [
		raw(voidEl('a:bodyPr')),
		raw(voidEl('a:lstStyle')),
		raw(el('a:p', null, raw(el('a:pPr', null, raw(defRPr))))),
	])
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
				raw(voidEl('c15:showLeaderLines', { val: opts.showLeaderLines ? 1 : 0 }, { openPrefix: '    ' })),
				{ openPrefix: '  ', closePrefix: '  ' }
			)
		)
	)
	return el('c:dLbls', null, [
		raw(voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })),
		raw(txPr),
		opts.dataLabelPosition ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition })) : null,
		raw(voidEl('c:showLegendKey', { val: 0 })),
		raw(voidEl('c:showVal', { val: opts.showValue ? 1 : 0 })),
		raw(voidEl('c:showCatName', { val: 0 })),
		raw(voidEl('c:showSerName', { val: opts.showSerName ? 1 : 0 })),
		raw(voidEl('c:showPercent', { val: 0 })),
		raw(voidEl('c:showBubbleSize', { val: opts.showBubbleSize ? 1 : 0 })),
		raw(extLst),
	])
}

/**
 * Plot a bubble / bubble3d chart into `<c:bubbleChart>` (X/Y plus per-point size).
 */
export function makeBubblePlot(
	chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valAxisId: string,
	catAxisId: string,
	valFmtCode: string
): string {
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
			const xVal = numRefBlock('c:xVal', `Sheet1!$A$2:$A$${xValues.length + 1}`, valFmtCode, xValues)
			const yVal = numRefBlock(
				'c:yVal',
				sheetRangeRef(idxColLtr + 1, 2, idxColLtr + 1, xValues.length + 1),
				valFmtCode,
				xValues.map((_value, i) => yValues[i]),
				'' // bubble's yVal has never indented its `<c:f>`; see numRefBlock
			)
			idxColLtr++
			const sizeRef = sheetRangeRef(idxColLtr + 1, 2, idxColLtr + 1, dataSizes(obj).length + 1)
			idxColLtr++
			return el('c:ser', null, [
				raw(voidEl('c:idx', { val: idx }, { openPrefix: '  ' })),
				raw(voidEl('c:order', { val: idx }, { openPrefix: '  ' })),
				raw(name),
				raw(spPr),
				// No `<c:dLbls>` per series — the chart-level block below carries the labels.
				raw(xVal),
				raw(yVal),
				raw(bubbleSizeBlock(obj, sizeRef)),
				raw(voidEl('c:bubble3D', { val: chartType === ChartType.bubble3d ? 1 : 0 }, { openPrefix: '  ' })),
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
