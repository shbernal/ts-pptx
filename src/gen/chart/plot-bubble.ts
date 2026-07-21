/**
 * PptxGenJS: Bubble Plot Assembly
 *
 * Emits the `<c:bubbleChart>` plot element for `bubble` and `bubble3D`. Like scatter it
 * takes X values from the first data row, and adds a third `<c:bubbleSize>` cache per
 * series. Reached through {@link ./chart-xml}'s `makeChartType` dispatch.
 */

import { ChartType } from '../../core-enums.js'
import { BARCHART_COLORS, DEF_FONT_COLOR, DEF_FONT_SIZE, DEF_SHAPE_SHADOW } from '../../core-enums-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { resolveBorderWidth } from '../drawingml/line.js'
import { valToPts } from '../../units-internal.js'
import { FIXED_PCT_PER_PERCENT, ptToHundredths } from '../../units.js'
import { dataSizes, dataValues, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, voidEl } from '../oxml/el.js'
import { createChartTextFonts, numCachePt } from './chart-parts.js'

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
	let colorIndex = -1 // Maintain the color index by region
	let idxColLtr = 1
	let strXml = ''
	/*
				`data` = [
					{ name:'X-Axis',     values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Values 1', values:[13, 20, 21, 25], sizes:[10, 5, 20, 15] },
					{ name:'Y-Values 2', values:[ 1,  2,  5,  9], sizes:[ 5, 3,  9,  3] }
				];
            */

	// 1: Start Chart
	strXml += '<c:bubbleChart>'
	strXml += '<c:varyColors val="0"/>'

	// 2: Series: (One for each Y-Axis)
	colorIndex = -1
	data
		.filter((_obj, idx) => idx > 0)
		.forEach((obj, idx) => {
			colorIndex++
			strXml += '<c:ser>'
			strXml += `  <c:idx val="${idx}"/>`
			strXml += `  <c:order val="${idx}"/>`

			// A: `<c:tx>`
			strXml += '  <c:tx>'
			strXml += '    <c:strRef>'
			strXml += `      <c:f>${sheetCellRef(idxColLtr + 1, 1)}</c:f>`
			strXml +=
				'      <c:strCache><c:ptCount val="1"/><c:pt idx="0">' +
				el('c:v', null, obj.name ?? '') +
				'</c:pt></c:strCache>'
			strXml += '    </c:strRef>'
			strXml += '  </c:tx>'

			// B: '<c:spPr>': Fill, Border, Line, LineStyle (dash, etc.), Shadow
			{
				strXml += '<c:spPr>'

				const chartColors = opts.chartColors?.length ? opts.chartColors : BARCHART_COLORS
				const tmpSerColor = chartColors[colorIndex % chartColors.length] ?? '000000'

				if (tmpSerColor === 'transparent') {
					strXml += '<a:noFill/>'
				} else if (opts.chartColorsOpacity) {
					strXml += `<a:solidFill>${createColorElement(tmpSerColor, '<a:alpha val="' + Math.round(opts.chartColorsOpacity * FIXED_PCT_PER_PERCENT).toString() + '"/>')}</a:solidFill>`
				} else {
					strXml += genXmlColorSelection(tmpSerColor)
				}

				if (opts.lineSize === 0) {
					strXml += '<a:ln><a:noFill/></a:ln>'
				} else if (opts.dataBorder) {
					strXml += `<a:ln w="${valToPts(resolveBorderWidth(opts.dataBorder, 0.75))}" cap="flat">${genXmlColorSelection({ color: opts.dataBorder.color ?? '363636', transparency: opts.dataBorder.transparency })}<a:prstDash val="solid"/><a:round/></a:ln>`
				} else {
					strXml += `<a:ln w="${valToPts(opts.lineSize ?? 2)}" cap="flat">${genXmlColorSelection(tmpSerColor)}`
					strXml +=
						voidEl('a:prstDash', { val: opts.lineDashValues?.[colorIndex] ?? opts.lineDash ?? 'solid' }) +
						voidEl('a:round') +
						'</a:ln>'
				}

				// Shadow
				strXml += createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)

				strXml += '</c:spPr>'
			}

			// C: '<c:dLbls>' "Data Labels"
			// Let it be defaulted for now

			// D: '<c:xVal>'/'<c:yVal>' "Values": Scatter Chart has 2: `xVal` and `yVal`
			{
				// X-Axis is always the same
				strXml += '<c:xVal>'
				strXml += '  <c:numRef>'
				strXml += `    <c:f>Sheet1!$A$2:$A$${dataValues(data[0]).length + 1}</c:f>`
				strXml += '    <c:numCache>'
				strXml += '      <c:formatCode>' + valFmtCode + '</c:formatCode>'
				strXml += `      <c:ptCount val="${dataValues(data[0]).length}"/>`
				dataValues(data[0]).forEach((value, idx) => {
					strXml += numCachePt(idx, value)
				})
				strXml += '    </c:numCache>'
				strXml += '  </c:numRef>'
				strXml += '</c:xVal>'

				// Y-Axis vals are this object's `values`
				strXml += '<c:yVal>'
				strXml += '  <c:numRef>'
				strXml += `<c:f>${sheetRangeRef(idxColLtr + 1, 2, idxColLtr + 1, dataValues(data[0]).length + 1)}</c:f>`
				idxColLtr++
				strXml += '    <c:numCache>'
				strXml += '      <c:formatCode>' + valFmtCode + '</c:formatCode>'
				// NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
				strXml += `      <c:ptCount val="${dataValues(data[0]).length}"/>`
				dataValues(data[0]).forEach((_value, idx) => {
					strXml += numCachePt(idx, dataValues(obj)[idx])
				})
				strXml += '    </c:numCache>'
				strXml += '  </c:numRef>'
				strXml += '</c:yVal>'
			}

			// E: '<c:bubbleSize>'
			strXml += '  <c:bubbleSize>'
			strXml += '    <c:numRef>'
			strXml += `<c:f>${sheetRangeRef(idxColLtr + 1, 2, idxColLtr + 1, dataSizes(obj).length + 1)}</c:f>`
			idxColLtr++
			strXml += '      <c:numCache>'
			strXml += '        <c:formatCode>General</c:formatCode>'
			strXml += `           <c:ptCount val="${dataSizes(obj).length}"/>`
			dataSizes(obj).forEach((value, idx) => {
				strXml += numCachePt(idx, value)
			})
			strXml += '      </c:numCache>'
			strXml += '    </c:numRef>'
			strXml += '  </c:bubbleSize>'
			strXml += '  <c:bubble3D val="' + (chartType === ChartType.bubble3d ? '1' : '0') + '"/>'

			// F: Close "SERIES"
			strXml += '</c:ser>'
		})

	// 3: Data Labels
	{
		strXml += '<c:dLbls>'
		strXml += voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })
		strXml += '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
		strXml += `<a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" u="none">`
		strXml += genXmlColorSelection(opts.dataLabelColor || DEF_FONT_COLOR)
		strXml += createChartTextFonts(opts.dataLabelFontFace || 'Arial')
		strXml += '</a:defRPr></a:pPr></a:p></c:txPr>'
		if (opts.dataLabelPosition) strXml += voidEl('c:dLblPos', { val: opts.dataLabelPosition })
		strXml += '<c:showLegendKey val="0"/>'
		strXml += `<c:showVal val="${opts.showValue ? '1' : '0'}"/>`
		strXml += `<c:showCatName val="0"/><c:showSerName val="${opts.showSerName ? '1' : '0'}"/><c:showPercent val="0"/><c:showBubbleSize val="${opts.showBubbleSize ? '1' : '0'}"/>`
		strXml += '<c:extLst>'
		strXml +=
			'  <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">'
		strXml += '    <c15:showLeaderLines val="' + (opts.showLeaderLines ? '1' : '0') + '"/>'
		strXml += '  </c:ext>'
		strXml += '</c:extLst>'
		strXml += '</c:dLbls>'
	}

	// 4: Bubble options
	// `<c:bubbleScale>` / `<c:showNegBubbles>` are intentionally omitted so PowerPoint
	// applies its own defaults; no library option exposes them yet.

	// 5: AxisId (NOTE: order matters! (category comes first))
	strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/>`

	// 6: Close Chart tag
	strXml += '</c:bubbleChart>'
	return strXml
}
