/**
 * ts-pptx: Pie & Doughnut Plot Assembly
 *
 * Emits the `<c:pieChart>` / `<c:doughnutChart>` plot elements. These are the only
 * families with no axes at all -- a single series, one `<c:dPt>` per slice carrying its
 * own fill, and optional leader lines -- so the builder takes no axis ids. Reached
 * through {@link ./chart-xml}'s `makeChartType` dispatch.
 */

import { ChartType } from '../../enums.js'
import { DEF_FONT_COLOR, DEF_FONT_SIZE, DEF_SHAPE_SHADOW } from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { encodeXmlEntities } from '../utils.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { resolveBorderWidth } from '../drawingml/line.js'
import { ptsToEmuLenient } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { dataValues, firstLabelGroup } from './data-refs.js'
import { el, voidEl } from '../oxml/el.js'
import {
	createChartBorderLine,
	createChartTextFonts,
	createLeaderLinesElement,
	paletteColor,
	resolveChartPalette,
} from './chart-parts.js'

/**
 * Plot a single-series pie / doughnut chart into `<c:pieChart>` / `<c:doughnutChart>`.
 */
export function makePiePlot(
	chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valFmtCode: string
): string {
	let optsChartData: OptsChartDataInternal
	let strXml = ''
	// Use the same let name so code blocks from barChart are interchangeable
	{
		const first = data[0]
		if (!first) return strXml
		optsChartData = first
	}

	/* EX:
				data: [
				 {
				   name: 'Project Status',
				   labels: ['Red', 'Amber', 'Green', 'Unknown'],
				   values: [10, 20, 38, 2]
				 }
				]
            */

	// 1: Start Chart
	strXml += '<c:' + chartType + 'Chart>'
	strXml += '  <c:varyColors val="1"/>'
	strXml += '<c:ser>'
	strXml += '  <c:idx val="0"/>'
	strXml += '  <c:order val="0"/>'
	strXml += '  <c:tx>'
	strXml += '    <c:strRef>'
	strXml += '      <c:f>Sheet1!$B$1</c:f>'
	strXml += '      <c:strCache>'
	strXml += '        <c:ptCount val="1"/>'
	strXml += '        <c:pt idx="0">' + el('c:v', null, optsChartData.name ?? '') + '</c:pt>'
	strXml += '      </c:strCache>'
	strXml += '    </c:strRef>'
	strXml += '  </c:tx>'
	strXml += '  <c:spPr>'
	strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'
	strXml +=
		'    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'
	if (opts.dataNoEffects) {
		strXml += '<a:effectLst/>'
	} else {
		strXml += createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)
	}
	strXml += '  </c:spPr>'

	// 2: "Data Point" block for every data row
	firstLabelGroup(optsChartData).forEach((_label, idx) => {
		const chartColors = resolveChartPalette(opts)
		const ptStyle = optsChartData.pointStyles?.[idx]
		strXml += '<c:dPt>'
		strXml += ` <c:idx val="${idx}"/>`
		strXml += ' <c:bubble3D val="0"/>'
		strXml += ' <c:spPr>'
		strXml += `<a:solidFill>${createColorElement(ptStyle?.fill || paletteColor(chartColors, idx))}</a:solidFill>`
		// Per-point border override takes precedence over chart-level `dataBorder`
		if (ptStyle?.border) {
			strXml += createChartBorderLine(ptStyle.border)
		} else if (opts.dataBorder) {
			strXml += `<a:ln w="${ptsToEmuLenient(resolveBorderWidth(opts.dataBorder, 0.75))}" cap="flat">${genXmlColorSelection(
				{
					color: opts.dataBorder.color ?? '363636',
					transparency: opts.dataBorder.transparency,
				}
			)}<a:prstDash val="solid"/><a:round/></a:ln>`
		}
		strXml += createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)
		strXml += '  </c:spPr>'
		strXml += '</c:dPt>'
	})

	// 3: "Data Label" block for every data Label
	strXml += '<c:dLbls>'
	firstLabelGroup(optsChartData).forEach((_label, idx) => {
		const customLbl = optsChartData.customLabels?.[idx]
		strXml += '<c:dLbl>'
		strXml += ` <c:idx val="${idx}"/>`
		// c:tx must precede c:numFmt per CT_DLbl / Group_DLbl / EG_DLblShared schema order
		if (customLbl) {
			strXml +=
				'<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>' +
				`<a:rPr lang="${encodeXmlEntities(opts.lang || 'en-US')}" dirty="0"/>` +
				el('a:t', null, customLbl) +
				'</a:r></a:p></c:rich></c:tx>'
		}
		strXml += '  ' + voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })
		strXml += '  <c:spPr/><c:txPr>'
		strXml += '   <a:bodyPr/><a:lstStyle/>'
		strXml += '   <a:p><a:pPr>'
		strXml += `   <a:defRPr sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" b="${opts.dataLabelFontBold ? 1 : 0}" i="${
			opts.dataLabelFontItalic ? 1 : 0
		}" u="none" strike="noStrike">`
		strXml += genXmlColorSelection(opts.dataLabelColor || DEF_FONT_COLOR)
		strXml += '    ' + createChartTextFonts(opts.dataLabelFontFace || 'Arial')
		strXml += '   </a:defRPr>'
		strXml += '      </a:pPr></a:p>'
		strXml += '    </c:txPr>'
		if (chartType === ChartType.pie && opts.dataLabelPosition)
			strXml += voidEl('c:dLblPos', { val: opts.dataLabelPosition })
		strXml += '    <c:showLegendKey val="0"/>'
		strXml += '    <c:showVal val="' + (customLbl ? '0' : opts.showValue ? '1' : '0') + '"/>'
		strXml += '    <c:showCatName val="' + (opts.showLabel ? '1' : '0') + '"/>'
		strXml += '    <c:showSerName val="' + (opts.showSerName ? '1' : '0') + '"/>'
		strXml += '    <c:showPercent val="' + (opts.showPercent ? '1' : '0') + '"/>'
		strXml += '    <c:showBubbleSize val="0"/>'
		strXml += '  </c:dLbl>'
	})
	strXml += ' ' + voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })
	strXml += '    <c:txPr>'
	strXml += '      <a:bodyPr/>'
	strXml += '      <a:lstStyle/>'
	strXml += '      <a:p>'
	strXml += '        <a:pPr>'
	strXml += `          <a:defRPr sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" b="${opts.dataLabelFontBold ? '1' : '0'}" i="${opts.dataLabelFontItalic ? '1' : '0'}" u="none" strike="noStrike">`
	strXml += genXmlColorSelection(opts.dataLabelColor || DEF_FONT_COLOR)
	strXml += '            ' + createChartTextFonts(opts.dataLabelFontFace || 'Arial')
	strXml += '          </a:defRPr>'
	strXml += '        </a:pPr>'
	strXml += '      </a:p>'
	strXml += '    </c:txPr>'
	strXml += chartType === ChartType.pie ? voidEl('c:dLblPos', { val: opts.dataLabelPosition || 'ctr' }) : ''
	strXml += '    <c:showLegendKey val="0"/>'
	strXml += '    <c:showVal val="0"/>'
	strXml += '    <c:showCatName val="1"/>'
	strXml += '    <c:showSerName val="0"/>'
	strXml += '    <c:showPercent val="1"/>'
	strXml += '    <c:showBubbleSize val="0"/>'
	strXml += ` <c:showLeaderLines val="${opts.showLeaderLines ? '1' : '0'}"/>`
	strXml += createLeaderLinesElement(opts)
	strXml += '</c:dLbls>'

	// 2: "Categories"
	strXml += '<c:cat>'
	strXml += '  <c:strRef>'
	strXml += `    <c:f>Sheet1!$A$2:$A$${firstLabelGroup(optsChartData).length + 1}</c:f>`
	strXml += '    <c:strCache>'
	strXml += `         <c:ptCount val="${firstLabelGroup(optsChartData).length}"/>`
	firstLabelGroup(optsChartData).forEach((label, idx) => {
		strXml += `<c:pt idx="${idx}">${el('c:v', null, label)}</c:pt>`
	})
	strXml += '    </c:strCache>'
	strXml += '  </c:strRef>'
	strXml += '</c:cat>'

	// 3: Create vals
	strXml += '  <c:val>'
	strXml += '    <c:numRef>'
	strXml += `      <c:f>Sheet1!$B$2:$B$${firstLabelGroup(optsChartData).length + 1}</c:f>`
	strXml += '      <c:numCache>'
	strXml += '        <c:formatCode>' + valFmtCode + '</c:formatCode>'
	strXml += `           <c:ptCount val="${firstLabelGroup(optsChartData).length}"/>`
	dataValues(optsChartData).forEach((value, idx) => {
		strXml += `<c:pt idx="${idx}"><c:v>${value || value === 0 ? value : ''}</c:v></c:pt>`
	})
	strXml += '      </c:numCache>'
	strXml += '    </c:numRef>'
	strXml += '  </c:val>'

	// 4: Close "SERIES"
	strXml += '  </c:ser>'
	strXml += `  <c:firstSliceAng val="${opts.firstSliceAng ? Math.round(opts.firstSliceAng) : 0}"/>`
	if (chartType === ChartType.doughnut)
		strXml += `<c:holeSize val="${typeof opts.holeSize === 'number' ? opts.holeSize : '50'}"/>`
	strXml += '</c:' + chartType + 'Chart>'
	return strXml
}

/**
 * Create Category axis
 * @param {ChartOptsInternal} opts - chart options
 * @param {string} axisId - value
 * @param {string} valAxisId - value
 * @return {string} XML
 */
