/**
 * ts-pptx: Scatter Plot Assembly
 *
 * Emits the `<c:scatterChart>` plot element. Scatter is the one family whose first data
 * row supplies X *values* rather than categories, so each `<c:ser>` carries an
 * `<c:xVal>`/`<c:yVal>` pair instead of `<c:cat>`/`<c:val>` -- which is why it does not
 * share the category-axis builder. Reached through {@link ./chart-xml}'s `makeChartType`
 * dispatch.
 */

import { ChartType } from '../../enums.js'
import { BARCHART_COLORS, DEF_FONT_COLOR, DEF_FONT_SIZE, DEF_SHAPE_SHADOW } from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { encodeXmlEntities, getUuid } from '../utils.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { createLineCap } from '../drawingml/line.js'
import { ptsToEmuLenient } from '../../units-internal.js'
import { FIXED_PCT_PER_PERCENT, ptToHundredths } from '../../units.js'
import { dataLabels, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, voidEl } from '../oxml/el.js'
import {
	chartColorLineFill,
	chartDataLabels,
	createChartTextFonts,
	makeChartErrorBarsXml,
	makeSeriesDataPointsXml,
	numRefBlock,
	paletteColor,
	resolveChartPalette,
	strRefBlock,
} from './chart-parts.js'

/**
 * Plot an XY scatter chart into `<c:scatterChart>` (paired X/Y numeric series).
 */
export function makeScatterPlot(
	chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valAxisId: string,
	catAxisId: string,
	valFmtCode: string
): string {
	let colorIndex = -1 // Maintain the color index by region
	let strXml = ''
	/*
				`data` = [
					{ name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Value 1', values:[13, 20, 21, 25] },
					{ name:'Y-Value 2', values:[ 1,  2,  5,  9] }
				];
            */

	// 1: Start Chart
	strXml += '<c:' + chartType + 'Chart>'
	strXml += '<c:scatterStyle val="lineMarker"/>'
	strXml += '<c:varyColors val="0"/>'

	// 2: Series: (One for each Y-Axis)
	colorIndex = -1
	data
		.filter((_obj, idx) => idx > 0)
		.forEach((obj, idx) => {
			colorIndex++
			const chartColors = resolveChartPalette(opts)
			strXml += '<c:ser>'
			strXml += `  <c:idx val="${idx}"/>`
			strXml += `  <c:order val="${idx}"/>`
			strXml += strRefBlock(sheetCellRef(idx + 2, 1), obj.name ?? '')

			// 'c:spPr': Fill, Border, Line, LineStyle (dash, etc.), Shadow
			strXml += '  <c:spPr>'
			{
				const tmpSerColor = paletteColor(chartColors, colorIndex)

				if (tmpSerColor === 'transparent') {
					strXml += '<a:noFill/>'
				} else if (opts.chartColorsOpacity) {
					strXml +=
						'<a:solidFill>' +
						createColorElement(
							tmpSerColor,
							'<a:alpha val="' + Math.round(opts.chartColorsOpacity * FIXED_PCT_PER_PERCENT).toString() + '"/>'
						) +
						'</a:solidFill>'
				} else {
					strXml += genXmlColorSelection(tmpSerColor)
				}

				if (opts.lineSize === 0) {
					strXml += '<a:ln><a:noFill/></a:ln>'
				} else {
					strXml += `<a:ln w="${ptsToEmuLenient(opts.lineSize ?? 2)}" cap="${createLineCap(opts.lineCap)}">${chartColorLineFill(tmpSerColor)}`
					strXml +=
						voidEl('a:prstDash', { val: opts.lineDashValues?.[colorIndex] ?? opts.lineDash ?? 'solid' }) +
						voidEl('a:round') +
						'</a:ln>'
				}

				// Shadow
				strXml += createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)
			}
			strXml += '  </c:spPr>'

			// 'c:marker' tag: `lineDataSymbol`
			{
				strXml += '<c:marker>'
				strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>'
				if (opts.lineDataSymbolSize) {
					// Defaults to "auto" otherwise (but this is usually too small, so there is a default)
					strXml += `<c:size val="${opts.lineDataSymbolSize}"/>`
				}
				strXml += '<c:spPr>'
				{
					const markerColor = paletteColor(chartColors, idx)
					strXml += markerColor === 'transparent' ? '<a:noFill/>' : genXmlColorSelection(markerColor)
				}
				strXml += `<a:ln w="${opts.lineDataSymbolLineSize}" cap="flat">${chartColorLineFill(opts.lineDataSymbolLineColor || paletteColor(chartColors, colorIndex))}<a:prstDash val="solid"/><a:round/></a:ln>`
				strXml += '<a:effectLst/>'
				strXml += '</c:spPr>'
				strXml += '</c:marker>'
			}

			// Per-point data points (`c:dPt`) MUST precede `c:dLbls` (CT_ScatterSer schema order).
			// Covers legacy single-series color-vary AND per-point `pointStyles` overrides.
			{
				const scatterVaryColors =
					data.length === 1 && opts.chartColors !== BARCHART_COLORS ? opts.chartColors || BARCHART_COLORS : null
				strXml += makeSeriesDataPointsXml(chartType, obj, opts, scatterVaryColors)
			}

			// Option: scatter data point labels
			//
			// Two GUIDs are minted below, and both are deliberately per-build. `chartUuid` tails
			// each point's `c16:uniqueId`, and each custom label's `a:fld` carries its own: a field
			// id has to be unique, which is a property a derived id would have to reproduce without
			// an oracle for how far that uniqueness has to reach. So the ids stay random and the
			// *comparison* gives, exactly as it already does for `c16:uniqueId` — both patterns are
			// erased by `NORMALIZERS` in `scripts/pptx-parts.mjs` before the byte-identity gate
			// diffs a part. Do not "fix" the nondeterminism here; it is not the palette case.
			if (opts.showLabel) {
				const chartUuid = getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
				if (
					dataLabels(obj)[0] &&
					(opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY')
				) {
					strXml += '<c:dLbls>'
					firstLabelGroup(obj).forEach((label, idx) => {
						if (opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY') {
							strXml += '  <c:dLbl>'
							strXml += `    <c:idx val="${idx}"/>`
							strXml += '    <c:tx>'
							strXml += '      <c:rich>'
							strXml += '            <a:bodyPr>'
							strXml += '                <a:spAutoFit/>'
							strXml += '            </a:bodyPr>'
							strXml += '            <a:lstStyle/>'
							strXml += '            <a:p>'
							strXml += '                <a:pPr>'
							strXml += `                    <a:defRPr sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" b="${opts.dataLabelFontBold ? '1' : '0'}" i="${opts.dataLabelFontItalic ? '1' : '0'}" u="none" strike="noStrike">`
							strXml +=
								'                        <a:solidFill>' +
								createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) +
								'</a:solidFill>'
							strXml += '                        ' + createChartTextFonts(opts.dataLabelFontFace || 'Arial')
							strXml += '                    </a:defRPr>'
							strXml += '                </a:pPr>'
							strXml += '              <a:r>'
							strXml += `                    <a:rPr lang="${encodeXmlEntities(opts.lang || 'en-US')}" sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" b="${opts.dataLabelFontBold ? '1' : '0'}" i="${opts.dataLabelFontItalic ? '1' : '0'}" u="none" strike="noStrike" dirty="0">`
							strXml +=
								'                        <a:solidFill>' +
								createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) +
								'</a:solidFill>'
							strXml += '                        ' + createChartTextFonts(opts.dataLabelFontFace || 'Arial')
							strXml += '                    </a:rPr>'
							strXml += '                    ' + el('a:t', null, label)
							strXml += '              </a:r>'
							// Apply XY values at end of custom label
							// Do not apply the values if the label was empty or just spaces
							// This allows for selective labelling where required
							if (opts.dataLabelFormatScatter === 'customXY' && !/^ *$/.test(label)) {
								strXml += '              <a:r>'
								strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
								strXml += '                  <a:t> (</a:t>'
								strXml += '              </a:r>'
								strXml +=
									'              <a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="XVALUE">'
								strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>'
								strXml += '                  <a:pPr>'
								strXml += '                      <a:defRPr/>'
								strXml += '                  </a:pPr>'
								strXml += '                  ' + el('a:t', null, '[' + (obj.name ?? ''))
								strXml += '              </a:fld>'
								strXml += '              <a:r>'
								strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
								strXml += '                  <a:t>, </a:t>'
								strXml += '              </a:r>'
								strXml +=
									'              <a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="YVALUE">'
								strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>'
								strXml += '                  <a:pPr>'
								strXml += '                      <a:defRPr/>'
								strXml += '                  </a:pPr>'
								strXml += '                  ' + el('a:t', null, '[' + (obj.name ?? '') + ']')
								strXml += '              </a:fld>'
								strXml += '              <a:r>'
								strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
								strXml += '                  <a:t>)</a:t>'
								strXml += '              </a:r>'
								strXml += '              <a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>'
							}
							strXml += '            </a:p>'
							strXml += '      </c:rich>'
							strXml += '    </c:tx>'
							strXml += '    <c:spPr>'
							strXml += '        <a:noFill/>'
							strXml += '        <a:ln>'
							strXml += '            <a:noFill/>'
							strXml += '        </a:ln>'
							strXml += '        <a:effectLst/>'
							strXml += '    </c:spPr>'
							if (opts.dataLabelPosition) strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>'
							strXml += '    <c:showLegendKey val="0"/>'
							strXml += '    <c:showVal val="0"/>'
							strXml += '    <c:showCatName val="0"/>'
							strXml += '    <c:showSerName val="0"/>'
							strXml += '    <c:showPercent val="0"/>'
							strXml += '    <c:showBubbleSize val="0"/>'
							strXml += '       <c:showLeaderLines val="1"/>'
							strXml += '    <c:extLst>'
							strXml +=
								'      <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"/>'
							strXml +=
								'      <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
							strXml += `            <c16:uniqueId val="{${String(idx + 1).padStart(8, '0')}${chartUuid}}"/>`
							strXml += '      </c:ext>'
							strXml += '        </c:extLst>'
							strXml += '</c:dLbl>'
						}
					})
					strXml += '</c:dLbls>'
				}
				if (opts.dataLabelFormatScatter === 'XY') {
					strXml += '<c:dLbls>'
					strXml += '    <c:spPr>'
					strXml += '        <a:noFill/>'
					strXml += '        <a:ln>'
					strXml += '            <a:noFill/>'
					strXml += '        </a:ln>'
					strXml += '          <a:effectLst/>'
					strXml += '    </c:spPr>'
					strXml += '    <c:txPr>'
					strXml += '        <a:bodyPr>'
					strXml += '            <a:spAutoFit/>'
					strXml += '        </a:bodyPr>'
					strXml += '        <a:lstStyle/>'
					strXml += '        <a:p>'
					strXml += '            <a:pPr>'
					strXml += `                <a:defRPr sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" b="${opts.dataLabelFontBold ? '1' : '0'}" i="${opts.dataLabelFontItalic ? '1' : '0'}" u="none" strike="noStrike">`
					strXml +=
						'                    <a:solidFill>' +
						createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) +
						'</a:solidFill>'
					strXml += '                    ' + createChartTextFonts(opts.dataLabelFontFace || 'Arial')
					strXml += '                </a:defRPr>'
					strXml += '            </a:pPr>'
					strXml += `            <a:endParaRPr lang="${encodeXmlEntities(opts.lang || 'en-US')}"/>`
					strXml += '        </a:p>'
					strXml += '    </c:txPr>'
					if (opts.dataLabelPosition) strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>'
					strXml += '    <c:showLegendKey val="0"/>'
					strXml += ` <c:showVal val="${opts.showLabel ? '1' : '0'}"/>`
					strXml += ` <c:showCatName val="${opts.showLabel ? '1' : '0'}"/>`
					strXml += ` <c:showSerName val="${opts.showSerName ? '1' : '0'}"/>`
					strXml += '    <c:showPercent val="0"/>'
					strXml += '    <c:showBubbleSize val="0"/>'
					strXml += '    <c:extLst>'
					strXml +=
						'        <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">'
					strXml += '            <c15:showLeaderLines val="1"/>'
					strXml += '        </c:ext>'
					strXml += '    </c:extLst>'
					strXml += '</c:dLbls>'
				}
			}

			// Error bars (`<c:errBars>`) — schema order places them after dLbls, before xVal/yVal.
			strXml += makeChartErrorBarsXml(chartType, obj.errorBars, obj)

			// 3: "Values": Scatter Chart has 2: `xVal` and `yVal`
			{
				// X-Axis is always the same; the Y series is cached against its length, so a
				// caller who supplied fewer Y values than X leaves gaps rather than a short cache.
				const xValues = dataValues(data[0])
				const yValues = dataValues(obj)
				strXml += numRefBlock('c:xVal', `Sheet1!$A$2:$A$${xValues.length + 1}`, valFmtCode, xValues)
				strXml += numRefBlock(
					'c:yVal',
					sheetRangeRef(idx + 2, 2, idx + 2, xValues.length + 1),
					valFmtCode,
					xValues.map((_value, i) => yValues[i])
				)
			}

			// Option: `smooth`
			strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>'

			// 4: Close "SERIES"
			strXml += '</c:ser>'
		})

	// 3: Data Labels
	strXml += chartDataLabels(opts, false)

	// 4: Add axis Id (NOTE: order matters! - category comes first)
	strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/>`

	// 5: Close Chart tag
	strXml += '</c:' + chartType + 'Chart>'
	return strXml
}
