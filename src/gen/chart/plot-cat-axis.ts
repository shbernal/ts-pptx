/**
 * PptxGenJS: Category-Axis Plot Assembly
 *
 * Emits the `<c:areaChart>` / `<c:barChart>` / `<c:bar3DChart>` / `<c:lineChart>` /
 * `<c:radarChart>` plot elements. These five chart types share one builder because they
 * share a plot shape -- grouping, per-series `<c:ser>` with a category reference and a
 * value cache, then the axis-id pair -- and differ only in a handful of type-gated
 * children. Reached through {@link ./chart-xml}'s `makeChartType` dispatch.
 */

import {
	AXIS_ID_SERIES_PRIMARY,
	BARCHART_COLORS,
	ChartType,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SHAPE_SHADOW,
} from '../../core-enums.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../core-interfaces.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { createLineCap, resolveBorderWidth } from '../drawingml/line.js'
import { valToPts } from '../../units-internal.js'
import { FIXED_PCT_PER_PERCENT, ptToHundredths } from '../../units.js'
import { dataLabels, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, voidEl } from '../oxml/el.js'
import {
	chartColorLineFill,
	createChartTextFonts,
	createSerLinesElement,
	makeChartErrorBarsXml,
	makeCustomDLblXml,
	makeSeriesDataPointsXml,
	numCachePt,
} from './chart-parts.js'

/**
 * Plot a category-axis chart family (area / bar / bar3d / line / radar) into a
 * `<c:xxxChart>` element. These share the grouping / series / cat+val axis structure.
 */
export function makeCatAxisPlot(
	chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valAxisId: string,
	catAxisId: string,
	valFmtCode: string
): string {
	let colorIndex = -1 // Maintain the color index by region
	let strXml = ''
	// 1: Start Chart
	strXml += `<c:${chartType}Chart>`
	if (chartType === ChartType.area || chartType === ChartType.line) {
		const lineGrouping =
			opts.barGrouping === 'stacked' || opts.barGrouping === 'percentStacked' ? opts.barGrouping : 'standard'
		strXml += '<c:grouping val="' + lineGrouping + '"/>'
	}

	if (chartType === ChartType.bar || chartType === ChartType.bar3d) {
		strXml += '<c:barDir val="' + opts.barDir + '"/>'
		strXml += '<c:grouping val="' + (opts.barGrouping || 'clustered') + '"/>'
	}

	if (chartType === ChartType.radar) {
		// Map the public PowerPoint-UI names to ST_RadarStyle wire values.
		const radarStyleWire =
			{ radar: 'standard', markers: 'marker', filled: 'filled' }[opts.radarStyle || 'radar'] ?? 'standard'
		strXml += '<c:radarStyle val="' + radarStyleWire + '"/>'
	}

	strXml += '<c:varyColors val="0"/>'

	// 2: "Series" block for every data row
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
	data.forEach((obj) => {
		colorIndex++
		strXml += '<c:ser>'
		strXml += `  <c:idx val="${obj._dataIndex}"/><c:order val="${obj._dataIndex}"/>`
		strXml += '  <c:tx>'
		strXml += '    <c:strRef>'
		strXml += `      <c:f>${sheetCellRef(obj._dataIndex + dataLabels(obj).length + 1, 1)}</c:f>`
		strXml +=
			'      <c:strCache><c:ptCount val="1"/><c:pt idx="0">' + el('c:v', null, obj.name ?? '') + '</c:pt></c:strCache>'
		strXml += '    </c:strRef>'
		strXml += '  </c:tx>'

		// Fill and Border
		// `chartColors` is always populated by addChartDefinition() (defaulting to BARCHART_COLORS); the
		// fallback here only satisfies the optional type and keeps `seriesColor` a non-null string.
		const chartColors = opts.chartColors?.length ? opts.chartColors : BARCHART_COLORS
		const seriesOverride = opts.seriesOptions?.[obj._dataIndex]
		const seriesColor = seriesOverride?.color ?? chartColors[colorIndex % chartColors.length] ?? '000000'

		strXml += '  <c:spPr>'
		if (seriesColor === 'transparent') {
			strXml += '<a:noFill/>'
		} else if (opts.chartColorsOpacity) {
			strXml +=
				'<a:solidFill>' +
				createColorElement(
					seriesColor,
					`<a:alpha val="${Math.round(opts.chartColorsOpacity * FIXED_PCT_PER_PERCENT)}"/>`
				) +
				'</a:solidFill>'
		} else {
			strXml += genXmlColorSelection(seriesColor)
		}

		if (chartType === ChartType.line || chartType === ChartType.radar) {
			const effectiveLineSize = seriesOverride?.lineSize ?? opts.lineSize ?? 2
			if (effectiveLineSize === 0) {
				strXml += '<a:ln><a:noFill/></a:ln>'
			} else {
				strXml += `<a:ln w="${valToPts(effectiveLineSize)}" cap="${createLineCap(opts.lineCap)}">${chartColorLineFill(seriesColor)}`
				strXml +=
					voidEl('a:prstDash', { val: opts.lineDashValues?.[colorIndex] ?? opts.lineDash ?? 'solid' }) +
					voidEl('a:round') +
					'</a:ln>'
			}
		} else if (opts.dataBorder) {
			strXml += `<a:ln w="${valToPts(resolveBorderWidth(opts.dataBorder, 0.75))}" cap="${createLineCap(opts.lineCap)}">${genXmlColorSelection({ color: opts.dataBorder.color ?? '363636', transparency: opts.dataBorder.transparency })}<a:prstDash val="solid"/><a:round/></a:ln>`
		}

		strXml += createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)

		strXml += '  </c:spPr>'
		// `invertIfNegative` is bar-only in the schema (CT_BarSer); area/line/radar series must omit it
		if (chartType === ChartType.bar || chartType === ChartType.bar3d) strXml += '  <c:invertIfNegative val="0"/>'

		// 'c:marker' must precede 'c:dLbls' in CT_LineSer (schema order: spPr → marker → dPt → dLbls)
		if (chartType === ChartType.line || chartType === ChartType.radar) {
			strXml += '<c:marker>'
			strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>'
			if (opts.lineDataSymbolSize) strXml += `<c:size val="${opts.lineDataSymbolSize}"/>` // Defaults to "auto" otherwise (but this is usually too small, so there is a default)
			strXml += '  <c:spPr>'
			{
				const markerColor =
					chartColors[
						obj._dataIndex + 1 > chartColors.length ? Math.floor(Math.random() * chartColors.length) : obj._dataIndex
					] ?? '000000'
				strXml += markerColor === 'transparent' ? '<a:noFill/>' : genXmlColorSelection(markerColor)
			}
			strXml += `    <a:ln w="${opts.lineDataSymbolLineSize}" cap="flat">${chartColorLineFill(opts.lineDataSymbolLineColor || seriesColor)}<a:prstDash val="solid"/><a:round/></a:ln>`
			strXml += '    <a:effectLst/>'
			strXml += '  </c:spPr>'
			strXml += '</c:marker>'
		}

		// Per-point data points (`c:dPt`) MUST precede `c:dLbls` in CT_*Ser schema order.
		// Covers legacy single-series bar color-vary AND per-point `pointStyles` overrides.
		{
			const barVaryColors =
				(chartType === ChartType.bar || chartType === ChartType.bar3d) &&
				data.length === 1 &&
				((opts.chartColors && opts.chartColors !== BARCHART_COLORS && opts.chartColors.length > 1) ||
					opts.invertedColors?.length)
					? opts.chartColors || BARCHART_COLORS
					: null
			strXml += makeSeriesDataPointsXml(chartType, obj, opts, barVaryColors)
		}

		// Data Labels per series
		// NOTE: [20190117] Adding these to RADAR chart causes unrecoverable corruption!
		if (chartType !== ChartType.radar) {
			const lblColor = seriesOverride?.dataLabelColor ?? opts.dataLabelColor ?? DEF_FONT_COLOR
			const lblBold = seriesOverride?.dataLabelFontBold ?? opts.dataLabelFontBold ?? false
			const lblItalic = seriesOverride?.dataLabelFontItalic ?? opts.dataLabelFontItalic ?? false
			const lblSize = seriesOverride?.dataLabelFontSize ?? opts.dataLabelFontSize ?? DEF_FONT_SIZE
			const lblFace = seriesOverride?.dataLabelFontFace ?? opts.dataLabelFontFace ?? 'Arial'
			const lblFmtCode = seriesOverride?.dataLabelFormatCode ?? opts.dataLabelFormatCode
			strXml += '<c:dLbls>'
			// Per-point custom labels must precede aggregate settings (CT_DLbls schema order: dLbl* then Group_DLbls)
			if (obj.customLabels?.length) {
				obj.customLabels.forEach((lbl, idx) => {
					if (lbl) strXml += makeCustomDLblXml(idx, lbl, opts)
				})
			}
			strXml += voidEl('c:numFmt', { formatCode: (lblFmtCode ?? '') || 'General', sourceLinked: 0 })
			if (opts.dataLabelBkgrdColors) strXml += `<c:spPr>${genXmlColorSelection(seriesColor)}</c:spPr>`
			strXml += '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
			strXml += `<a:defRPr b="${lblBold ? 1 : 0}" i="${lblItalic ? 1 : 0}" strike="noStrike" sz="${ptToHundredths(lblSize)}" u="none">`
			strXml += genXmlColorSelection(lblColor)
			strXml += createChartTextFonts(lblFace)
			strXml += '</a:defRPr></a:pPr></a:p></c:txPr>'
			if (opts.dataLabelPosition) strXml += voidEl('c:dLblPos', { val: opts.dataLabelPosition })
			strXml += '<c:showLegendKey val="0"/>'
			strXml += `<c:showVal val="${opts.showValue ? '1' : '0'}"/>`
			strXml += `<c:showCatName val="0"/><c:showSerName val="${opts.showSerName ? '1' : '0'}"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>`
			strXml += `<c:showLeaderLines val="${opts.showLeaderLines ? '1' : '0'}"/>`
			strXml += '</c:dLbls>'
		}

		// Error bars (`<c:errBars>`) — schema order places them after dLbls, before cat.
		// RADAR has no error bars in CT_RadarSer, so it is excluded.
		if (chartType !== ChartType.radar) strXml += makeChartErrorBarsXml(chartType, obj.errorBars, obj)

		// 2: "Categories"
		{
			strXml += '<c:cat>'
			if (opts.catLabelFormatCode) {
				// Use 'numRef' as catLabelFormatCode implies that we are expecting numbers here
				strXml += '  <c:numRef>'
				strXml += `    <c:f>Sheet1!$A$2:$A$${firstLabelGroup(obj).length + 1}</c:f>`
				strXml += '    <c:numCache>'
				strXml += '      <c:formatCode>' + (opts.catLabelFormatCode || 'General') + '</c:formatCode>'
				strXml += `      <c:ptCount val="${firstLabelGroup(obj).length}"/>`
				firstLabelGroup(obj).forEach((label, idx) => (strXml += `<c:pt idx="${idx}">${el('c:v', null, label)}</c:pt>`))
				strXml += '    </c:numCache>'
				strXml += '  </c:numRef>'
			} else if (dataLabels(obj).length === 1) {
				strXml += '  <c:strRef>'
				strXml += `    <c:f>Sheet1!$A$2:$A$${firstLabelGroup(obj).length + 1}</c:f>`
				strXml += '    <c:strCache>'
				strXml += `      <c:ptCount val="${firstLabelGroup(obj).length}"/>`
				firstLabelGroup(obj).forEach((label, idx) => (strXml += `<c:pt idx="${idx}">${el('c:v', null, label)}</c:pt>`))
				strXml += '    </c:strCache>'
				strXml += '  </c:strRef>'
			} else {
				strXml += '  <c:multiLvlStrRef>'
				strXml += `    <c:f>${sheetRangeRef(1, 2, dataLabels(obj).length, firstLabelGroup(obj).length + 1)}</c:f>`
				strXml += '    <c:multiLvlStrCache>'
				strXml += `      <c:ptCount val="${firstLabelGroup(obj).length}"/>`
				dataLabels(obj).forEach((labelsGroup) => {
					strXml += '<c:lvl>'
					labelsGroup.forEach((label, idx) => (strXml += `<c:pt idx="${idx}">${el('c:v', null, label)}</c:pt>`))
					strXml += '</c:lvl>'
				})
				strXml += '    </c:multiLvlStrCache>'
				strXml += '  </c:multiLvlStrRef>'
			}
			strXml += '</c:cat>'
		}

		// 3: "Values"
		{
			strXml += '<c:val>'
			strXml += '  <c:numRef>'
			strXml += `<c:f>${sheetRangeRef(obj._dataIndex + dataLabels(obj).length + 1, 2, obj._dataIndex + dataLabels(obj).length + 1, firstLabelGroup(obj).length + 1)}</c:f>`
			strXml += '    <c:numCache>'
			strXml += '      <c:formatCode>' + valFmtCode + '</c:formatCode>'
			strXml += `      <c:ptCount val="${firstLabelGroup(obj).length}"/>`
			dataValues(obj).forEach((value, idx) => {
				strXml += numCachePt(idx, value)
			})
			strXml += '    </c:numCache>'
			strXml += '  </c:numRef>'
			strXml += '</c:val>'
		}

		// Option: `smooth`
		if (chartType === ChartType.line) strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>'

		// 4: Close "SERIES"
		strXml += '</c:ser>'
	})

	// 3: "Data Labels"
	{
		strXml += '  <c:dLbls>'
		strXml +=
			'    ' + voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })
		strXml += '    <c:txPr>'
		strXml += '      <a:bodyPr/>'
		strXml += '      <a:lstStyle/>'
		strXml += '      <a:p><a:pPr>'
		strXml += `        <a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" u="none">`
		strXml += genXmlColorSelection(opts.dataLabelColor || DEF_FONT_COLOR)
		strXml += '          ' + createChartTextFonts(opts.dataLabelFontFace || 'Arial')
		strXml += '        </a:defRPr>'
		strXml += '      </a:pPr></a:p>'
		strXml += '    </c:txPr>'
		if (opts.dataLabelPosition) strXml += ' <c:dLblPos val="' + opts.dataLabelPosition + '"/>'
		strXml += '    <c:showLegendKey val="0"/>'
		strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
		strXml += '    <c:showCatName val="0"/>'
		strXml += '    <c:showSerName val="' + (opts.showSerName ? '1' : '0') + '"/>'
		strXml += '    <c:showPercent val="0"/>'
		strXml += '    <c:showBubbleSize val="0"/>'
		strXml += `    <c:showLeaderLines val="${opts.showLeaderLines ? '1' : '0'}"/>`
		strXml += '  </c:dLbls>'
	}

	// 4: Add more chart options (gapWidth, line Marker, etc.)
	if (chartType === ChartType.bar) {
		strXml += `  <c:gapWidth val="${opts.barGapWidthPct}"/>`
		strXml += `  <c:overlap val="${opts.barOverlapPct != null ? opts.barOverlapPct : (opts.barGrouping || '').includes('tacked') ? 100 : 0}"/>`
		// `<c:serLines>` ("Series Lines") connects data points across stacked bar/column series.
		// Schema order (CT_BarChart): gapWidth → overlap → serLines → axId.
		strXml += createSerLinesElement(opts.barSeriesLine)
	} else if (chartType === ChartType.bar3d) {
		strXml += `  <c:gapWidth val="${opts.barGapWidthPct}"/>`
		strXml += `  <c:gapDepth val="${opts.barGapDepthPct}"/>`
		strXml += '  <c:shape val="' + opts.bar3DShape + '"/>'
	} else if (chartType === ChartType.line) {
		strXml += '  <c:marker val="1"/>'
	}

	// 5: Add axisId (NOTE: order matters! (category comes first))
	// Only 3D charts (BAR3D) get a series axis def; emitting a
	// SERIES_PRIMARY axId for 2D charts produced a dangling reference
	// that violated the OOXML invariant (every axId in <c:plotArea>
	// must resolve to a defined catAx/valAx).
	strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/>`
	if (chartType === ChartType.bar3d) {
		strXml += `<c:axId val="${AXIS_ID_SERIES_PRIMARY}"/>`
	}

	// 6: Close Chart tag
	strXml += `</c:${chartType}Chart>`
	return strXml
}
