/**
 * ts-pptx: Chart DrawingML Assembly
 *
 * Builds a chart's `ppt/charts/chartN.xml` -- the `<c:chartSpace>` DrawingML that
 * PowerPoint renders. `makeXmlCharts` assembles the top-level envelope (header, plot
 * area, axes region, legend, metadata) and `makeChartType` dispatches the plot itself to
 * the per-family builder: {@link ./plot-cat-axis} for area/bar/bar3D/line/radar,
 * {@link ./plot-scatter}, {@link ./plot-bubble}, {@link ./plot-pie}. Axes come from
 * {@link ./chart-axes} and the shared leaf fragments from {@link ./chart-parts}.
 *
 * Every function in this directory is a pure string builder -- no I/O, no mutation of
 * the presentation model. The `<c:f>` series formulas here point back at the cells
 * written by the embedded workbook ({@link ./embed-xlsx}); that mapping lives in
 * {@link ./data-refs}.
 */

import { asChartType, ChartType } from '../../core-enums.js'
import {
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_SERIES_PRIMARY,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
	DEF_FONT_SIZE,
	DEF_FONT_TITLE_SIZE,
	XML_DECL,
} from '../../core-enums-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal, SlideRelChart } from '../../types/internal.js'
import { warn } from '../../diagnostics.js'
import { encodeXmlEntities } from '../../gen-utils.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { resolveBorderWidth } from '../drawingml/line.js'
import { valToPts } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { createChartTextFonts, genXmlTitle } from './chart-parts.js'
import { makeCatAxis, makeSerAxis, makeValAxis } from './chart-axes.js'
import { makeCatAxisPlot } from './plot-cat-axis.js'
import { makeScatterPlot } from './plot-scatter.js'
import { makeBubblePlot } from './plot-bubble.js'
import { makePiePlot } from './plot-pie.js'
import { isVolumeStockStyle, makeStockPlot } from './plot-stock.js'
import { makeSurfacePlot, makeSurfaceScene } from './plot-surface.js'

/**
 * Build the chartSpace/chart header: chartSpace open, title (or autoTitleDeleted),
 * optional 3D view, and the plotArea open with optional manual layout.
 */
function makeChartHeaderXml(rel: SlideRelChart): string {
	const chartArea = rel.opts.chartArea ?? {}
	let strXml = ''
	// CHARTSPACE: BEGIN vvv
	strXml +=
		'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
	strXml += voidEl('c:date1904', { val: 0 }) // ppt defaults to 1904 dates, excel to 1900
	strXml += voidEl('c:roundedCorners', { val: chartArea.roundedCorners ? 1 : 0 })
	strXml += '<c:chart>'

	// OPTION: Title
	if (rel.opts.showTitle) {
		strXml += genXmlTitle(
			{
				title: rel.opts.title || 'Chart Title',
				color: rel.opts.titleColor,
				fontFace: rel.opts.titleFontFace,
				fontSize: rel.opts.titleFontSize || DEF_FONT_TITLE_SIZE,
				titleAlign: rel.opts.titleAlign,
				titleBold: rel.opts.titleBold,
				titleItalic: rel.opts.titleItalic,
				titleUnderline: rel.opts.titleUnderline,
				titlePos: rel.opts.titlePos,
				titleRotate: rel.opts.titleRotate,
			},
			rel.opts.x as number,
			rel.opts.y as number
		)
		strXml += voidEl('c:autoTitleDeleted', { val: 0 })
	} else {
		// NOTE: Add autoTitleDeleted tag in else to prevent default creation of chart title even when showTitle is set to false
		strXml += voidEl('c:autoTitleDeleted', { val: 1 })
	}
	/** Add 3D view tag
	 * @see: https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_perspective_topic_ID0E6BUQB.html
	 */
	if (rel.opts._type === ChartType.bar3d) {
		strXml += el('c:view3D', null, [
			raw(voidEl('c:rotX', { val: rel.opts.v3DRotX })),
			raw(voidEl('c:rotY', { val: rel.opts.v3DRotY })),
			raw(voidEl('c:rAngAx', { val: !rel.opts.v3DRAngAx ? 0 : 1 })),
			raw(voidEl('c:perspective', { val: rel.opts.v3DPerspective })),
		])
	} else if (rel.opts._type === ChartType.surface) {
		// A surface chart is a 3-D scene: view3D + floor/side/back walls precede the plotArea.
		strXml += makeSurfaceScene(rel.opts)
	}

	strXml += '<c:plotArea>'
	// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
	if (rel.opts.layout) {
		const manualLayout = el(
			'c:manualLayout',
			null,
			[
				raw(voidEl('c:layoutTarget', { val: 'inner' }, { closePrefix: ' ' })),
				raw(voidEl('c:xMode', { val: 'edge' }, { closePrefix: ' ' })),
				raw(voidEl('c:yMode', { val: 'edge' }, { closePrefix: ' ' })),
				raw(voidEl('c:x', { val: rel.opts.layout.x || 0 }, { closePrefix: ' ' })),
				raw(voidEl('c:y', { val: rel.opts.layout.y || 0 }, { closePrefix: ' ' })),
				raw(voidEl('c:w', { val: rel.opts.layout.w || 1 }, { closePrefix: ' ' })),
				raw(voidEl('c:h', { val: rel.opts.layout.h || 1 }, { closePrefix: ' ' })),
			],
			{ childPrefix: '  ', closePrefix: ' ' }
		)
		strXml += el('c:layout', null, raw(manualLayout), { childPrefix: ' ' })
	} else {
		strXml += voidEl('c:layout')
	}
	return strXml
}

/**
 * Build the category/value/series axis XML (empty for pie/doughnut). Resolves combo-chart
 * category axes to val axes when owned by a scatter/bubble subchart, per the tracked flags.
 */
function makeChartAxesXml(
	rel: SlideRelChart,
	usesSecondaryValAxis: boolean,
	usesSecondaryCatAxis: boolean,
	primaryCatAxisValType: ChartType | null,
	secondaryCatAxisValType: ChartType | null,
	primaryCatAxisHasCategoryChart: boolean,
	secondaryCatAxisHasCategoryChart: boolean
): string {
	let strXml = ''
	// Stock charts drive their own (date/category)+value axis pair. The non-volume styles use just
	// the primary pair (the generic path below would also produce it), but the volume styles add a
	// secondary pair for the price series (the volume bar owns the primary pair) — so route all stock
	// axes here, reusing the shared cat/val emitters. `barDir` defaults to 'col' in the define layer,
	// which already gives the correct axis positions (category at bottom, value at left).
	if (rel.opts._type === ChartType.stock) {
		let stockXml = makeCatAxis(rel.opts, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY)
		stockXml += makeValAxis(rel.opts, AXIS_ID_VALUE_PRIMARY)
		if (isVolumeStockStyle(rel.opts.stockStyle)) {
			// Secondary value axis (right) for the price series, then a hidden secondary category axis.
			stockXml += makeValAxis(rel.opts, AXIS_ID_VALUE_SECONDARY)
			stockXml += makeCatAxis({ ...rel.opts, catAxisHidden: true }, AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_VALUE_SECONDARY)
		}
		return stockXml
	}
	if (rel.opts._type !== ChartType.pie && rel.opts._type !== ChartType.doughnut) {
		// Param check
		if (rel.opts.valAxes && rel.opts.valAxes.length > 1 && !usesSecondaryValAxis) {
			throw new Error('Secondary axis must be used by one of the multiple charts')
		}

		// Resolve the effective `_type` for a combo category axis so scatter/bubble
		// subcharts get a `<c:valAx>` X axis. Returns the scatter/bubble type when
		// that axis is owned only by such a subchart, else null (category axis).
		const comboCatAxisType = (isSecondary: boolean): { _type: ChartType } | Record<string, never> => {
			const valType = isSecondary ? secondaryCatAxisValType : primaryCatAxisValType
			const hasCategoryChart = isSecondary ? secondaryCatAxisHasCategoryChart : primaryCatAxisHasCategoryChart
			if (!valType) return {}
			if (hasCategoryChart) {
				// A category-based chart and a scatter/bubble chart cannot share one
				// axis (one needs <c:catAx>, the other <c:valAx>). Keep the category
				// axis and warn rather than silently emit a repair-triggering file.
				warn(
					'chart/axis-type-conflict',
					`A category-based chart and a scatter/bubble chart cannot share the same ${isSecondary ? 'secondary' : 'primary'} category axis; emitting a category axis. Put the scatter/bubble series on a separate axis.`
				)
				return {}
			}
			return { _type: valType }
		}

		if (rel.opts.catAxes) {
			if (!rel.opts.valAxes || rel.opts.valAxes.length !== rel.opts.catAxes.length) {
				throw new Error('There must be the same number of value and category axes.')
			}
			strXml += makeCatAxis(
				{ ...rel.opts, ...rel.opts.catAxes[0], ...comboCatAxisType(false) },
				AXIS_ID_CATEGORY_PRIMARY,
				AXIS_ID_VALUE_PRIMARY
			)
		} else {
			strXml += makeCatAxis(
				{ ...rel.opts, ...comboCatAxisType(false) },
				AXIS_ID_CATEGORY_PRIMARY,
				AXIS_ID_VALUE_PRIMARY
			)
		}

		if (rel.opts.valAxes) {
			strXml += makeValAxis({ ...rel.opts, ...rel.opts.valAxes[0] }, AXIS_ID_VALUE_PRIMARY)
			if (rel.opts.valAxes[1]) {
				strXml += makeValAxis({ ...rel.opts, ...rel.opts.valAxes[1] }, AXIS_ID_VALUE_SECONDARY)
			}
		} else {
			strXml += makeValAxis(rel.opts, AXIS_ID_VALUE_PRIMARY)

			// Add series axis for 3D bar and surface (both plot over a category × series grid)
			if (rel.opts._type === ChartType.bar3d || rel.opts._type === ChartType.surface) {
				strXml += makeSerAxis(rel.opts, AXIS_ID_SERIES_PRIMARY, AXIS_ID_VALUE_PRIMARY)
			}

			// For combo charts referencing a secondary value axis via the
			// `secondaryValAxis: true` flag (without a `valAxes` array),
			// auto-synthesise the missing secondary value axis def so that
			// the axId references in <c:plotArea> all resolve.
			if (usesSecondaryValAxis) {
				strXml += makeValAxis(rel.opts, AXIS_ID_VALUE_SECONDARY)
			}
		}

		// Combo Charts: Add secondary axes after all vals
		if (rel.opts?.catAxes && rel.opts?.catAxes[1]) {
			strXml += makeCatAxis(
				{ ...rel.opts, ...rel.opts.catAxes[1], ...comboCatAxisType(true) },
				AXIS_ID_CATEGORY_SECONDARY,
				AXIS_ID_VALUE_SECONDARY
			)
		} else if (usesSecondaryCatAxis && (!rel.opts.catAxes || !rel.opts.catAxes[1])) {
			// Same as above for the secondary category axis.
			strXml += makeCatAxis(
				{ ...rel.opts, ...comboCatAxisType(true) },
				AXIS_ID_CATEGORY_SECONDARY,
				AXIS_ID_VALUE_SECONDARY
			)
		}
	}
	return strXml
}

/**
 * Build the plotArea properties (data table, fill, border), close plotArea, then the legend.
 */
function makeChartPlotAreaPropsXml(rel: SlideRelChart): string {
	const plotArea = rel.opts.plotArea ?? {}
	let strXml = ''
	// NOTE: DataTable goes between '</c:valAx>' and '<c:spPr>'
	if (rel.opts.showDataTable) {
		strXml += '<c:dTable>'
		strXml += `  <c:showHorzBorder val="${!rel.opts.showDataTableHorzBorder ? 0 : 1}"/>`
		strXml += `  <c:showVertBorder val="${!rel.opts.showDataTableVertBorder ? 0 : 1}"/>`
		strXml += `  <c:showOutline    val="${!rel.opts.showDataTableOutline ? 0 : 1}"/>`
		strXml += `  <c:showKeys       val="${!rel.opts.showDataTableKeys ? 0 : 1}"/>`
		strXml += '  <c:spPr>'
		strXml += '    <a:noFill/>'
		strXml +=
			'    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>'
		strXml += '    <a:effectLst/>'
		strXml += '  </c:spPr>'
		strXml += '  <c:txPr>'
		strXml +=
			'   <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>'
		strXml += '   <a:lstStyle/>'
		strXml += '   <a:p>'
		strXml += '     <a:pPr rtl="0">'
		strXml += `       <a:defRPr sz="${ptToHundredths(rel.opts.dataTableFontSize || DEF_FONT_SIZE)}" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">`
		strXml +=
			'         <a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>'
		strXml += '         <a:latin typeface="+mn-lt"/>'
		strXml += '         <a:ea typeface="+mn-ea"/>'
		strXml += '         <a:cs typeface="+mn-cs"/>'
		strXml += '       </a:defRPr>'
		strXml += '     </a:pPr>'
		strXml += '    <a:endParaRPr lang="en-US"/>'
		strXml += '   </a:p>'
		strXml += ' </c:txPr>'
		strXml += '</c:dTable>'
	}

	strXml += '  <c:spPr>'

	// OPTION: Fill
	strXml += plotArea.fill?.color ? genXmlColorSelection(plotArea.fill) : '<a:noFill/>'

	// OPTION: Border
	strXml += plotArea.border
		? `<a:ln w="${valToPts(resolveBorderWidth(plotArea.border, 1))}" cap="flat">${genXmlColorSelection({ color: plotArea.border.color ?? '363636', transparency: plotArea.border.transparency })}</a:ln>`
		: '<a:ln><a:noFill/></a:ln>'

	// Close shapeProp/plotArea before Legend
	strXml += '    <a:effectLst/>'
	strXml += '  </c:spPr>'
	strXml += '</c:plotArea>'

	// OPTION: Legend
	// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
	if (rel.opts.showLegend) {
		strXml += '<c:legend>'
		strXml += '<c:legendPos val="' + rel.opts.legendPos + '"/>'
		// For combo charts: suppress series from subcharts that set showLegend: false
		if (Array.isArray(rel.opts._type)) {
			let seriesIdx = 0
			rel.opts._type.forEach((type) => {
				if (type.options?.showLegend === false) {
					for (let i = 0; i < type.data.length; i++) {
						strXml += `<c:legendEntry><c:idx val="${seriesIdx + i}"/><c:delete val="1"/></c:legendEntry>`
					}
				}
				seriesIdx += type.data.length
			})
		}
		// OPTION: Manual legend placement
		// Each axis of CT_ManualLayout is independent: omitting xMode/x (or
		// yMode/y, etc.) leaves that axis on automatic layout. x/y use edge
		// mode so they are absolute fractions of the chart; w/h are fractions
		// of the chart size. Schema order: xMode, yMode, x, y, w, h.
		const legendLayout = rel.opts.legendLayout
		const hasLegendX = legendLayout && typeof legendLayout.x === 'number'
		const hasLegendY = legendLayout && typeof legendLayout.y === 'number'
		const hasLegendW = legendLayout && typeof legendLayout.w === 'number'
		const hasLegendH = legendLayout && typeof legendLayout.h === 'number'
		if (hasLegendX || hasLegendY || hasLegendW || hasLegendH) {
			let modes = ''
			let vals = ''
			if (hasLegendX) {
				modes += '<c:xMode val="edge"/>'
				vals += `<c:x val="${legendLayout.x}"/>`
			}
			if (hasLegendY) {
				modes += '<c:yMode val="edge"/>'
				vals += `<c:y val="${legendLayout.y}"/>`
			}
			if (hasLegendW) vals += `<c:w val="${legendLayout.w}"/>`
			if (hasLegendH) vals += `<c:h val="${legendLayout.h}"/>`
			strXml += `<c:layout><c:manualLayout>${modes}${vals}</c:manualLayout></c:layout>`
		}
		strXml += '<c:overlay val="0"/>'
		if (rel.opts.legendFontFace || rel.opts.legendFontSize || rel.opts.legendColor) {
			strXml += '<c:txPr>'
			strXml += '  <a:bodyPr/>'
			strXml += '  <a:lstStyle/>'
			strXml += '  <a:p>'
			strXml += '    <a:pPr>'
			strXml += rel.opts.legendFontSize
				? `<a:defRPr sz="${ptToHundredths(Number(rel.opts.legendFontSize))}">`
				: '<a:defRPr>'
			if (rel.opts.legendColor) strXml += genXmlColorSelection(rel.opts.legendColor)
			if (rel.opts.legendFontFace) strXml += createChartTextFonts(rel.opts.legendFontFace)
			strXml += '      </a:defRPr>'
			strXml += '    </a:pPr>'
			strXml += '    <a:endParaRPr lang="en-US"/>'
			strXml += '  </a:p>'
			strXml += '</c:txPr>'
		}
		strXml += '</c:legend>'
	}
	return strXml
}

/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 * @param {SlideRelChart} rel - chart object
 * @return {string} XML
 */
export function makeXmlCharts(rel: SlideRelChart): string {
	let strXml = XML_DECL
	// `chartArea`/`plotArea` are always populated by addChartDefinition() but stay optional on the type.
	const chartArea = rel.opts.chartArea ?? {}
	let usesSecondaryValAxis = false
	let usesSecondaryCatAxis = false
	// Combo charts: a scatter/bubble subchart draws numbers on its category (X)
	// axis, so that axis must be emitted as a `<c:valAx>` rather than a `<c:catAx>`
	// or PowerPoint flags the file for repair. Track, per category axis,
	// the scatter/bubble subchart type that owns it (if any) and whether a
	// category-based subchart also references it (an unsatisfiable conflict).
	let primaryCatAxisValType: ChartType | null = null
	let secondaryCatAxisValType: ChartType | null = null
	let primaryCatAxisHasCategoryChart = false
	let secondaryCatAxisHasCategoryChart = false

	// STEP 1: Create chart
	strXml += makeChartHeaderXml(rel)

	// STEP 2: Create chart-type XML (plot each subchart's series/points)
	if (Array.isArray(rel.opts._type)) {
		rel.opts._type.forEach((type) => {
			const options = { ...rel.opts, ...type.options }
			const valAxisId = options.secondaryValAxis ? AXIS_ID_VALUE_SECONDARY : AXIS_ID_VALUE_PRIMARY
			const catAxisId = options.secondaryCatAxis ? AXIS_ID_CATEGORY_SECONDARY : AXIS_ID_CATEGORY_PRIMARY
			usesSecondaryValAxis = usesSecondaryValAxis || (options.secondaryValAxis ?? false)
			usesSecondaryCatAxis = usesSecondaryCatAxis || (options.secondaryCatAxis ?? false)
			const subType = asChartType(type.type)
			// Record whether this subchart needs a value-based X axis (scatter/bubble)
			// or a category-based X axis, keyed to the primary/secondary cat axis it uses.
			const usesValueXAxis =
				subType === ChartType.scatter || subType === ChartType.bubble || subType === ChartType.bubble3d
			if (options.secondaryCatAxis) {
				if (usesValueXAxis) secondaryCatAxisValType = subType
				else secondaryCatAxisHasCategoryChart = true
			} else {
				if (usesValueXAxis) primaryCatAxisValType = subType
				else primaryCatAxisHasCategoryChart = true
			}
			strXml += makeChartType(subType, type.data as OptsChartDataInternal[], options, valAxisId, catAxisId)
		})
	} else if (rel.opts._type) {
		strXml += makeChartType(rel.opts._type, rel.data, rel.opts, AXIS_ID_VALUE_PRIMARY, AXIS_ID_CATEGORY_PRIMARY)
	}

	// STEP 3: Axes
	strXml += makeChartAxesXml(
		rel,
		usesSecondaryValAxis,
		usesSecondaryCatAxis,
		primaryCatAxisValType,
		secondaryCatAxisValType,
		primaryCatAxisHasCategoryChart,
		secondaryCatAxisHasCategoryChart
	)

	// STEP 4: Chart properties and plotArea options: border, data table, fill, legend
	strXml += makeChartPlotAreaPropsXml(rel)

	strXml += '  <c:plotVisOnly val="1"/>'
	strXml += '  <c:dispBlanksAs val="' + rel.opts.displayBlanksAs + '"/>'
	if (rel.opts._type === ChartType.scatter) strXml += '<c:showDLblsOverMax val="1"/>'

	strXml += '</c:chart>'

	// STEP 5: chartSpace shape props
	strXml += '<c:spPr>'
	strXml += chartArea.fill?.color ? genXmlColorSelection(chartArea.fill) : '<a:noFill/>'
	strXml += chartArea.border
		? `<a:ln w="${valToPts(resolveBorderWidth(chartArea.border, 1))}" cap="flat">${genXmlColorSelection({ color: chartArea.border.color ?? '363636', transparency: chartArea.border.transparency })}</a:ln>`
		: '<a:ln><a:noFill/></a:ln>'
	strXml += '  <a:effectLst/>'
	strXml += '</c:spPr>'

	// STEP 6: Data (add relId)
	strXml += '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'

	// STEP 7: Metadata (custom chart-level annotations via the schema-valid extension list)
	// CT_ChartSpace document order: externalData → printSettings → userShapes → extLst (extLst LAST).
	strXml += genXmlChartMetadata(rel.opts.metadata)

	// LAST: chartSpace end
	strXml += '</c:chartSpace>'

	return strXml
}

/**
 * Stable ts-pptx vendor GUID identifying the chart-metadata extension on `c:chartSpace/c:extLst`.
 * Custom data rides under this URI in a foreign namespace so PowerPoint preserves it (the extLst
 * mechanism) instead of stripping/repairing it as it would an unrecognised sibling element.
 */
const CHART_METADATA_EXT_URI = '{094A432E-1F6C-499B-95B8-B57DC9536949}'

/** Foreign namespace for the chart-metadata extension payload. */
const CHART_METADATA_NS = 'http://ts-pptx.com/schema/chart/metadata'

/**
 * Create the chart-space extension-list XML carrying custom `metadata` key/value annotations.
 * Emits nothing when metadata is absent or contains no valid entries. Keys must be non-empty
 * strings and values must be strings; invalid entries are dropped with a warning rather than
 * emitting degenerate XML (per the "no silent coercion" policy).
 * @param {Record<string, string>} [metadata] custom chart-level metadata
 * @return {string} `<c:extLst>…</c:extLst>` XML, or '' when there is nothing valid to emit
 */
function genXmlChartMetadata(metadata?: Record<string, string>): string {
	if (metadata == null) return ''
	if (typeof metadata !== 'object' || Array.isArray(metadata)) {
		warn('chart/invalid-metadata', 'chart `metadata` must be a plain object of string key/value pairs; ignored.')
		return ''
	}

	let items = ''
	for (const [key, value] of Object.entries(metadata)) {
		if (typeof key !== 'string' || key.length === 0) {
			warn(
				'chart/invalid-metadata-key',
				`chart metadata key "${String(key)}" is not a non-empty string; entry skipped.`
			)
			continue
		}
		if (typeof value !== 'string') {
			warn('chart/invalid-metadata-value', `chart metadata value for key "${key}" is not a string; entry skipped.`)
			continue
		}
		items += voidEl('pgm:item', { key, value })
	}
	if (items === '') return ''

	return el(
		'c:extLst',
		null,
		raw(
			el(
				'c:ext',
				{ uri: CHART_METADATA_EXT_URI },
				raw(el('pgm:metadata', { 'xmlns:pgm': CHART_METADATA_NS }, raw(items)))
			)
		)
	)
}

/**
 * Create XML string for any given chart type.
 * Dispatches to a per-family plot helper; resolves the effective value number format first.
 * @param {ChartType} chartType chart type name
 * @param {OptsChartDataInternal[]} data chart data
 * @param {ChartOptsInternal} opts chart options
 * @param {string} valAxisId chart val axis id
 * @param {string} catAxisId chart cat axis id
 * @example 'bubble' returns <c:bubbleChart></c>
 * @example '<c:lineChart>'
 * @return {string} XML chart
 */

function makeChartType(
	chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valAxisId: string,
	catAxisId: string
): string {
	// NOTE: "Chart Range" (as shown in "select Chart Area dialog") is calculated.
	// ....: Ensure each X/Y Axis/Col has same row height (esp. applicable to XY Scatter where X can often be larger than Y's)
	//
	// PowerPoint and Google Slides render values using the cached *source* number format carried in
	// each series' `<c:numCache><c:formatCode>` (mirroring the embedded workbook cell format), NOT the
	// `<c:dLbls><c:numFmt>` mask — so when the value cache is left as "General" those engines display
	// raw values (e.g. `0.1` instead of `10%`) even though LibreOffice honors the dLbls mask. Resolve a
	// single effective value format and stamp it onto every value cache below so all three engines agree.
	// Precedence keeps the historical `valLabelFormatCode` winner,
	// then the data-table format, and finally falls back to `dataLabelFormatCode` (the most common knob).
	const valFmtCode = encodeXmlEntities(
		opts.valLabelFormatCode || opts.dataTableFormatCode || opts.dataLabelFormatCode || 'General'
	)

	switch (chartType) {
		case ChartType.area:
		case ChartType.bar:
		case ChartType.bar3d:
		case ChartType.line:
		case ChartType.radar:
			return makeCatAxisPlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.scatter:
			return makeScatterPlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.bubble:
		case ChartType.bubble3d:
			return makeBubblePlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.stock:
			return makeStockPlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.surface:
			return makeSurfacePlot(chartType, data, opts, valAxisId, catAxisId, valFmtCode)
		case ChartType.doughnut:
		case ChartType.pie:
			return makePiePlot(chartType, data, opts, valFmtCode)
		default:
			return ''
	}
}
