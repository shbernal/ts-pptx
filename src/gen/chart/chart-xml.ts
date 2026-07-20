/**
 * PptxGenJS: Chart DrawingML Assembly
 *
 * Builds a chart's `ppt/charts/chartN.xml` — the `<c:chartSpace>` DrawingML that
 * PowerPoint renders. `makeXmlCharts` assembles the top-level envelope; the rest are
 * the per-region fragment builders it delegates to (chart-type plot areas, axes,
 * titles, gridlines, error bars, data labels, number caches). Every function is a
 * pure string builder — no I/O, no mutation of the presentation model. The `<c:f>`
 * series formulas here point back at the cells written by the embedded workbook
 * ({@link ./embed-xlsx}); that mapping lives in {@link ./data-refs}.
 */

import {
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_SERIES_PRIMARY,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
	asChartType,
	BARCHART_COLORS,
	ChartType,
	DEF_CHART_GRIDLINE,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_FONT_TITLE_SIZE,
	DEF_SHAPE_SHADOW,
	ONEPT,
	XML_DECL,
} from '../../core-enums.js'
import type {
	ChartOptsInternal,
	SlideRelChart,
	ChartPropsTitle,
	OptsChartGridLine,
	OptsChartDataInternal,
	BorderProps,
	ChartErrorBarOptions,
} from '../../core-interfaces.js'
import { warn } from '../../log.js'
import {
	createColorElement,
	createLineCap,
	createShadowEffectLst,
	genXmlColorSelection,
	genXmlPatternFill,
	convertRotationDegrees,
	encodeXmlEntities,
	getUuid,
	resolveBorderWidth,
	valToPts,
} from '../../gen-utils.js'
import { FIXED_PCT_PER_PERCENT, ptToHundredths } from '../../units.js'
import { dataLabels, dataValues, dataSizes, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'

const VALID_CHART_TIME_UNITS = ['days', 'months', 'years']

// DEF_CHART_GRIDLINE.color is optional on the type but always present on the constant.
const DEF_GRIDLINE_COLOR: string = DEF_CHART_GRIDLINE.color ?? '888888'

/**
 * Fill fragment for a `chartColors`-derived series/line/marker colour.
 *
 * A `'transparent'` entry means "no fill" — an invisible series, connecting line, or marker
 * stroke — and maps to `<a:noFill/>`. Any real colour goes through the normal solid-fill path.
 * The series and marker *fill* paths already special-case `'transparent'`; without this the
 * *stroke* paths (`<a:ln>` on the series line and marker border) would instead pass the literal
 * `'transparent'` through colour validation, warn "not a valid scheme color or hex RGB", and
 * render as black — leaving a stray black line/border where a transparent series was requested.
 */
function chartColorLineFill(color: string): string {
	return color === 'transparent' ? voidEl('a:noFill') : genXmlColorSelection(color)
}
/**
 * Emit the `<a:latin>/<a:ea>/<a:cs>` font trio for a chart text run.
 *
 * In DrawingML run properties a typeface applies only to the script class of
 * its element: `<a:latin>` covers Latin/ASCII, `<a:ea>` covers East Asian, and
 * `<a:cs>` covers complex scripts. Emitting `<a:latin>` alone leaves East Asian
 * (e.g. Chinese) and complex-script glyphs falling back to the theme font, so a
 * user-specified font never takes effect for that text — most visibly on
 * PowerPoint for Mac. Stamping the same typeface onto all three classes is what
 * choosing a font in PowerPoint's UI does.
 * @param {string} typeface - font face name
 * @return {string} `<a:latin/><a:ea/><a:cs/>` XML
 */
// ===== Chart XML assembly =====

function createChartTextFonts(typeface: string): string {
	// Every caller passes a caller-supplied font option (dataLabelFontFace, catAxisLabelFontFace,
	// legendFontFace, ...), so escaping happens here — one site covers all of them. voidEl()'s
	// attrs escape by construction: an unescaped `"` or `&` would otherwise close the attribute
	// early and emit a non-parseable chart part.
	return voidEl('a:latin', { typeface }) + voidEl('a:ea', { typeface }) + voidEl('a:cs', { typeface })
}

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

			// Add series axis for 3D bar
			if (rel.opts._type === ChartType.bar3d) {
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
 * Stable PptxGenJS vendor GUID identifying the chart-metadata extension on `c:chartSpace/c:extLst`.
 * Custom data rides under this URI in a foreign namespace so PowerPoint preserves it (the extLst
 * mechanism) instead of stripping/repairing it as it would an unrecognised sibling element.
 */
const CHART_METADATA_EXT_URI = '{094A432E-1F6C-499B-95B8-B57DC9536949}'
/** Foreign namespace for the chart-metadata extension payload. */
const CHART_METADATA_NS = 'http://pptxgenjs.com/schema/chart/metadata'

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
		warn('chart `metadata` must be a plain object of string key/value pairs; ignored.')
		return ''
	}

	let items = ''
	for (const [key, value] of Object.entries(metadata)) {
		if (typeof key !== 'string' || key.length === 0) {
			warn(`chart metadata key "${String(key)}" is not a non-empty string; entry skipped.`)
			continue
		}
		if (typeof value !== 'string') {
			warn(`chart metadata value for key "${key}" is not a string; entry skipped.`)
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
// ===== Chart-type plotting =====

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
		case ChartType.doughnut:
		case ChartType.pie:
			return makePiePlot(chartType, data, opts, valFmtCode)
		default:
			return ''
	}
}
/**
 * Plot a category-axis chart family (area / bar / bar3d / line / radar) into a
 * `<c:xxxChart>` element. These share the grouping / series / cat+val axis structure.
 */
function makeCatAxisPlot(
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
		// Map the public PowerPoint-UI names to ST_RadarStyle wire values (also accepts the
		// deprecated wire spellings directly, in case an un-normalized value reaches here).
		const radarStyleWire =
			{ radar: 'standard', markers: 'marker', filled: 'filled', standard: 'standard', marker: 'marker' }[
				opts.radarStyle || 'radar'
			] ?? 'standard'
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
				strXml += `<a:prstDash val="${opts.lineDashValues?.[colorIndex] ?? opts.lineDash ?? 'solid'}"/><a:round/></a:ln>`
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
			if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
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

/**
 * Plot an XY scatter chart into `<c:scatterChart>` (paired X/Y numeric series).
 */
function makeScatterPlot(
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
			const chartColors = opts.chartColors?.length ? opts.chartColors : BARCHART_COLORS
			strXml += '<c:ser>'
			strXml += `  <c:idx val="${idx}"/>`
			strXml += `  <c:order val="${idx}"/>`
			strXml += '  <c:tx>'
			strXml += '    <c:strRef>'
			strXml += `      <c:f>${sheetCellRef(idx + 2, 1)}</c:f>`
			strXml +=
				'      <c:strCache><c:ptCount val="1"/><c:pt idx="0">' +
				el('c:v', null, obj.name ?? '') +
				'</c:pt></c:strCache>'
			strXml += '    </c:strRef>'
			strXml += '  </c:tx>'

			// 'c:spPr': Fill, Border, Line, LineStyle (dash, etc.), Shadow
			strXml += '  <c:spPr>'
			{
				const tmpSerColor = chartColors[colorIndex % chartColors.length] ?? '000000'

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
					strXml += `<a:ln w="${valToPts(opts.lineSize ?? 2)}" cap="${createLineCap(opts.lineCap)}">${chartColorLineFill(tmpSerColor)}`
					strXml += `<a:prstDash val="${opts.lineDashValues?.[colorIndex] ?? opts.lineDash ?? 'solid'}"/><a:round/></a:ln>`
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
					const markerColor =
						chartColors[idx + 1 > chartColors.length ? Math.floor(Math.random() * chartColors.length) : idx] ?? '000000'
					strXml += markerColor === 'transparent' ? '<a:noFill/>' : genXmlColorSelection(markerColor)
				}
				strXml += `<a:ln w="${opts.lineDataSymbolLineSize}" cap="flat">${chartColorLineFill(opts.lineDataSymbolLineColor || (chartColors[colorIndex % chartColors.length] ?? '000000'))}<a:prstDash val="solid"/><a:round/></a:ln>`
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
							strXml += `                    <a:rPr lang="${opts.lang || 'en-US'}" sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" b="${opts.dataLabelFontBold ? '1' : '0'}" i="${opts.dataLabelFontItalic ? '1' : '0'}" u="none" strike="noStrike" dirty="0">`
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
					strXml += `            <a:endParaRPr lang="${opts.lang || 'en-US'}"/>`
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
				strXml += `    <c:f>${sheetRangeRef(idx + 2, 2, idx + 2, dataValues(data[0]).length + 1)}</c:f>`
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

			// Option: `smooth`
			strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>'

			// 4: Close "SERIES"
			strXml += '</c:ser>'
		})

	// 3: Data Labels
	{
		strXml += '  <c:dLbls>'
		strXml +=
			'    ' + voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })
		strXml += '    <c:txPr>'
		strXml += '      <a:bodyPr/>'
		strXml += '      <a:lstStyle/>'
		strXml += '      <a:p><a:pPr>'
		strXml += `        <a:defRPr b="${opts.dataLabelFontBold ? '1' : '0'}" i="${opts.dataLabelFontItalic ? '1' : '0'}" strike="noStrike" sz="${ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)}" u="none">`
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
		strXml += '  </c:dLbls>'
	}

	// 4: Add axis Id (NOTE: order matters! - category comes first)
	strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/>`

	// 5: Close Chart tag
	strXml += '</c:' + chartType + 'Chart>'
	return strXml
}

/**
 * Plot a bubble / bubble3d chart into `<c:bubbleChart>` (X/Y plus per-point size).
 */
function makeBubblePlot(
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
					strXml += `<a:prstDash val="${opts.lineDashValues?.[colorIndex] ?? opts.lineDash ?? 'solid'}"/><a:round/></a:ln>`
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
		if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
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

/**
 * Plot a single-series pie / doughnut chart into `<c:pieChart>` / `<c:doughnutChart>`.
 */
function makePiePlot(
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
		const chartColors = opts.chartColors?.length ? opts.chartColors : BARCHART_COLORS
		const ptStyle = optsChartData.pointStyles?.[idx]
		strXml += '<c:dPt>'
		strXml += ` <c:idx val="${idx}"/>`
		strXml += ' <c:bubble3D val="0"/>'
		strXml += ' <c:spPr>'
		strXml += `<a:solidFill>${createColorElement(
			ptStyle?.fill ||
				(chartColors[idx + 1 > chartColors.length ? Math.floor(Math.random() * chartColors.length) : idx] ?? '000000')
		)}</a:solidFill>`
		// Per-point border override takes precedence over chart-level `dataBorder`
		if (ptStyle?.border) {
			strXml += createChartBorderLine(ptStyle.border)
		} else if (opts.dataBorder) {
			strXml += `<a:ln w="${valToPts(resolveBorderWidth(opts.dataBorder, 0.75))}" cap="flat">${genXmlColorSelection({
				color: opts.dataBorder.color ?? '363636',
				transparency: opts.dataBorder.transparency,
			})}<a:prstDash val="solid"/><a:round/></a:ln>`
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
				`<a:rPr lang="${opts.lang || 'en-US'}" dirty="0"/>` +
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
		if (chartType === ChartType.pie && opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
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
	strXml += chartType === ChartType.pie ? `<c:dLblPos val="${opts.dataLabelPosition || 'ctr'}"/>` : ''
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
// ===== Axes =====

function makeCatAxis(opts: ChartOptsInternal, axisId: string, valAxisId: string): string {
	let strXml = ''
	const usesValueAxisForCategories =
		opts._type === ChartType.scatter || opts._type === ChartType.bubble || opts._type === ChartType.bubble3d
	const usesCategoryAxis = !usesValueAxisForCategories && !opts.catLabelFormatCode

	// Build cat axis tag
	// NOTE: Scatter and Bubble chart need two Val axises as they display numbers on x axis
	if (usesValueAxisForCategories) {
		strXml += '<c:valAx>'
	} else {
		strXml += '<c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>'
	}
	strXml += '  <c:axId val="' + axisId + '"/>'
	strXml += '  <c:scaling>'
	strXml += '<c:orientation val="' + (opts.catAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) + '"/>'
	if (opts.catAxisMaxVal || opts.catAxisMaxVal === 0) strXml += `<c:max val="${opts.catAxisMaxVal}"/>`
	if (opts.catAxisMinVal || opts.catAxisMinVal === 0) strXml += `<c:min val="${opts.catAxisMinVal}"/>`
	strXml += '</c:scaling>'
	strXml += '  <c:delete val="' + (opts.catAxisHidden ? '1' : '0') + '"/>'
	strXml += '  <c:axPos val="' + (opts.barDir === 'col' ? 'b' : 'l') + '"/>'
	strXml += opts.catGridLine && opts.catGridLine.style !== 'none' ? createGridLineElement(opts.catGridLine) : ''
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showCatAxisTitle) {
		strXml += genXmlTitle({
			color: opts.catAxisTitleColor,
			fontFace: opts.catAxisTitleFontFace,
			fontSize: opts.catAxisTitleFontSize,
			titleRotate: opts.catAxisTitleRotate,
			title: opts.catAxisTitle || 'Axis Title',
		})
	}
	// NOTE: Adding Val Axis Formatting if scatter or bubble charts
	if (opts._type === ChartType.scatter || opts._type === ChartType.bubble || opts._type === ChartType.bubble3d) {
		const xAxisFmtCode = opts.catAxisLabelFormatCode ?? opts.valAxisLabelFormatCode
		strXml +=
			'  <c:numFmt formatCode="' + (xAxisFmtCode ? encodeXmlEntities(xAxisFmtCode) : 'General') + '" sourceLinked="1"/>'
	} else {
		strXml +=
			'  <c:numFmt formatCode="' +
			(encodeXmlEntities(opts.catLabelFormatCode ?? '') || 'General') +
			'" sourceLinked="1"/>'
	}
	if (opts._type === ChartType.scatter) {
		strXml += '  <c:majorTickMark val="none"/>'
		strXml += '  <c:minorTickMark val="none"/>'
		strXml += '  <c:tickLblPos val="' + (opts.catAxisLabelPos || 'nextTo') + '"/>'
	} else {
		strXml += '  <c:majorTickMark val="' + (opts.catAxisMajorTickMark || 'out') + '"/>'
		strXml += '  <c:minorTickMark val="' + (opts.catAxisMinorTickMark || 'none') + '"/>'
		strXml += '  <c:tickLblPos val="' + (opts.catAxisLabelPos || (opts.barDir === 'col' ? 'low' : 'nextTo')) + '"/>'
	}
	strXml += '  <c:spPr>'
	strXml += `    <a:ln w="${opts.catAxisLineSize ? valToPts(opts.catAxisLineSize) : ONEPT}" cap="flat">`
	strXml += !opts.catAxisLineShow ? '<a:noFill/>' : genXmlColorSelection(opts.catAxisLineColor || DEF_GRIDLINE_COLOR)
	strXml += '      <a:prstDash val="' + (opts.catAxisLineStyle || 'solid') + '"/>'
	strXml += '      <a:round/>'
	strXml += '    </a:ln>'
	strXml += '  </c:spPr>'
	strXml += '  <c:txPr>'
	if (opts.catAxisLabelRotate) {
		strXml += `<a:bodyPr rot="${convertRotationDegrees(opts.catAxisLabelRotate)}"/>`
	} else {
		// NOTE: don't specify "`rot=0" - that way the object will be auto behavior
		strXml += '<a:bodyPr/>'
	}
	strXml += '    <a:lstStyle/>'
	strXml += '    <a:p>'
	strXml += '    <a:pPr>'
	strXml += `      <a:defRPr sz="${ptToHundredths(opts.catAxisLabelFontSize || DEF_FONT_SIZE)}" b="${opts.catAxisLabelFontBold ? 1 : 0}" i="${opts.catAxisLabelFontItalic ? 1 : 0}" u="none" strike="noStrike">`
	strXml += genXmlColorSelection(opts.catAxisLabelColor || DEF_FONT_COLOR)
	strXml += '      ' + createChartTextFonts(opts.catAxisLabelFontFace || 'Arial')
	strXml += '   </a:defRPr>'
	strXml += '  </a:pPr>'
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>'
	strXml += '  </a:p>'
	strXml += ' </c:txPr>'
	strXml += ' <c:crossAx val="' + valAxisId + '"/>'
	const valAxisCrossTag = typeof opts.valAxisCrossesAt === 'number' ? 'crossesAt' : 'crosses'
	const valAxisCrossValue =
		typeof opts.valAxisCrossesAt === 'number' ? opts.valAxisCrossesAt : opts.valAxisCrossesAt || 'autoZero'
	strXml += ` <c:${valAxisCrossTag} val="${valAxisCrossValue}"/>`
	if (!usesValueAxisForCategories) strXml += ' <c:auto val="1"/>'
	if (usesCategoryAxis) {
		strXml += ' <c:lblAlgn val="ctr"/>'
		if (opts.catAxisLabelFrequency) strXml += ' <c:tickLblSkip val="' + opts.catAxisLabelFrequency + '"/>'
		strXml += ` <c:noMultiLvlLbl val="${opts.catAxisMultiLevelLabels ? 0 : 1}"/>`
	}

	// PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
	// Allow major and minor units to be set for double value axis charts
	if (opts.catLabelFormatCode || usesValueAxisForCategories) {
		if (opts.catLabelFormatCode) {
			;(['catAxisBaseTimeUnit', 'catAxisMajorTimeUnit', 'catAxisMinorTimeUnit'] as const).forEach((opt) => {
				// Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
				const optVal = opts[opt]
				if (optVal && (typeof optVal !== 'string' || !VALID_CHART_TIME_UNITS.includes(optVal.toLowerCase()))) {
					warn(`"${opt}" must be one of: 'days','months','years' !`)
					opts[opt] = undefined
				}
			})
			if (opts.catAxisBaseTimeUnit) strXml += '<c:baseTimeUnit val="' + opts.catAxisBaseTimeUnit.toLowerCase() + '"/>'
			if (opts.catAxisMajorTimeUnit)
				strXml += '<c:majorTimeUnit val="' + opts.catAxisMajorTimeUnit.toLowerCase() + '"/>'
			if (opts.catAxisMinorTimeUnit)
				strXml += '<c:minorTimeUnit val="' + opts.catAxisMinorTimeUnit.toLowerCase() + '"/>'
		}
		if (opts.catAxisMajorUnit) strXml += `<c:majorUnit val="${opts.catAxisMajorUnit}"/>`
		if (opts.catAxisMinorUnit) strXml += `<c:minorUnit val="${opts.catAxisMinorUnit}"/>`
	}

	// Close cat axis tag
	// NOTE: Added closing tag of val or cat axis based on chart type
	if (usesValueAxisForCategories) {
		strXml += '</c:valAx>'
	} else {
		strXml += '</c:' + (opts.catLabelFormatCode ? 'dateAx' : 'catAx') + '>'
	}

	return strXml
}

/**
 * Create Value Axis (Used by `bar3D`)
 * @param {ChartOptsInternal} opts - chart options
 * @param {string} valAxisId - value
 * @return {string} XML
 */
function makeValAxis(opts: ChartOptsInternal, valAxisId: string): string {
	let axisPos =
		valAxisId === AXIS_ID_VALUE_PRIMARY ? (opts.barDir === 'col' ? 'l' : 'b') : opts.barDir !== 'col' ? 'r' : 't'
	if (valAxisId === AXIS_ID_VALUE_SECONDARY) axisPos = 'r' // default behavior for PPT is showing 2nd val axis on right (primary axis on left)
	const crossAxId = valAxisId === AXIS_ID_VALUE_PRIMARY ? AXIS_ID_CATEGORY_PRIMARY : AXIS_ID_CATEGORY_SECONDARY
	let strXml = ''

	strXml += '<c:valAx>'
	strXml += '  <c:axId val="' + valAxisId + '"/>'
	strXml += '  <c:scaling>'
	if (opts.valAxisLogScaleBase) strXml += `<c:logBase val="${opts.valAxisLogScaleBase}"/>`
	strXml += '<c:orientation val="' + (opts.valAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) + '"/>'
	if (opts.valAxisMaxVal || opts.valAxisMaxVal === 0) strXml += `<c:max val="${opts.valAxisMaxVal}"/>`
	if (opts.valAxisMinVal || opts.valAxisMinVal === 0) strXml += `<c:min val="${opts.valAxisMinVal}"/>`
	strXml += '  </c:scaling>'
	strXml += `  <c:delete val="${opts.valAxisHidden ? 1 : 0}"/>`
	strXml += '  <c:axPos val="' + axisPos + '"/>'
	if (opts.valGridLine && opts.valGridLine.style !== 'none') strXml += createGridLineElement(opts.valGridLine)
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showValAxisTitle) {
		strXml += genXmlTitle({
			color: opts.valAxisTitleColor,
			fontFace: opts.valAxisTitleFontFace,
			fontSize: opts.valAxisTitleFontSize,
			titleRotate: opts.valAxisTitleRotate,
			title: opts.valAxisTitle || 'Axis Title',
		})
	}
	strXml += `<c:numFmt formatCode="${opts.valAxisLabelFormatCode ? encodeXmlEntities(opts.valAxisLabelFormatCode) : 'General'}" sourceLinked="0"/>`
	if (opts._type === ChartType.scatter) {
		strXml += '  <c:majorTickMark val="none"/>'
		strXml += '  <c:minorTickMark val="none"/>'
		strXml += '  <c:tickLblPos val="nextTo"/>'
	} else {
		strXml += ' <c:majorTickMark val="' + (opts.valAxisMajorTickMark || 'out') + '"/>'
		strXml += ' <c:minorTickMark val="' + (opts.valAxisMinorTickMark || 'none') + '"/>'
		strXml += ' <c:tickLblPos val="' + (opts.valAxisLabelPos || (opts.barDir === 'col' ? 'nextTo' : 'low')) + '"/>'
	}
	strXml += ' <c:spPr>'
	strXml += `   <a:ln w="${opts.valAxisLineSize ? valToPts(opts.valAxisLineSize) : ONEPT}" cap="flat">`
	strXml += !opts.valAxisLineShow ? '<a:noFill/>' : genXmlColorSelection(opts.valAxisLineColor || DEF_GRIDLINE_COLOR)
	strXml += '     <a:prstDash val="' + (opts.valAxisLineStyle || 'solid') + '"/>'
	strXml += '     <a:round/>'
	strXml += '   </a:ln>'
	strXml += ' </c:spPr>'
	strXml += ' <c:txPr>'
	strXml += `  <a:bodyPr${opts.valAxisLabelRotate ? ' rot="' + convertRotationDegrees(opts.valAxisLabelRotate).toString() + '"' : ''}/>` // don't specify rot 0 so we get the auto behavior
	strXml += '  <a:lstStyle/>'
	strXml += '  <a:p>'
	strXml += '    <a:pPr>'
	strXml += `      <a:defRPr sz="${ptToHundredths(opts.valAxisLabelFontSize || DEF_FONT_SIZE)}" b="${opts.valAxisLabelFontBold ? 1 : 0}" i="${opts.valAxisLabelFontItalic ? 1 : 0}" u="none" strike="noStrike">`
	strXml += genXmlColorSelection(opts.valAxisLabelColor || DEF_FONT_COLOR)
	strXml += '        ' + createChartTextFonts(opts.valAxisLabelFontFace || 'Arial')
	strXml += '      </a:defRPr>'
	strXml += '    </a:pPr>'
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>'
	strXml += '  </a:p>'
	strXml += ' </c:txPr>'
	strXml += ' <c:crossAx val="' + crossAxId + '"/>'
	if (typeof opts.catAxisCrossesAt === 'number') {
		strXml += ` <c:crossesAt val="${opts.catAxisCrossesAt}"/>`
	} else if (typeof opts.catAxisCrossesAt === 'string') {
		strXml += ' <c:crosses val="' + opts.catAxisCrossesAt + '"/>'
	} else {
		const isRight = axisPos === 'r' || axisPos === 't'
		const crosses = isRight ? 'max' : 'autoZero'
		strXml += ' <c:crosses val="' + crosses + '"/>'
	}
	strXml +=
		' <c:crossBetween val="' +
		(opts.valAxisCrossBetween
			? opts.valAxisCrossBetween
			: opts._type === ChartType.scatter ||
				  !!(Array.isArray(opts._type) && opts._type.some((type) => asChartType(type.type) === ChartType.area))
				? 'midCat'
				: 'between') +
		'"/>'
	if (opts.valAxisMajorUnit) strXml += ` <c:majorUnit val="${opts.valAxisMajorUnit}"/>`
	if (opts.valAxisDisplayUnit) {
		strXml += `<c:dispUnits><c:builtInUnit val="${opts.valAxisDisplayUnit}"/>${opts.valAxisDisplayUnitLabel ? '<c:dispUnitsLbl/>' : ''}</c:dispUnits>`
	}

	strXml += '</c:valAx>'

	return strXml
}

/**
 * Create Series Axis (Used by `bar3D`)
 * @param {ChartOptsInternal} opts - chart options
 * @param {string} axisId - axis ID
 * @param {string} valAxisId - value
 * @return {string} XML
 */
function makeSerAxis(opts: ChartOptsInternal, axisId: string, valAxisId: string): string {
	let strXml = ''

	// Build ser axis tag
	strXml += '<c:serAx>'
	strXml += '  <c:axId val="' + axisId + '"/>'
	strXml +=
		'  <c:scaling><c:orientation val="' +
		(opts.serAxisOrientation || (opts.barDir === 'col' ? 'minMax' : 'minMax')) +
		'"/></c:scaling>'
	strXml += '  <c:delete val="' + (opts.serAxisHidden ? '1' : '0') + '"/>'
	strXml += '  <c:axPos val="' + (opts.barDir === 'col' ? 'b' : 'l') + '"/>'
	strXml += opts.serGridLine && opts.serGridLine.style !== 'none' ? createGridLineElement(opts.serGridLine) : ''
	// '<c:title>' comes between '</c:majorGridlines>' and '<c:numFmt>'
	if (opts.showSerAxisTitle) {
		strXml += genXmlTitle({
			color: opts.serAxisTitleColor,
			fontFace: opts.serAxisTitleFontFace,
			fontSize: opts.serAxisTitleFontSize,
			titleRotate: opts.serAxisTitleRotate,
			title: opts.serAxisTitle || 'Axis Title',
		})
	}
	strXml += `  <c:numFmt formatCode="${encodeXmlEntities(opts.serLabelFormatCode ?? '') || 'General'}" sourceLinked="0"/>`
	strXml += '  <c:majorTickMark val="out"/>'
	strXml += '  <c:minorTickMark val="none"/>'
	strXml += `  <c:tickLblPos val="${opts.serAxisLabelPos || opts.barDir === 'col' ? 'low' : 'nextTo'}"/>`
	strXml += '  <c:spPr>'
	strXml += '    <a:ln w="12700" cap="flat">'
	strXml += !opts.serAxisLineShow ? '<a:noFill/>' : genXmlColorSelection(opts.serAxisLineColor || DEF_GRIDLINE_COLOR)
	strXml += '      <a:prstDash val="solid"/>'
	strXml += '      <a:round/>'
	strXml += '    </a:ln>'
	strXml += '  </c:spPr>'
	strXml += '  <c:txPr>'
	strXml += '    <a:bodyPr/>' // don't specify rot 0 so we get the auto behavior
	strXml += '    <a:lstStyle/>'
	strXml += '    <a:p>'
	strXml += '    <a:pPr>'
	strXml += `    <a:defRPr sz="${ptToHundredths(opts.serAxisLabelFontSize || DEF_FONT_SIZE)}" b="${opts.serAxisLabelFontBold ? '1' : '0'}" i="${opts.serAxisLabelFontItalic ? '1' : '0'}" u="none" strike="noStrike">`
	strXml += `      ${genXmlColorSelection(opts.serAxisLabelColor || DEF_FONT_COLOR)}`
	strXml += '      ' + createChartTextFonts(opts.serAxisLabelFontFace || 'Arial')
	strXml += '   </a:defRPr>'
	strXml += '  </a:pPr>'
	strXml += '  <a:endParaRPr lang="' + (opts.lang || 'en-US') + '"/>'
	strXml += '  </a:p>'
	strXml += ' </c:txPr>'
	strXml += ' <c:crossAx val="' + valAxisId + '"/>'
	strXml += ' <c:crosses val="autoZero"/>'
	if (opts.serAxisLabelFrequency) strXml += ' <c:tickLblSkip val="' + opts.serAxisLabelFrequency + '"/>'

	// PPT will auto-adjust these as needed after calcing the date bounds, so we only include them when specified by user
	if (opts.serLabelFormatCode) {
		;(['serAxisBaseTimeUnit', 'serAxisMajorTimeUnit', 'serAxisMinorTimeUnit'] as const).forEach((opt) => {
			// Validate input as poorly chosen/garbage options will cause chart corruption and it wont render at all!
			const optVal = opts[opt]
			if (optVal && (typeof optVal !== 'string' || !VALID_CHART_TIME_UNITS.includes(optVal.toLowerCase()))) {
				warn(`"${opt}" must be one of: 'days','months','years' !`)
				opts[opt] = undefined
			}
		})
		if (opts.serAxisBaseTimeUnit) strXml += ` <c:baseTimeUnit  val="${opts.serAxisBaseTimeUnit.toLowerCase()}"/>`
		if (opts.serAxisMajorTimeUnit) strXml += ` <c:majorTimeUnit val="${opts.serAxisMajorTimeUnit.toLowerCase()}"/>`
		if (opts.serAxisMinorTimeUnit) strXml += ` <c:minorTimeUnit val="${opts.serAxisMinorTimeUnit.toLowerCase()}"/>`
		if (opts.serAxisMajorUnit) strXml += ` <c:majorUnit val="${opts.serAxisMajorUnit}"/>`
		if (opts.serAxisMinorUnit) strXml += ` <c:minorUnit val="${opts.serAxisMinorUnit}"/>`
	}

	// Close ser axis tag
	strXml += '</c:serAx>'

	return strXml
}

/**
 * Create char title elements
 * @param {ChartPropsTitle} opts - options
 * @return {string} XML `<c:title>`
 */
// ===== Titles & shared builders =====

function genXmlTitle(opts: ChartPropsTitle, chartX?: number, chartY?: number): string {
	const align =
		opts.titleAlign === 'left' || opts.titleAlign === 'right'
			? `<a:pPr algn="${opts.titleAlign.slice(0, 1)}">`
			: '<a:pPr>'
	const rotate = opts.titleRotate ? `<a:bodyPr rot="${convertRotationDegrees(opts.titleRotate)}"/>` : '<a:bodyPr/>' // don't specify rotation to get default (ex. vertical for cat axis)
	const sizeAttr = opts.fontSize ? `sz="${ptToHundredths(opts.fontSize)}"` : '' // only set the font size if specified.  Powerpoint will handle the default size
	const titleBold = opts.titleBold ? 1 : 0
	const titleItalic = opts.titleItalic ? 1 : 0
	const titleUnderline = opts.titleUnderline ? 'sng' : 'none'

	let layout = '<c:layout/>'
	const hasX = opts.titlePos && typeof opts.titlePos.x === 'number'
	const hasY = opts.titlePos && typeof opts.titlePos.y === 'number'
	if (hasX || hasY) {
		// NOTE: manualLayout x/y vals are *relative to entire slide*. Each axis is
		// independent in CT_ManualLayout: omitting xMode/x (or yMode/y) leaves that
		// axis on automatic layout, so a caller can center horizontally while still
		// applying a manual vertical offset (and vice-versa).
		// Schema order is xMode, yMode, x, y.
		let modes = ''
		let vals = ''
		if (hasX) {
			const totalX = (opts.titlePos?.x ?? 0) + (chartX ?? 0)
			let valX = totalX === 0 ? 0 : (totalX * (totalX / 5)) / 10
			if (valX >= 1) valX = valX / 10
			if (valX >= 0.1) valX = valX / 10
			modes += '<c:xMode val="edge"/>'
			vals += `<c:x val="${valX}"/>`
		}
		if (hasY) {
			const totalY = (opts.titlePos?.y ?? 0) + (chartY ?? 0)
			let valY = totalY === 0 ? 0 : (totalY * (totalY / 5)) / 10
			if (valY >= 1) valY = valY / 10
			if (valY >= 0.1) valY = valY / 10
			modes += '<c:yMode val="edge"/>'
			vals += `<c:y val="${valY}"/>`
		}
		layout = `<c:layout><c:manualLayout>${modes}${vals}</c:manualLayout></c:layout>`
	}

	return `<c:title>
      <c:tx>
        <c:rich>
          ${rotate}
          <a:lstStyle/>
          <a:p>
            ${align}
            <a:defRPr ${sizeAttr} b="${titleBold}" i="${titleItalic}" u="${titleUnderline}" strike="noStrike">
              ${genXmlColorSelection(opts.color || DEF_FONT_COLOR)}
              ${createChartTextFonts(opts.fontFace || 'Arial')}
            </a:defRPr>
          </a:pPr>
          <a:r>
            <a:rPr ${sizeAttr} b="${titleBold}" i="${titleItalic}" u="${titleUnderline}" strike="noStrike">
              ${genXmlColorSelection(opts.color || DEF_FONT_COLOR)}
              ${createChartTextFonts(opts.fontFace || 'Arial')}
            </a:rPr>
            <a:t>${encodeXmlEntities(opts.title ?? '') || ''}</a:t>
          </a:r>
        </a:p>
        </c:rich>
      </c:tx>
      ${layout}
      <c:overlay val="0"/>
    </c:title>`
}

/**
 * Create Grid Line Element
 * @param {OptsChartGridLine} glOpts {size, color, style}
 * @return {string} XML
 */
function createGridLineElement(glOpts: OptsChartGridLine): string {
	let strXml = '<c:majorGridlines>'
	strXml += ' <c:spPr>'
	strXml += `  <a:ln w="${valToPts(glOpts.size || DEF_CHART_GRIDLINE.size || 1)}" cap="${createLineCap(glOpts.cap || DEF_CHART_GRIDLINE.cap)}">`
	strXml += '  <a:solidFill><a:srgbClr val="' + (glOpts.color || DEF_CHART_GRIDLINE.color) + '"/></a:solidFill>' // should accept scheme colors
	strXml += '   <a:prstDash val="' + (glOpts.style || DEF_CHART_GRIDLINE.style) + '"/><a:round/>'
	strXml += '  </a:ln>'
	strXml += ' </c:spPr>'
	strXml += '</c:majorGridlines>'

	return strXml
}

/**
 * Build a `<c:pt>` numeric-cache data point, or '' to leave a gap.
 *
 * `<c:v>` inside a `<c:numCache>` is an `xsd:double`; emitting `NaN`, `Infinity`
 * or an empty string yields an invalid value that makes PowerPoint report the
 * package as needing repair. Null/undefined are intentional gaps and are skipped
 * silently (a sparse, idx-keyed cache is valid); other non-finite numbers are
 * skipped with a warning, per the library's "warn rather than emit a degenerate
 * result" policy.
 * @param idx - zero-based data-point index (emitted as `idx`)
 * @param value - numeric value (or null/undefined gap)
 */
function numCachePt(idx: number, value: number | null | undefined): string {
	if (value == null) return ''
	if (!Number.isFinite(value)) {
		warn(`chart value "${value}" at index ${idx} is not a finite number; data point omitted.`)
		return ''
	}
	return `<c:pt idx="${idx}"><c:v>${value}</c:v></c:pt>`
}

/**
 * Build the error-bar elements (`<c:errBars>`) for a single series.
 *
 * Schema position (CT_*Ser): after `dLbls`/`trendline`, before `cat`/`val` (bar/line/area)
 * or `xVal`/`yVal` (scatter). CT_ErrBars child order is errDir → errBarType → errValType →
 * noEndCap → plus → minus → val → spPr.
 *
 * @param chartType - chart this series belongs to (used to bound how many bars are legal)
 * @param errorBars - one config, or an array (X+Y) for scatter/area; bar/line keep only the first
 * @param obj - the series data object (only `name`, for warnings)
 */
function makeChartErrorBarsXml(
	chartType: ChartType,
	errorBars: ChartErrorBarOptions | ChartErrorBarOptions[] | undefined,
	obj: OptsChartDataInternal
): string {
	if (!errorBars) return ''
	const bars = Array.isArray(errorBars) ? errorBars : [errorBars]
	// CT_BarSer/CT_LineSer allow a single <c:errBars>; only scatter/area permit two (x + y).
	const maxBars = chartType === ChartType.scatter || chartType === ChartType.area ? 2 : 1
	let strXml = ''

	bars.slice(0, maxBars).forEach((eb) => {
		if (!eb) return
		const valueType = eb.valueType || 'fixedVal'
		const barType = eb.barType || 'both'
		const direction = eb.direction || 'y'

		strXml += '<c:errBars>'
		strXml += `<c:errDir val="${direction}"/>`
		strXml += `<c:errBarType val="${barType}"/>`
		strXml += `<c:errValType val="${valueType}"/>`
		strXml += `<c:noEndCap val="${eb.noEndCap ? '1' : '0'}"/>`

		if (valueType === 'cust') {
			// Custom amounts: <c:plus>/<c:minus> each hold a number source (we emit <c:numLit>).
			// `barType` decides which sides are present; warn (don't silently drop) on a missing side.
			if (barType !== 'minus') {
				if (!eb.plusValues?.length)
					warn(`chart series "${obj.name}" errorBars valueType 'cust' needs \`plusValues\` for barType '${barType}'.`)
				strXml += makeErrBarNumLit('plus', eb.plusValues || [])
			}
			if (barType !== 'plus') {
				if (!eb.minusValues?.length)
					warn(`chart series "${obj.name}" errorBars valueType 'cust' needs \`minusValues\` for barType '${barType}'.`)
				strXml += makeErrBarNumLit('minus', eb.minusValues || [])
			}
		} else if (valueType !== 'stdErr') {
			// fixedVal / percentage / stdDev use a single magnitude (stdErr derives it from the data).
			strXml += `<c:val val="${eb.value ?? 1}"/>`
		}

		if (eb.color || eb.size != null) {
			strXml += '<c:spPr><a:ln'
			strXml += eb.size != null ? ` w="${valToPts(eb.size)}"` : ''
			strXml += '>'
			strXml += eb.color ? genXmlColorSelection(eb.color) : ''
			strXml += '</a:ln></c:spPr>'
		}

		strXml += '</c:errBars>'
	})

	return strXml
}

/**
 * Build a `<c:plus>`/`<c:minus>` number-literal source for custom error-bar amounts.
 * @param tag - `'plus'` or `'minus'`
 * @param values - per-point magnitudes (index-aligned with the series values)
 */
function makeErrBarNumLit(tag: 'plus' | 'minus', values: number[]): string {
	let strXml = `<c:${tag}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>`
	values.forEach((value, idx) => {
		strXml += numCachePt(idx, value)
	})
	strXml += `</c:numLit></c:${tag}>`

	return strXml
}

/**
 * Build a `<c:serLines>` ("Series Lines") element for a bar chart.
 * @param opt - `true` for PowerPoint automatic styling, an {@link OptsChartGridLine}
 *   to customize the line, or falsy / `{ style: 'none' }` to omit the element.
 */
function createSerLinesElement(opt?: boolean | OptsChartGridLine): string {
	if (!opt) return ''
	if (opt === true) return '<c:serLines/>'
	if (opt.style === 'none') return ''
	let strXml = '<c:serLines><c:spPr>'
	strXml += `<a:ln w="${valToPts(opt.size || DEF_CHART_GRIDLINE.size || 1)}" cap="${createLineCap(opt.cap || DEF_CHART_GRIDLINE.cap)}">`
	strXml += `<a:solidFill><a:srgbClr val="${opt.color || DEF_CHART_GRIDLINE.color}"/></a:solidFill>`
	strXml += `<a:prstDash val="${opt.style || DEF_CHART_GRIDLINE.style}"/><a:round/>`
	strXml += '</a:ln></c:spPr></c:serLines>'

	return strXml
}

/**
 * Build the `<c:leaderLines>` element for pie/doughnut data labels.
 *
 * Schema position: inside `<c:dLbls>`, immediately after `<c:showLeaderLines>`
 * (CT_DLbls / Group_DLbls order: showLeaderLines → leaderLines).
 *
 * Returns `''` unless the caller both enabled leader lines (`showLeaderLines`)
 * and configured their appearance (`leaderLineColor` / `leaderLineSize`). When
 * appearance is unset we leave the element off so PowerPoint applies its
 * automatic leader-line color, matching prior behavior.
 *
 * @param opts - chart options (reads `showLeaderLines`, `leaderLineColor`, `leaderLineSize`)
 */
function createLeaderLinesElement(opts: ChartOptsInternal): string {
	if (!opts.showLeaderLines) return ''
	if (!opts.leaderLineColor && opts.leaderLineSize == null) return ''
	const w = valToPts(opts.leaderLineSize ?? 0.75)
	const color = opts.leaderLineColor || '808080'
	return (
		'<c:leaderLines><c:spPr>' +
		`<a:ln w="${w}" cap="flat">${genXmlColorSelection(color)}<a:prstDash val="solid"/><a:round/></a:ln>` +
		'<a:effectLst/></c:spPr></c:leaderLines>'
	)
}

/**
 * Build a single custom `<c:dLbl>` (rich-text data label) overriding one data point's label.
 *
 * Used when a series supplies explicit per-point label text: the emitted `<c:idx>` pins the
 * override to that point, and the `<c:rich>` run carries the label's own font styling (size,
 * bold, italic, color, face) resolved from the chart-level dataLabel* options. All the
 * `show*` flags are forced off so only the literal `text` renders (no value/category/percent).
 * @param idx - zero-based data-point index this label overrides
 * @param text - the literal label text (XML-escaped here)
 * @param opts - chart options supplying dataLabel font/color defaults and `lang`
 * @return {string} a `<c:dLbl>` element
 */
function makeCustomDLblXml(idx: number, text: string, opts: ChartOptsInternal): string {
	const sz = ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE)
	const bold = opts.dataLabelFontBold ? '1' : '0'
	const italic = opts.dataLabelFontItalic ? '1' : '0'
	const color = createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)
	const face = opts.dataLabelFontFace || 'Arial'
	const lang = opts.lang || 'en-US'
	return (
		`<c:dLbl><c:idx val="${idx}"/>` +
		'<c:tx><c:rich><a:bodyPr/><a:lstStyle/>' +
		`<a:p><a:pPr><a:defRPr sz="${sz}" b="${bold}" i="${italic}" u="none" strike="noStrike">` +
		`<a:solidFill>${color}</a:solidFill>${createChartTextFonts(face)}</a:defRPr></a:pPr>` +
		`<a:r><a:rPr lang="${lang}" sz="${sz}" b="${bold}" i="${italic}" u="none" strike="noStrike" dirty="0">` +
		`<a:solidFill>${color}</a:solidFill>${createChartTextFonts(face)}</a:rPr>` +
		`<a:t>${encodeXmlEntities(text)}</a:t></a:r></a:p>` +
		'</c:rich></c:tx>' +
		'<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>' +
		'<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>'
	)
}

/**
 * Build an `<a:ln>` border element from a per-data-point `BorderProps`.
 * @param border - point border style (`type`, `color`, `pt`)
 */
function createChartBorderLine(border: BorderProps): string {
	if (border.type === 'none') return '<a:ln><a:noFill/></a:ln>'
	const dash = border.type === 'dash' ? 'dash' : 'solid'
	return `<a:ln w="${valToPts(resolveBorderWidth(border, 1))}" cap="flat">${genXmlColorSelection({ color: border.color || '666666', transparency: border.transparency })}<a:prstDash val="${dash}"/><a:round/></a:ln>`
}

/**
 * Build `<c:dPt>` entries for a series in the bar/line/area/scatter loops.
 *
 * Merges two sources into a single `c:dPt` per index so we never emit a
 * duplicate `<c:idx>` (which corrupts the chart):
 * - legacy single-series color-vary fills (bar/scatter), supplied via `varyColors`
 * - per-point `pointStyles` border/fill overrides
 *
 * Must be emitted in schema position *before* `c:dLbls` (CT_*Ser order).
 * RADAR is skipped: extra per-point markup historically corrupts the chart.
 *
 * @param chartType  - series chart type
 * @param obj        - series data (reads `values`, `pointStyles`)
 * @param opts       - chart options (fill/shadow/lineSize context)
 * @param varyColors - color array when single-series color-vary applies, else `null`
 */
function makeSeriesDataPointsXml(
	chartType: ChartType,
	obj: OptsChartDataInternal,
	opts: ChartOptsInternal,
	varyColors: string[] | null
): string {
	if (chartType === ChartType.radar) return ''
	const pointStyles = obj.pointStyles
	if (!varyColors && !pointStyles?.length) return ''

	const isBar = chartType === ChartType.bar || chartType === ChartType.bar3d
	const isScatter = chartType === ChartType.scatter
	let xml = ''
	dataValues(obj).forEach((value, index) => {
		const ptStyle = pointStyles?.[index]
		const arrColors = varyColors
			? value < 0
				? opts.invertedColors || opts.chartColors || BARCHART_COLORS
				: varyColors
			: null
		const fillColor = ptStyle?.fill || (arrColors ? arrColors[index % arrColors.length] : null)
		const pattern = ptStyle?.pattern
		const border = ptStyle?.border
		// Nothing to style for this point -> omit the c:dPt entirely
		if (!fillColor && !pattern && !border) return

		xml += '<c:dPt>'
		xml += `<c:idx val="${index}"/>`
		if (isBar) xml += '<c:invertIfNegative val="0"/>'
		xml += '<c:bubble3D val="0"/>'
		xml += '<c:spPr>'
		if ((isBar || isScatter) && opts.lineSize === 0 && !border && !ptStyle?.fill && !pattern) {
			// Preserve legacy color-vary behavior: hide outline when lineSize===0
			xml += '<a:ln><a:noFill/></a:ln>'
		} else {
			// Pattern fill takes precedence over a solid fill (OOXML allows only one fill per c:dPt).
			// Default the pattern foreground to this point's resolved color so it reads as a hatched bar.
			if (pattern) {
				xml += genXmlPatternFill(fillColor && !pattern.fgColor ? { ...pattern, fgColor: fillColor } : pattern)
			} else if (fillColor) {
				// BAR3D color-vary historically tints the edge line, not the face fill
				if (chartType === ChartType.bar3d) xml += `<a:ln>${genXmlColorSelection(fillColor)}</a:ln>`
				else xml += genXmlColorSelection(fillColor)
			}
			if (border) xml += createChartBorderLine(border)
		}
		xml += createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)
		xml += '</c:spPr>'
		xml += '</c:dPt>'
	})
	return xml
}
