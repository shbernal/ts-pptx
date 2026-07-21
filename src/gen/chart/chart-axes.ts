/**
 * PptxGenJS: Chart Axis Assembly
 *
 * Builds the `<c:catAx>` / `<c:valAx>` / `<c:serAx>` elements. The three share a
 * structure -- scaling, delete flag, tick marks, label position, text properties,
 * gridlines, optional axis title -- but PowerPoint requires the child elements in a
 * fixed order that differs per axis, so they stay as three separate builders rather
 * than one parameterized one. Called by {@link ./chart-xml}'s axis region.
 */

import {
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
	asChartType,
	ChartType,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
} from '../../core-enums.js'
import type { ChartOptsInternal } from '../../types/internal.js'
import { warn } from '../../log.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { convertRotationDegrees, valToPts } from '../../units-internal.js'
import { EMU_PER_POINT, ptToHundredths } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	createChartTextFonts,
	createGridLineElement,
	DEF_GRIDLINE_COLOR,
	genXmlTitle,
	VALID_CHART_TIME_UNITS,
} from './chart-parts.js'

export function makeCatAxis(opts: ChartOptsInternal, axisId: string, valAxisId: string): string {
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
		strXml += '  ' + voidEl('c:numFmt', { formatCode: xAxisFmtCode || 'General', sourceLinked: 1 })
	} else {
		strXml += '  ' + voidEl('c:numFmt', { formatCode: (opts.catLabelFormatCode ?? '') || 'General', sourceLinked: 1 })
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
	strXml += `    <a:ln w="${opts.catAxisLineSize ? valToPts(opts.catAxisLineSize) : EMU_PER_POINT}" cap="flat">`
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
				strXml += voidEl('c:majorTimeUnit', { val: opts.catAxisMajorTimeUnit.toLowerCase() })
			if (opts.catAxisMinorTimeUnit)
				strXml += voidEl('c:minorTimeUnit', { val: opts.catAxisMinorTimeUnit.toLowerCase() })
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
export function makeValAxis(opts: ChartOptsInternal, valAxisId: string): string {
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
	strXml += voidEl('c:numFmt', { formatCode: opts.valAxisLabelFormatCode || 'General', sourceLinked: 0 })
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
	strXml += `   <a:ln w="${opts.valAxisLineSize ? valToPts(opts.valAxisLineSize) : EMU_PER_POINT}" cap="flat">`
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
		strXml += el('c:dispUnits', null, [
			raw(voidEl('c:builtInUnit', { val: opts.valAxisDisplayUnit })),
			raw(opts.valAxisDisplayUnitLabel ? voidEl('c:dispUnitsLbl') : ''),
		])
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
export function makeSerAxis(opts: ChartOptsInternal, axisId: string, valAxisId: string): string {
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
	strXml += '  ' + voidEl('c:numFmt', { formatCode: (opts.serLabelFormatCode ?? '') || 'General', sourceLinked: 0 })
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
		// `baseTimeUnit` keeps its template string on purpose: it emits TWO spaces before
		// `val`, and voidEl() joins attributes with exactly one. Normalizing the spacing
		// would be a byte change, so the quirk stays visible here rather than being
		// silently "fixed" by the builder.
		if (opts.serAxisBaseTimeUnit) strXml += ` <c:baseTimeUnit  val="${opts.serAxisBaseTimeUnit.toLowerCase()}"/>`
		if (opts.serAxisMajorTimeUnit)
			strXml += voidEl('c:majorTimeUnit', { val: opts.serAxisMajorTimeUnit.toLowerCase() }, { openPrefix: ' ' })
		if (opts.serAxisMinorTimeUnit)
			strXml += voidEl('c:minorTimeUnit', { val: opts.serAxisMinorTimeUnit.toLowerCase() }, { openPrefix: ' ' })
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
