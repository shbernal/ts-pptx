/**
 * ts-pptx: Chart Axis Assembly
 *
 * Builds the `<c:catAx>` / `<c:valAx>` / `<c:serAx>` elements. The three share a
 * structure -- scaling, delete flag, tick marks, label position, text properties,
 * gridlines, optional axis title -- but PowerPoint requires the child elements in a
 * fixed order that differs per axis, so they stay as three separate builders rather
 * than one parameterized one. Called by {@link ./chart-xml}'s axis region.
 *
 * Sub-blocks are shared where they can be shared *exactly*: {@link axisLineSpPr} covers
 * `<c:spPr>`, which differs between the three only by indent width. `<c:txPr>` is
 * deliberately left duplicated — its inherited indentation is irregular *within* each copy
 * as well as between them (the category axis closes `</a:defRPr>` at three spaces, the value
 * axis at six), so parameterising it would take about as many indent arguments as it has
 * lines and would hide the quirks rather than share the structure.
 */

import { asChartType, ChartType } from '../../enums.js'
import {
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
} from '../../constants-internal.js'
import type { ChartOptsInternal } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { convertAngleUnits, ptsToEmuLenient } from '../../units-internal.js'
import { EMU_PER_POINT, ptToHundredths } from '../../units.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	createChartTextFonts,
	createGridLineElement,
	DEF_GRIDLINE_COLOR,
	genXmlTitle,
	validTimeUnit,
} from './chart-parts.js'

/**
 * The `<c:spPr>` axis-line block, identical in shape on all three axes.
 *
 * The only thing that differed between the three copies was how far each was indented — two
 * spaces on the category and series axes, one on the value axis — and that irregularity is
 * inherited, not meaningful. It is reproduced here rather than normalised because inter-element
 * whitespace being inert does not make changing it free: byte-identity is the gate that proves
 * an emitter refactor changed nothing, and it cannot tell a "harmless" whitespace edit from a
 * real one. `indent` keeps the quirk visible at the call site.
 * @param {number} indent - leading spaces on the `<c:spPr>` line; children step out by 2 and 4
 * @param {number} widthEmu - line width for `a:ln@w`
 * @param {boolean | undefined} show - false emits `<a:noFill/>` instead of a colour
 * @param {string | undefined} color - line colour, defaulting to {@link DEF_GRIDLINE_COLOR}
 * @param {string} dash - `a:prstDash@val`
 * @return {string} XML `<c:spPr>`
 */
function axisLineSpPr(
	indent: number,
	widthEmu: number,
	show: boolean | undefined,
	color: string | undefined,
	dash: string
): string {
	const at = (extra: number): string => ' '.repeat(indent + extra)
	return (
		`${at(0)}<c:spPr>` +
		`${at(2)}<a:ln w="${widthEmu}" cap="flat">` +
		(!show ? '<a:noFill/>' : genXmlColorSelection(color || DEF_GRIDLINE_COLOR)) +
		`${at(4)}<a:prstDash val="${dash}"/>` +
		`${at(4)}<a:round/>` +
		`${at(2)}</a:ln>` +
		`${at(0)}</c:spPr>`
	)
}

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
	strXml += '<c:orientation val="' + (opts.catAxisOrientation || 'minMax') + '"/>'
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
	strXml += axisLineSpPr(
		2,
		opts.catAxisLineSize ? ptsToEmuLenient(opts.catAxisLineSize) : EMU_PER_POINT,
		opts.catAxisLineShow,
		opts.catAxisLineColor,
		opts.catAxisLineStyle || 'solid'
	)
	strXml += '  <c:txPr>'
	if (opts.catAxisLabelRotate) {
		strXml += `<a:bodyPr rot="${convertAngleUnits(opts.catAxisLabelRotate, 'catAxisLabelRotate')}"/>`
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
			// All three resolved before any is emitted, so the warnings still arrive in option order.
			const baseTimeUnit = validTimeUnit(opts.catAxisBaseTimeUnit, 'catAxisBaseTimeUnit')
			const majorTimeUnit = validTimeUnit(opts.catAxisMajorTimeUnit, 'catAxisMajorTimeUnit')
			const minorTimeUnit = validTimeUnit(opts.catAxisMinorTimeUnit, 'catAxisMinorTimeUnit')
			if (baseTimeUnit) strXml += '<c:baseTimeUnit val="' + baseTimeUnit + '"/>'
			if (majorTimeUnit) strXml += voidEl('c:majorTimeUnit', { val: majorTimeUnit })
			if (minorTimeUnit) strXml += voidEl('c:minorTimeUnit', { val: minorTimeUnit })
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
	strXml += '<c:orientation val="' + (opts.valAxisOrientation || 'minMax') + '"/>'
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
	strXml += axisLineSpPr(
		1,
		opts.valAxisLineSize ? ptsToEmuLenient(opts.valAxisLineSize) : EMU_PER_POINT,
		opts.valAxisLineShow,
		opts.valAxisLineColor,
		opts.valAxisLineStyle || 'solid'
	)
	strXml += ' <c:txPr>'
	strXml += `  <a:bodyPr${opts.valAxisLabelRotate ? ' rot="' + convertAngleUnits(opts.valAxisLabelRotate, 'valAxisLabelRotate').toString() + '"' : ''}/>` // don't specify rot 0 so we get the auto behavior
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
	strXml += '  <c:scaling><c:orientation val="' + (opts.serAxisOrientation || 'minMax') + '"/></c:scaling>'
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
	strXml += `  <c:tickLblPos val="${opts.serAxisLabelPos || (opts.barDir === 'col' ? 'low' : 'nextTo')}"/>`
	strXml += axisLineSpPr(2, EMU_PER_POINT, opts.serAxisLineShow, opts.serAxisLineColor, 'solid')
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
		// All three resolved before any is emitted, so the warnings still arrive in option order.
		const baseTimeUnit = validTimeUnit(opts.serAxisBaseTimeUnit, 'serAxisBaseTimeUnit')
		const majorTimeUnit = validTimeUnit(opts.serAxisMajorTimeUnit, 'serAxisMajorTimeUnit')
		const minorTimeUnit = validTimeUnit(opts.serAxisMinorTimeUnit, 'serAxisMinorTimeUnit')
		// `baseTimeUnit` keeps its template string on purpose: it emits TWO spaces before
		// `val`, and voidEl() joins attributes with exactly one. Normalizing the spacing
		// would be a byte change, so the quirk stays visible here rather than being
		// silently "fixed" by the builder.
		if (baseTimeUnit) strXml += ` <c:baseTimeUnit  val="${baseTimeUnit}"/>`
		if (majorTimeUnit) strXml += voidEl('c:majorTimeUnit', { val: majorTimeUnit }, { openPrefix: ' ' })
		if (minorTimeUnit) strXml += voidEl('c:minorTimeUnit', { val: minorTimeUnit }, { openPrefix: ' ' })
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
