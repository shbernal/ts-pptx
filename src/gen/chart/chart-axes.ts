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
 * `<c:spPr>` and {@link axisTextProps} the whole `<c:txPr>`. Both used to take indentation
 * arguments, which is what had kept them copied rather than shared; the emitters are flat now
 * (`docs/chart-whitespace-flatten.md`), so they take only what they are about.
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
	axisCrossing,
	createChartTextFonts,
	createGridLineElement,
	DEF_GRIDLINE_COLOR,
	genXmlTitle,
	positiveIntAttr,
	validTimeUnit,
} from './chart-parts.js'
import { isScatterChart, isXyChart } from './chart-kind.js'

/**
 * The `<c:spPr>` axis-line block, identical on all three axes.
 * @param {number} widthEmu - line width for `a:ln@w`
 * @param {boolean | undefined} show - false emits `<a:noFill/>` instead of a colour
 * @param {string | undefined} color - line colour, defaulting to {@link DEF_GRIDLINE_COLOR}
 * @param {string} dash - `a:prstDash@val`
 * @return {string} XML `<c:spPr>`
 */
function axisLineSpPr(widthEmu: number, show: boolean | undefined, color: string | undefined, dash: string): string {
	const line = el('a:ln', { w: widthEmu, cap: 'flat' }, [
		raw(!show ? voidEl('a:noFill') : genXmlColorSelection(color || DEF_GRIDLINE_COLOR)),
		raw(voidEl('a:prstDash', { val: dash })),
		raw(voidEl('a:round', null)),
	])
	return el('c:spPr', null, raw(line))
}

/**
 * The `<c:txPr>` an axis wraps its label run properties in: `<a:bodyPr>`, an empty
 * `<a:lstStyle>`, then the `<a:pPr>` carrying `defRPr` and an empty `<a:endParaRPr>`. All three
 * axis builders emit exactly this; only the rotation differs, and the series axis has none.
 *
 * @param defRPr - the already-built `<a:defRPr>`
 * @param lang - the deck's language tag, for `endParaRPr`
 * @param rot - `a:bodyPr@rot`, already in 60000ths of a degree. `undefined` omits the
 *   attribute, which is what gets the auto behaviour -- do not pass `0` for it.
 */
function axisTextProps(defRPr: string, lang: string, rot?: number): string {
	return el('c:txPr', null, [
		raw(voidEl('a:bodyPr', { rot })),
		raw(voidEl('a:lstStyle', null)),
		raw(el('a:p', null, [raw(el('a:pPr', null, raw(defRPr))), raw(voidEl('a:endParaRPr', { lang }))])),
	])
}

/** One axis' label font, as the three axis builders each spelled it. */
interface AxisLabelFont {
	size: number | undefined
	bold: boolean | undefined
	italic: boolean | undefined
	color: string | undefined
	fontFace: string | undefined
}

/**
 * The `<a:defRPr>` an axis' tick labels carry, in the `sz, b, i, u, strike` order all three
 * axes emit.
 *
 * Three copies differed only in the `catAxis…`/`valAxis…`/`serAxis…` option prefix. The header
 * of this module explains that the rest of the axis blocks stayed copied because the emitters
 * took indentation arguments; the flatten removed that reason and left these behind.
 */
function axisLabelDefRPr(font: AxisLabelFont): string {
	return el(
		'a:defRPr',
		{
			sz: ptToHundredths(font.size || DEF_FONT_SIZE),
			b: font.bold ? 1 : 0,
			i: font.italic ? 1 : 0,
			u: 'none',
			strike: 'noStrike',
		},
		[raw(genXmlColorSelection(font.color || DEF_FONT_COLOR)), raw(createChartTextFonts(font.fontFace || 'Arial'))]
	)
}

/** One axis' `<c:title>` as the three builders each spelled it, down to the shared default. */
interface AxisTitle {
	text: string | undefined
	color: string | undefined
	fontFace: string | undefined
	fontSize: number | undefined
	rotate: number | undefined
}

/** The axis `<c:title>`, or `''` when the caller asked for none. */
function axisTitleXml(show: boolean | undefined, title: AxisTitle): string {
	if (!show) return ''
	return genXmlTitle({
		color: title.color,
		fontFace: title.fontFace,
		fontSize: title.fontSize,
		titleRotate: title.rotate,
		title: title.text || 'Axis Title',
	})
}

export function makeCatAxis(opts: ChartOptsInternal, axisId: string, valAxisId: string): string {
	const usesValueAxisForCategories = isXyChart(opts._type)
	const usesCategoryAxis = !usesValueAxisForCategories && !opts.catLabelFormatCode
	// NOTE: scatter and bubble charts display numbers on the X axis, so their category axis is a
	// second value axis; a `catLabelFormatCode` means dates, which is a `<c:dateAx>`.
	const tag = usesValueAxisForCategories ? 'c:valAx' : opts.catLabelFormatCode ? 'c:dateAx' : 'c:catAx'

	const scaling = el('c:scaling', null, [
		raw(voidEl('c:orientation', { val: opts.catAxisOrientation || 'minMax' })),
		opts.catAxisMaxVal || opts.catAxisMaxVal === 0 ? raw(voidEl('c:max', { val: opts.catAxisMaxVal })) : null,
		opts.catAxisMinVal || opts.catAxisMinVal === 0 ? raw(voidEl('c:min', { val: opts.catAxisMinVal })) : null,
	])

	// Scatter/bubble read the X format off the category option, falling back to the value one.
	const numFmt = usesValueAxisForCategories
		? voidEl('c:numFmt', {
				formatCode: (opts.catAxisLabelFormatCode ?? opts.valAxisLabelFormatCode) || 'General',
				sourceLinked: 1,
			})
		: voidEl('c:numFmt', { formatCode: (opts.catLabelFormatCode ?? '') || 'General', sourceLinked: 1 })

	const ticks = isScatterChart(opts._type)
		? voidEl('c:majorTickMark', { val: 'none' }) +
			voidEl('c:minorTickMark', { val: 'none' }) +
			voidEl('c:tickLblPos', { val: opts.catAxisLabelPos || 'nextTo' })
		: voidEl('c:majorTickMark', { val: opts.catAxisMajorTickMark || 'out' }) +
			voidEl('c:minorTickMark', { val: opts.catAxisMinorTickMark || 'none' }) +
			voidEl('c:tickLblPos', { val: opts.catAxisLabelPos || (opts.barDir === 'col' ? 'low' : 'nextTo') })

	const defRPr = axisLabelDefRPr({
		size: opts.catAxisLabelFontSize,
		bold: opts.catAxisLabelFontBold,
		italic: opts.catAxisLabelFontItalic,
		color: opts.catAxisLabelColor,
		fontFace: opts.catAxisLabelFontFace,
	})
	const txPr = axisTextProps(
		defRPr,
		opts.lang || 'en-US',
		opts.catAxisLabelRotate ? convertAngleUnits(opts.catAxisLabelRotate, 'catAxisLabelRotate') : undefined
	)

	const catLabelSkip = positiveIntAttr(opts.catAxisLabelFrequency, 'catAxisLabelFrequency')
	const valAxisCrossing = axisCrossing(opts.valAxisCrossesAt, 'autoZero', 'valAxisCrossesAt')

	// The TIME units are gated on a format code because they describe a date axis and PowerPoint
	// auto-adjusts them once it has calculated the date bounds. The numeric major/minor units are
	// not: they are an `xsd:double` spacing that applies to any category axis, and the value axis
	// below has always emitted its own with no gate at all. Gating them here meant
	// `{ type: 'bar3d', catAxisMajorUnit: 3, valAxisMajorUnit: 4 }` emitted exactly one element.
	let units = ''
	if (opts.catLabelFormatCode) {
		// All three resolved before any is emitted, so the warnings still arrive in option order.
		const baseTimeUnit = validTimeUnit(opts.catAxisBaseTimeUnit, 'catAxisBaseTimeUnit')
		const majorTimeUnit = validTimeUnit(opts.catAxisMajorTimeUnit, 'catAxisMajorTimeUnit')
		const minorTimeUnit = validTimeUnit(opts.catAxisMinorTimeUnit, 'catAxisMinorTimeUnit')
		if (baseTimeUnit) units += voidEl('c:baseTimeUnit', { val: baseTimeUnit })
		if (majorTimeUnit) units += voidEl('c:majorTimeUnit', { val: majorTimeUnit })
		if (minorTimeUnit) units += voidEl('c:minorTimeUnit', { val: minorTimeUnit })
	}
	if (opts.catAxisMajorUnit) units += voidEl('c:majorUnit', { val: opts.catAxisMajorUnit })
	if (opts.catAxisMinorUnit) units += voidEl('c:minorUnit', { val: opts.catAxisMinorUnit })

	return el(tag, null, [
		raw(voidEl('c:axId', { val: axisId })),
		raw(scaling),
		raw(voidEl('c:delete', { val: opts.catAxisHidden ? 1 : 0 })),
		raw(voidEl('c:axPos', { val: opts.barDir === 'col' ? 'b' : 'l' })),
		raw(opts.catGridLine && opts.catGridLine.style !== 'none' ? createGridLineElement(opts.catGridLine) : ''),
		// `<c:title>` comes between `</c:majorGridlines>` and `<c:numFmt>`.
		raw(
			axisTitleXml(opts.showCatAxisTitle, {
				text: opts.catAxisTitle,
				color: opts.catAxisTitleColor,
				fontFace: opts.catAxisTitleFontFace,
				fontSize: opts.catAxisTitleFontSize,
				rotate: opts.catAxisTitleRotate,
			})
		),
		raw(numFmt),
		raw(ticks),
		raw(
			axisLineSpPr(
				opts.catAxisLineSize ? ptsToEmuLenient(opts.catAxisLineSize) : EMU_PER_POINT,
				opts.catAxisLineShow,
				opts.catAxisLineColor,
				opts.catAxisLineStyle || 'solid'
			)
		),
		raw(txPr),
		raw(voidEl('c:crossAx', { val: valAxisId })),
		raw(valAxisCrossing),
		usesValueAxisForCategories ? null : raw(voidEl('c:auto', { val: 1 })),
		usesCategoryAxis ? raw(voidEl('c:lblAlgn', { val: 'ctr' })) : null,
		usesCategoryAxis && catLabelSkip !== undefined ? raw(voidEl('c:tickLblSkip', { val: catLabelSkip })) : null,
		usesCategoryAxis ? raw(voidEl('c:noMultiLvlLbl', { val: opts.catAxisMultiLevelLabels ? 0 : 1 })) : null,
		raw(units),
	])
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

	const scaling = el('c:scaling', null, [
		opts.valAxisLogScaleBase ? raw(voidEl('c:logBase', { val: opts.valAxisLogScaleBase })) : null,
		raw(voidEl('c:orientation', { val: opts.valAxisOrientation || 'minMax' })),
		opts.valAxisMaxVal || opts.valAxisMaxVal === 0 ? raw(voidEl('c:max', { val: opts.valAxisMaxVal })) : null,
		opts.valAxisMinVal || opts.valAxisMinVal === 0 ? raw(voidEl('c:min', { val: opts.valAxisMinVal })) : null,
	])

	const ticks = isScatterChart(opts._type)
		? voidEl('c:majorTickMark', { val: 'none' }) +
			voidEl('c:minorTickMark', { val: 'none' }) +
			voidEl('c:tickLblPos', { val: 'nextTo' })
		: voidEl('c:majorTickMark', { val: opts.valAxisMajorTickMark || 'out' }) +
			voidEl('c:minorTickMark', { val: opts.valAxisMinorTickMark || 'none' }) +
			voidEl('c:tickLblPos', { val: opts.valAxisLabelPos || (opts.barDir === 'col' ? 'nextTo' : 'low') })

	const defRPr = axisLabelDefRPr({
		size: opts.valAxisLabelFontSize,
		bold: opts.valAxisLabelFontBold,
		italic: opts.valAxisLabelFontItalic,
		color: opts.valAxisLabelColor,
		fontFace: opts.valAxisLabelFontFace,
	})
	const txPr = axisTextProps(
		defRPr,
		opts.lang || 'en-US',
		opts.valAxisLabelRotate ? convertAngleUnits(opts.valAxisLabelRotate, 'valAxisLabelRotate') : undefined
	)

	// Where this axis meets its category axis: an explicit position, an explicit rule, or the
	// default — a right/top axis crosses at the maximum, everything else at zero.
	const crosses = axisCrossing(
		opts.catAxisCrossesAt,
		axisPos === 'r' || axisPos === 't' ? 'max' : 'autoZero',
		'catAxisCrossesAt'
	)
	const crossBetween =
		opts.valAxisCrossBetween ||
		(isScatterChart(opts._type) ||
		!!(Array.isArray(opts._type) && opts._type.some((type) => asChartType(type.type) === ChartType.area))
			? 'midCat'
			: 'between')

	return el('c:valAx', null, [
		raw(voidEl('c:axId', { val: valAxisId })),
		raw(scaling),
		raw(voidEl('c:delete', { val: opts.valAxisHidden ? 1 : 0 })),
		raw(voidEl('c:axPos', { val: axisPos })),
		opts.valGridLine && opts.valGridLine.style !== 'none' ? raw(createGridLineElement(opts.valGridLine)) : null,
		// `<c:title>` comes between `</c:majorGridlines>` and `<c:numFmt>`.
		raw(
			axisTitleXml(opts.showValAxisTitle, {
				text: opts.valAxisTitle,
				color: opts.valAxisTitleColor,
				fontFace: opts.valAxisTitleFontFace,
				fontSize: opts.valAxisTitleFontSize,
				rotate: opts.valAxisTitleRotate,
			})
		),
		raw(voidEl('c:numFmt', { formatCode: opts.valAxisLabelFormatCode || 'General', sourceLinked: 0 })),
		raw(ticks),
		raw(
			axisLineSpPr(
				opts.valAxisLineSize ? ptsToEmuLenient(opts.valAxisLineSize) : EMU_PER_POINT,
				opts.valAxisLineShow,
				opts.valAxisLineColor,
				opts.valAxisLineStyle || 'solid'
			)
		),
		raw(txPr),
		raw(voidEl('c:crossAx', { val: crossAxId })),
		raw(crosses),
		raw(voidEl('c:crossBetween', { val: crossBetween })),
		opts.valAxisMajorUnit ? raw(voidEl('c:majorUnit', { val: opts.valAxisMajorUnit })) : null,
		opts.valAxisMinorUnit ? raw(voidEl('c:minorUnit', { val: opts.valAxisMinorUnit })) : null,
		opts.valAxisDisplayUnit
			? raw(
					el('c:dispUnits', null, [
						raw(voidEl('c:builtInUnit', { val: opts.valAxisDisplayUnit })),
						raw(opts.valAxisDisplayUnitLabel ? voidEl('c:dispUnitsLbl') : ''),
					])
				)
			: null,
	])
}

/**
 * Create Series Axis (Used by `bar3D`)
 * @param {ChartOptsInternal} opts - chart options
 * @param {string} axisId - axis ID
 * @param {string} valAxisId - value
 * @return {string} XML
 */
export function makeSerAxis(opts: ChartOptsInternal, axisId: string, valAxisId: string): string {
	const defRPr = axisLabelDefRPr({
		size: opts.serAxisLabelFontSize,
		bold: opts.serAxisLabelFontBold,
		italic: opts.serAxisLabelFontItalic,
		color: opts.serAxisLabelColor,
		fontFace: opts.serAxisLabelFontFace,
	})
	// No `serAxisLabelRotate` option exists, so the series axis always takes the auto rotation.
	const txPr = axisTextProps(defRPr, opts.lang || 'en-US')
	const serLabelSkip = positiveIntAttr(opts.serAxisLabelFrequency, 'serAxisLabelFrequency')

	// Time units on a format code, numeric units unconditionally — the same split as the
	// category axis above and the value axis below. See the note there.
	let units = ''
	if (opts.serLabelFormatCode) {
		// All three resolved before any is emitted, so the warnings still arrive in option order.
		const baseTimeUnit = validTimeUnit(opts.serAxisBaseTimeUnit, 'serAxisBaseTimeUnit')
		const majorTimeUnit = validTimeUnit(opts.serAxisMajorTimeUnit, 'serAxisMajorTimeUnit')
		const minorTimeUnit = validTimeUnit(opts.serAxisMinorTimeUnit, 'serAxisMinorTimeUnit')
		// `baseTimeUnit` keeps its template string on purpose: it emits TWO spaces before
		// `val`, and voidEl() joins attributes with exactly one. Normalizing the spacing
		// would be a byte change, so the quirk stays visible here rather than being
		// silently "fixed" by the builder.
		if (baseTimeUnit) units += ` <c:baseTimeUnit  val="${baseTimeUnit}"/>`
		if (majorTimeUnit) units += voidEl('c:majorTimeUnit', { val: majorTimeUnit })
		if (minorTimeUnit) units += voidEl('c:minorTimeUnit', { val: minorTimeUnit })
	}
	if (opts.serAxisMajorUnit) units += voidEl('c:majorUnit', { val: opts.serAxisMajorUnit })
	if (opts.serAxisMinorUnit) units += voidEl('c:minorUnit', { val: opts.serAxisMinorUnit })

	return el('c:serAx', null, [
		raw(voidEl('c:axId', { val: axisId })),
		raw(el('c:scaling', null, raw(voidEl('c:orientation', { val: opts.serAxisOrientation || 'minMax' })))),
		raw(voidEl('c:delete', { val: opts.serAxisHidden ? 1 : 0 })),
		raw(voidEl('c:axPos', { val: opts.barDir === 'col' ? 'b' : 'l' })),
		raw(opts.serGridLine && opts.serGridLine.style !== 'none' ? createGridLineElement(opts.serGridLine) : ''),
		// `<c:title>` comes between `</c:majorGridlines>` and `<c:numFmt>`.
		raw(
			axisTitleXml(opts.showSerAxisTitle, {
				text: opts.serAxisTitle,
				color: opts.serAxisTitleColor,
				fontFace: opts.serAxisTitleFontFace,
				fontSize: opts.serAxisTitleFontSize,
				rotate: opts.serAxisTitleRotate,
			})
		),
		raw(voidEl('c:numFmt', { formatCode: (opts.serLabelFormatCode ?? '') || 'General', sourceLinked: 0 })),
		raw(voidEl('c:majorTickMark', { val: 'out' })),
		raw(voidEl('c:minorTickMark', { val: 'none' })),
		raw(voidEl('c:tickLblPos', { val: opts.serAxisLabelPos || (opts.barDir === 'col' ? 'low' : 'nextTo') })),
		raw(axisLineSpPr(EMU_PER_POINT, opts.serAxisLineShow, opts.serAxisLineColor, 'solid')),
		raw(txPr),
		raw(voidEl('c:crossAx', { val: valAxisId })),
		raw(voidEl('c:crosses', { val: 'autoZero' })),
		serLabelSkip !== undefined ? raw(voidEl('c:tickLblSkip', { val: serLabelSkip })) : null,
		raw(units),
	])
}
