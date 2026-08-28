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
 * lines and would hide the quirks rather than share the structure. Building all three through
 * `el()` does not change that: it makes each copy's quirks explicit at its own call site,
 * which is the point.
 *
 * The other way out -- regularising the whitespace so one builder *could* serve all three --
 * is a byte change, which AGENTS.md is explicit about: a whitespace-only diff is a STOP, not
 * a known divergence. Revisit this only if the emitted whitespace in these parts is ever
 * regularised deliberately, as its own decision.
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
	const line = el(
		'a:ln',
		{ w: widthEmu, cap: 'flat' },
		[
			raw(!show ? voidEl('a:noFill') : genXmlColorSelection(color || DEF_GRIDLINE_COLOR)),
			raw(voidEl('a:prstDash', { val: dash }, { openPrefix: at(4) })),
			raw(voidEl('a:round', null, { openPrefix: at(4) })),
		],
		{ openPrefix: at(2), closePrefix: at(2) }
	)
	return el('c:spPr', null, raw(line), { openPrefix: at(0), closePrefix: at(0) })
}

export function makeCatAxis(opts: ChartOptsInternal, axisId: string, valAxisId: string): string {
	const usesValueAxisForCategories =
		opts._type === ChartType.scatter || opts._type === ChartType.bubble || opts._type === ChartType.bubble3d
	const usesCategoryAxis = !usesValueAxisForCategories && !opts.catLabelFormatCode
	// NOTE: scatter and bubble charts display numbers on the X axis, so their category axis is a
	// second value axis; a `catLabelFormatCode` means dates, which is a `<c:dateAx>`.
	const tag = usesValueAxisForCategories ? 'c:valAx' : opts.catLabelFormatCode ? 'c:dateAx' : 'c:catAx'

	const scaling = el(
		'c:scaling',
		null,
		[
			raw(voidEl('c:orientation', { val: opts.catAxisOrientation || 'minMax' })),
			opts.catAxisMaxVal || opts.catAxisMaxVal === 0 ? raw(voidEl('c:max', { val: opts.catAxisMaxVal })) : null,
			opts.catAxisMinVal || opts.catAxisMinVal === 0 ? raw(voidEl('c:min', { val: opts.catAxisMinVal })) : null,
		],
		{ openPrefix: '  ' }
	)

	// Scatter/bubble read the X format off the category option, falling back to the value one.
	const numFmt = usesValueAxisForCategories
		? voidEl(
				'c:numFmt',
				{ formatCode: (opts.catAxisLabelFormatCode ?? opts.valAxisLabelFormatCode) || 'General', sourceLinked: 1 },
				{ openPrefix: '  ' }
			)
		: voidEl(
				'c:numFmt',
				{ formatCode: (opts.catLabelFormatCode ?? '') || 'General', sourceLinked: 1 },
				{ openPrefix: '  ' }
			)

	const ticks =
		opts._type === ChartType.scatter
			? voidEl('c:majorTickMark', { val: 'none' }, { openPrefix: '  ' }) +
				voidEl('c:minorTickMark', { val: 'none' }, { openPrefix: '  ' }) +
				voidEl('c:tickLblPos', { val: opts.catAxisLabelPos || 'nextTo' }, { openPrefix: '  ' })
			: voidEl('c:majorTickMark', { val: opts.catAxisMajorTickMark || 'out' }, { openPrefix: '  ' }) +
				voidEl('c:minorTickMark', { val: opts.catAxisMinorTickMark || 'none' }, { openPrefix: '  ' }) +
				voidEl(
					'c:tickLblPos',
					{ val: opts.catAxisLabelPos || (opts.barDir === 'col' ? 'low' : 'nextTo') },
					{ openPrefix: '  ' }
				)

	const defRPr = el(
		'a:defRPr',
		{
			sz: ptToHundredths(opts.catAxisLabelFontSize || DEF_FONT_SIZE),
			b: opts.catAxisLabelFontBold ? 1 : 0,
			i: opts.catAxisLabelFontItalic ? 1 : 0,
			u: 'none',
			strike: 'noStrike',
		},
		[
			raw(genXmlColorSelection(opts.catAxisLabelColor || DEF_FONT_COLOR)),
			raw('      ' + createChartTextFonts(opts.catAxisLabelFontFace || 'Arial')),
		],
		{ openPrefix: '      ', closePrefix: '   ' }
	)
	const txPr = el(
		'c:txPr',
		null,
		[
			// NOTE: don't specify `rot="0"` — leaving it off is what gets the auto behavior.
			raw(
				voidEl('a:bodyPr', {
					rot: opts.catAxisLabelRotate ? convertAngleUnits(opts.catAxisLabelRotate, 'catAxisLabelRotate') : undefined,
				})
			),
			raw(voidEl('a:lstStyle', null, { openPrefix: '    ' })),
			raw(
				el(
					'a:p',
					null,
					[
						raw(el('a:pPr', null, raw(defRPr), { openPrefix: '    ', closePrefix: '  ' })),
						raw(voidEl('a:endParaRPr', { lang: opts.lang || 'en-US' }, { openPrefix: '  ' })),
					],
					{ openPrefix: '    ', closePrefix: '  ' }
				)
			),
		],
		{ openPrefix: '  ', closePrefix: ' ' }
	)

	const valAxisCrossTag = typeof opts.valAxisCrossesAt === 'number' ? 'crossesAt' : 'crosses'
	const valAxisCrossValue =
		typeof opts.valAxisCrossesAt === 'number' ? opts.valAxisCrossesAt : opts.valAxisCrossesAt || 'autoZero'

	// PPT auto-adjusts these once it has calculated the date bounds, so they are emitted only when
	// the caller asked for them. Major/minor units are also allowed on a double value axis.
	let units = ''
	if (opts.catLabelFormatCode || usesValueAxisForCategories) {
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
	}

	return el(tag, null, [
		raw(voidEl('c:axId', { val: axisId }, { openPrefix: '  ' })),
		raw(scaling),
		raw(voidEl('c:delete', { val: opts.catAxisHidden ? 1 : 0 }, { openPrefix: '  ' })),
		raw(voidEl('c:axPos', { val: opts.barDir === 'col' ? 'b' : 'l' }, { openPrefix: '  ' })),
		raw(opts.catGridLine && opts.catGridLine.style !== 'none' ? createGridLineElement(opts.catGridLine) : ''),
		// `<c:title>` comes between `</c:majorGridlines>` and `<c:numFmt>`.
		opts.showCatAxisTitle
			? raw(
					genXmlTitle({
						color: opts.catAxisTitleColor,
						fontFace: opts.catAxisTitleFontFace,
						fontSize: opts.catAxisTitleFontSize,
						titleRotate: opts.catAxisTitleRotate,
						title: opts.catAxisTitle || 'Axis Title',
					})
				)
			: null,
		raw(numFmt),
		raw(ticks),
		raw(
			axisLineSpPr(
				2,
				opts.catAxisLineSize ? ptsToEmuLenient(opts.catAxisLineSize) : EMU_PER_POINT,
				opts.catAxisLineShow,
				opts.catAxisLineColor,
				opts.catAxisLineStyle || 'solid'
			)
		),
		raw(txPr),
		raw(voidEl('c:crossAx', { val: valAxisId }, { openPrefix: ' ' })),
		raw(voidEl(`c:${valAxisCrossTag}`, { val: valAxisCrossValue }, { openPrefix: ' ' })),
		usesValueAxisForCategories ? null : raw(voidEl('c:auto', { val: 1 }, { openPrefix: ' ' })),
		usesCategoryAxis ? raw(voidEl('c:lblAlgn', { val: 'ctr' }, { openPrefix: ' ' })) : null,
		usesCategoryAxis && opts.catAxisLabelFrequency
			? raw(voidEl('c:tickLblSkip', { val: opts.catAxisLabelFrequency }, { openPrefix: ' ' }))
			: null,
		usesCategoryAxis
			? raw(voidEl('c:noMultiLvlLbl', { val: opts.catAxisMultiLevelLabels ? 0 : 1 }, { openPrefix: ' ' }))
			: null,
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

	const scaling = el(
		'c:scaling',
		null,
		[
			opts.valAxisLogScaleBase ? raw(voidEl('c:logBase', { val: opts.valAxisLogScaleBase })) : null,
			raw(voidEl('c:orientation', { val: opts.valAxisOrientation || 'minMax' })),
			opts.valAxisMaxVal || opts.valAxisMaxVal === 0 ? raw(voidEl('c:max', { val: opts.valAxisMaxVal })) : null,
			opts.valAxisMinVal || opts.valAxisMinVal === 0 ? raw(voidEl('c:min', { val: opts.valAxisMinVal })) : null,
		],
		{ openPrefix: '  ', closePrefix: '  ' }
	)

	const ticks =
		opts._type === ChartType.scatter
			? voidEl('c:majorTickMark', { val: 'none' }, { openPrefix: '  ' }) +
				voidEl('c:minorTickMark', { val: 'none' }, { openPrefix: '  ' }) +
				voidEl('c:tickLblPos', { val: 'nextTo' }, { openPrefix: '  ' })
			: voidEl('c:majorTickMark', { val: opts.valAxisMajorTickMark || 'out' }, { openPrefix: ' ' }) +
				voidEl('c:minorTickMark', { val: opts.valAxisMinorTickMark || 'none' }, { openPrefix: ' ' }) +
				voidEl(
					'c:tickLblPos',
					{ val: opts.valAxisLabelPos || (opts.barDir === 'col' ? 'nextTo' : 'low') },
					{ openPrefix: ' ' }
				)

	const defRPr = el(
		'a:defRPr',
		{
			sz: ptToHundredths(opts.valAxisLabelFontSize || DEF_FONT_SIZE),
			b: opts.valAxisLabelFontBold ? 1 : 0,
			i: opts.valAxisLabelFontItalic ? 1 : 0,
			u: 'none',
			strike: 'noStrike',
		},
		[
			raw(genXmlColorSelection(opts.valAxisLabelColor || DEF_FONT_COLOR)),
			raw('        ' + createChartTextFonts(opts.valAxisLabelFontFace || 'Arial')),
		],
		{ openPrefix: '      ', closePrefix: '      ' }
	)
	const txPr = el(
		'c:txPr',
		null,
		[
			// Don't specify `rot="0"`, so we get the auto behavior.
			raw(
				voidEl(
					'a:bodyPr',
					{
						rot: opts.valAxisLabelRotate ? convertAngleUnits(opts.valAxisLabelRotate, 'valAxisLabelRotate') : undefined,
					},
					{ openPrefix: '  ' }
				)
			),
			raw(voidEl('a:lstStyle', null, { openPrefix: '  ' })),
			raw(
				el(
					'a:p',
					null,
					[
						raw(el('a:pPr', null, raw(defRPr), { openPrefix: '    ', closePrefix: '    ' })),
						raw(voidEl('a:endParaRPr', { lang: opts.lang || 'en-US' }, { openPrefix: '  ' })),
					],
					{ openPrefix: '  ', closePrefix: '  ' }
				)
			),
		],
		{ openPrefix: ' ', closePrefix: ' ' }
	)

	// Where this axis meets its category axis: an explicit position, an explicit rule, or the
	// default — a right/top axis crosses at the maximum, everything else at zero.
	const crosses =
		typeof opts.catAxisCrossesAt === 'number'
			? voidEl('c:crossesAt', { val: opts.catAxisCrossesAt }, { openPrefix: ' ' })
			: typeof opts.catAxisCrossesAt === 'string'
				? voidEl('c:crosses', { val: opts.catAxisCrossesAt }, { openPrefix: ' ' })
				: voidEl('c:crosses', { val: axisPos === 'r' || axisPos === 't' ? 'max' : 'autoZero' }, { openPrefix: ' ' })
	const crossBetween =
		opts.valAxisCrossBetween ||
		(opts._type === ChartType.scatter ||
		!!(Array.isArray(opts._type) && opts._type.some((type) => asChartType(type.type) === ChartType.area))
			? 'midCat'
			: 'between')

	return el('c:valAx', null, [
		raw(voidEl('c:axId', { val: valAxisId }, { openPrefix: '  ' })),
		raw(scaling),
		raw(voidEl('c:delete', { val: opts.valAxisHidden ? 1 : 0 }, { openPrefix: '  ' })),
		raw(voidEl('c:axPos', { val: axisPos }, { openPrefix: '  ' })),
		opts.valGridLine && opts.valGridLine.style !== 'none' ? raw(createGridLineElement(opts.valGridLine)) : null,
		// `<c:title>` comes between `</c:majorGridlines>` and `<c:numFmt>`.
		opts.showValAxisTitle
			? raw(
					genXmlTitle({
						color: opts.valAxisTitleColor,
						fontFace: opts.valAxisTitleFontFace,
						fontSize: opts.valAxisTitleFontSize,
						titleRotate: opts.valAxisTitleRotate,
						title: opts.valAxisTitle || 'Axis Title',
					})
				)
			: null,
		raw(voidEl('c:numFmt', { formatCode: opts.valAxisLabelFormatCode || 'General', sourceLinked: 0 })),
		raw(ticks),
		raw(
			axisLineSpPr(
				1,
				opts.valAxisLineSize ? ptsToEmuLenient(opts.valAxisLineSize) : EMU_PER_POINT,
				opts.valAxisLineShow,
				opts.valAxisLineColor,
				opts.valAxisLineStyle || 'solid'
			)
		),
		raw(txPr),
		raw(voidEl('c:crossAx', { val: crossAxId }, { openPrefix: ' ' })),
		raw(crosses),
		raw(voidEl('c:crossBetween', { val: crossBetween }, { openPrefix: ' ' })),
		opts.valAxisMajorUnit ? raw(voidEl('c:majorUnit', { val: opts.valAxisMajorUnit }, { openPrefix: ' ' })) : null,
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
	const defRPr = el(
		'a:defRPr',
		{
			sz: ptToHundredths(opts.serAxisLabelFontSize || DEF_FONT_SIZE),
			b: opts.serAxisLabelFontBold ? 1 : 0,
			i: opts.serAxisLabelFontItalic ? 1 : 0,
			u: 'none',
			strike: 'noStrike',
		},
		[
			raw('      ' + genXmlColorSelection(opts.serAxisLabelColor || DEF_FONT_COLOR)),
			raw('      ' + createChartTextFonts(opts.serAxisLabelFontFace || 'Arial')),
		],
		{ openPrefix: '    ', closePrefix: '   ' }
	)
	const txPr = el(
		'c:txPr',
		null,
		[
			// Don't specify `rot="0"`, so we get the auto behavior.
			raw(voidEl('a:bodyPr', null, { openPrefix: '    ' })),
			raw(voidEl('a:lstStyle', null, { openPrefix: '    ' })),
			raw(
				el(
					'a:p',
					null,
					[
						raw(el('a:pPr', null, raw(defRPr), { openPrefix: '    ', closePrefix: '  ' })),
						raw(voidEl('a:endParaRPr', { lang: opts.lang || 'en-US' }, { openPrefix: '  ' })),
					],
					{ openPrefix: '    ', closePrefix: '  ' }
				)
			),
		],
		{ openPrefix: '  ', closePrefix: ' ' }
	)

	// PPT auto-adjusts these once it has calculated the date bounds, so they are emitted only when
	// the caller asked for them.
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
		if (majorTimeUnit) units += voidEl('c:majorTimeUnit', { val: majorTimeUnit }, { openPrefix: ' ' })
		if (minorTimeUnit) units += voidEl('c:minorTimeUnit', { val: minorTimeUnit }, { openPrefix: ' ' })
		if (opts.serAxisMajorUnit) units += voidEl('c:majorUnit', { val: opts.serAxisMajorUnit }, { openPrefix: ' ' })
		if (opts.serAxisMinorUnit) units += voidEl('c:minorUnit', { val: opts.serAxisMinorUnit }, { openPrefix: ' ' })
	}

	return el('c:serAx', null, [
		raw(voidEl('c:axId', { val: axisId }, { openPrefix: '  ' })),
		raw(
			el('c:scaling', null, raw(voidEl('c:orientation', { val: opts.serAxisOrientation || 'minMax' })), {
				openPrefix: '  ',
			})
		),
		raw(voidEl('c:delete', { val: opts.serAxisHidden ? 1 : 0 }, { openPrefix: '  ' })),
		raw(voidEl('c:axPos', { val: opts.barDir === 'col' ? 'b' : 'l' }, { openPrefix: '  ' })),
		raw(opts.serGridLine && opts.serGridLine.style !== 'none' ? createGridLineElement(opts.serGridLine) : ''),
		// `<c:title>` comes between `</c:majorGridlines>` and `<c:numFmt>`.
		opts.showSerAxisTitle
			? raw(
					genXmlTitle({
						color: opts.serAxisTitleColor,
						fontFace: opts.serAxisTitleFontFace,
						fontSize: opts.serAxisTitleFontSize,
						titleRotate: opts.serAxisTitleRotate,
						title: opts.serAxisTitle || 'Axis Title',
					})
				)
			: null,
		raw(
			voidEl(
				'c:numFmt',
				{ formatCode: (opts.serLabelFormatCode ?? '') || 'General', sourceLinked: 0 },
				{ openPrefix: '  ' }
			)
		),
		raw(voidEl('c:majorTickMark', { val: 'out' }, { openPrefix: '  ' })),
		raw(voidEl('c:minorTickMark', { val: 'none' }, { openPrefix: '  ' })),
		raw(
			voidEl(
				'c:tickLblPos',
				{ val: opts.serAxisLabelPos || (opts.barDir === 'col' ? 'low' : 'nextTo') },
				{ openPrefix: '  ' }
			)
		),
		raw(axisLineSpPr(2, EMU_PER_POINT, opts.serAxisLineShow, opts.serAxisLineColor, 'solid')),
		raw(txPr),
		raw(voidEl('c:crossAx', { val: valAxisId }, { openPrefix: ' ' })),
		raw(voidEl('c:crosses', { val: 'autoZero' }, { openPrefix: ' ' })),
		opts.serAxisLabelFrequency
			? raw(voidEl('c:tickLblSkip', { val: opts.serAxisLabelFrequency }, { openPrefix: ' ' }))
			: null,
		raw(units),
	])
}
