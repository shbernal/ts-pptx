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
import type { ChartDataPointStyle } from '../../types/chart.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { resolveBorderWidth } from '../drawingml/line.js'
import { ptsToEmuLenient } from '../../units-internal.js'
import { ptToHundredths } from '../../units.js'
import { dataValues, firstLabelGroup } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	createChartBorderLine,
	createChartTextFonts,
	createLeaderLinesElement,
	paletteColor,
	resolveChartPalette,
	strRefBlock,
} from './chart-parts.js'

/** The label run properties both `<c:txPr>` spellings share, differing only in their indentation. */
function labelDefRPr(opts: ChartOptsInternal, fontsIndent: string, fmt: { openPrefix: string; closePrefix: string }) {
	return el(
		'a:defRPr',
		{
			sz: ptToHundredths(opts.dataLabelFontSize || DEF_FONT_SIZE),
			b: opts.dataLabelFontBold ? 1 : 0,
			i: opts.dataLabelFontItalic ? 1 : 0,
			u: 'none',
			strike: 'noStrike',
		},
		[
			raw(genXmlColorSelection(opts.dataLabelColor || DEF_FONT_COLOR)),
			raw(fontsIndent + createChartTextFonts(opts.dataLabelFontFace || 'Arial')),
		],
		fmt
	)
}

/** One `<c:dPt>`: the slice's own fill, its border override, and the shared shadow. */
function pieDataPoint(
	idx: number,
	ptStyle: ChartDataPointStyle | undefined,
	opts: ChartOptsInternal,
	fallbackColor: string
): string {
	// A per-point border override takes precedence over the chart-level `dataBorder`.
	const border = ptStyle?.border
		? createChartBorderLine(ptStyle.border)
		: opts.dataBorder
			? el('a:ln', { w: ptsToEmuLenient(resolveBorderWidth(opts.dataBorder, 0.75)), cap: 'flat' }, [
					raw(
						genXmlColorSelection({
							color: opts.dataBorder.color ?? '363636',
							transparency: opts.dataBorder.transparency,
						})
					),
					raw(voidEl('a:prstDash', { val: 'solid' })),
					raw(voidEl('a:round')),
				])
			: ''
	const spPr = el(
		'c:spPr',
		null,
		[
			raw(el('a:solidFill', null, raw(createColorElement(ptStyle?.fill || fallbackColor)))),
			raw(border),
			raw(createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)),
		],
		{ openPrefix: ' ', closePrefix: '  ' }
	)
	return el('c:dPt', null, [
		raw(voidEl('c:idx', { val: idx }, { openPrefix: ' ' })),
		raw(voidEl('c:bubble3D', { val: 0 }, { openPrefix: ' ' })),
		raw(spPr),
	])
}

/**
 * One `<c:dLbl>`: a pie labels its points, not its series, so every slice carries its own copy of
 * the number format, text style and show flags. A `customLabels` entry replaces the value with
 * literal rich text, which is why it also forces `<c:showVal>` off.
 */
function pieDataLabel(idx: number, customLbl: string | undefined, opts: ChartOptsInternal, chartType: ChartType) {
	const numFmt = voidEl(
		'c:numFmt',
		{ formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 },
		{ openPrefix: '  ' }
	)
	const txPr = el(
		'c:txPr',
		null,
		[
			raw(voidEl('a:bodyPr', null, { openPrefix: '   ' })),
			raw(voidEl('a:lstStyle')),
			raw(
				el(
					'a:p',
					null,
					raw(
						el('a:pPr', null, raw(labelDefRPr(opts, '    ', { openPrefix: '   ', closePrefix: '   ' })), {
							closePrefix: '      ',
						})
					),
					{ openPrefix: '   ' }
				)
			),
		],
		{ closePrefix: '    ' }
	)
	return el(
		'c:dLbl',
		null,
		[
			raw(voidEl('c:idx', { val: idx }, { openPrefix: ' ' })),
			// `c:tx` must precede `c:numFmt` per CT_DLbl / Group_DLbl / EG_DLblShared schema order.
			customLbl
				? raw(
						el(
							'c:tx',
							null,
							raw(
								el('c:rich', null, [
									raw(voidEl('a:bodyPr')),
									raw(voidEl('a:lstStyle')),
									raw(
										el(
											'a:p',
											null,
											raw(
												el('a:r', null, [
													raw(voidEl('a:rPr', { lang: opts.lang || 'en-US', dirty: 0 })),
													raw(el('a:t', null, customLbl)),
												])
											)
										)
									),
								])
							)
						)
					)
				: null,
			raw(numFmt),
			raw(voidEl('c:spPr', null, { openPrefix: '  ' })),
			raw(txPr),
			chartType === ChartType.pie && opts.dataLabelPosition
				? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition }))
				: null,
			raw(voidEl('c:showLegendKey', { val: 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showVal', { val: customLbl ? 0 : opts.showValue ? 1 : 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showCatName', { val: opts.showLabel ? 1 : 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showSerName', { val: opts.showSerName ? 1 : 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showPercent', { val: opts.showPercent ? 1 : 0 }, { openPrefix: '    ' })),
			raw(voidEl('c:showBubbleSize', { val: 0 }, { openPrefix: '    ' })),
		],
		{ closePrefix: '  ' }
	)
}

/** The `<c:cat>` slice-name reference and the `<c:val>` cache, both keyed on the label count. */
function pieCategories(labels: string[]): string {
	const strCache = el(
		'c:strCache',
		null,
		[
			raw(voidEl('c:ptCount', { val: labels.length }, { openPrefix: '         ' })),
			raw(labels.map((label, idx) => el('c:pt', { idx }, raw(el('c:v', null, label)))).join('')),
		],
		{ openPrefix: '    ', closePrefix: '    ' }
	)
	const strRef = el(
		'c:strRef',
		null,
		[raw(el('c:f', null, `Sheet1!$A$2:$A$${labels.length + 1}`, { openPrefix: '    ' })), raw(strCache)],
		{ openPrefix: '  ', closePrefix: '  ' }
	)
	return el('c:cat', null, raw(strRef))
}

/** The `<c:val>` numeric cache. A missing value keeps its `<c:pt>` with an empty `<c:v>`. */
function pieValues(obj: OptsChartDataInternal, count: number, valFmtCode: string): string {
	const points = dataValues(obj)
		.map((value, idx) => el('c:pt', { idx }, raw(el('c:v', null, value || value === 0 ? value : ''))))
		.join('')
	const numCache = el(
		'c:numCache',
		null,
		[
			// `valFmtCode` arrives ALREADY ESCAPED — `chart-xml.ts` runs the option through
			// `encodeXmlEntities` once and hands the same string to all five plot emitters — so it goes
			// in as `raw`. A text child would escape it a second time and turn a user's `0"A&B"` from
			// `0&quot;A&amp;B&quot;` into `0&amp;quot;A&amp;amp;B&amp;quot;`.
			raw(el('c:formatCode', null, raw(valFmtCode), { openPrefix: '        ' })),
			raw(voidEl('c:ptCount', { val: count }, { openPrefix: '           ' })),
			raw(points),
		],
		{ openPrefix: '      ', closePrefix: '      ' }
	)
	const numRef = el(
		'c:numRef',
		null,
		[raw(el('c:f', null, `Sheet1!$B$2:$B$${count + 1}`, { openPrefix: '      ' })), raw(numCache)],
		{ openPrefix: '    ', closePrefix: '    ' }
	)
	return el('c:val', null, raw(numRef), { openPrefix: '  ', closePrefix: '  ' })
}

/**
 * Plot a single-series pie / doughnut chart into `<c:pieChart>` / `<c:doughnutChart>`.
 */
export function makePiePlot(
	chartType: ChartType,
	data: OptsChartDataInternal[],
	opts: ChartOptsInternal,
	valFmtCode: string
): string {
	/* EX:
				data: [
				 {
				   name: 'Project Status',
				   labels: ['Red', 'Amber', 'Green', 'Unknown'],
				   values: [10, 20, 38, 2]
				 }
				]
            */
	const optsChartData = data[0]
	if (!optsChartData) return ''
	const chartColors = resolveChartPalette(opts)
	const labels = firstLabelGroup(optsChartData)

	// The series' own shape props are a placeholder — every slice overrides them in its `<c:dPt>`.
	const spPr = el(
		'c:spPr',
		null,
		[
			raw(el('a:solidFill', null, raw(voidEl('a:schemeClr', { val: 'accent1' })), { openPrefix: '    ' })),
			raw(
				el(
					'a:ln',
					{ w: 9525, cap: 'flat' },
					[
						raw(el('a:solidFill', null, raw(voidEl('a:srgbClr', { val: 'F9F9F9' })))),
						raw(voidEl('a:prstDash', { val: 'solid' })),
						raw(voidEl('a:round')),
					],
					{ openPrefix: '    ' }
				)
			),
			raw(opts.dataNoEffects ? voidEl('a:effectLst') : createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)),
		],
		{ openPrefix: '  ', closePrefix: '  ' }
	)

	// The plot-level `<c:dLbls>` carries the defaults; the per-point ones above it carry the overrides.
	const dLbls = el('c:dLbls', null, [
		raw(labels.map((_label, idx) => pieDataLabel(idx, optsChartData.customLabels?.[idx], opts, chartType)).join('')),
		raw(
			voidEl(
				'c:numFmt',
				{ formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 },
				{ openPrefix: ' ' }
			)
		),
		raw(
			el(
				'c:txPr',
				null,
				[
					raw(voidEl('a:bodyPr', null, { openPrefix: '      ' })),
					raw(voidEl('a:lstStyle', null, { openPrefix: '      ' })),
					raw(
						el(
							'a:p',
							null,
							raw(
								el(
									'a:pPr',
									null,
									raw(labelDefRPr(opts, '            ', { openPrefix: '          ', closePrefix: '          ' })),
									{ openPrefix: '        ', closePrefix: '        ' }
								)
							),
							{ openPrefix: '      ', closePrefix: '      ' }
						)
					),
				],
				{ openPrefix: '    ', closePrefix: '    ' }
			)
		),
		chartType === ChartType.pie ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition || 'ctr' })) : null,
		raw(voidEl('c:showLegendKey', { val: 0 }, { openPrefix: '    ' })),
		raw(voidEl('c:showVal', { val: 0 }, { openPrefix: '    ' })),
		raw(voidEl('c:showCatName', { val: 1 }, { openPrefix: '    ' })),
		raw(voidEl('c:showSerName', { val: 0 }, { openPrefix: '    ' })),
		raw(voidEl('c:showPercent', { val: 1 }, { openPrefix: '    ' })),
		raw(voidEl('c:showBubbleSize', { val: 0 }, { openPrefix: '    ' })),
		raw(voidEl('c:showLeaderLines', { val: opts.showLeaderLines ? 1 : 0 }, { openPrefix: ' ' })),
		raw(createLeaderLinesElement(opts)),
	])

	const ser = el(
		'c:ser',
		null,
		[
			raw(voidEl('c:idx', { val: 0 }, { openPrefix: '  ' })),
			raw(voidEl('c:order', { val: 0 }, { openPrefix: '  ' })),
			raw(strRefBlock('Sheet1!$B$1', optsChartData.name ?? '', 'expanded')),
			raw(spPr),
			raw(
				labels
					.map((_label, idx) =>
						pieDataPoint(idx, optsChartData.pointStyles?.[idx], opts, paletteColor(chartColors, idx))
					)
					.join('')
			),
			raw(dLbls),
			raw(pieCategories(labels)),
			raw(pieValues(optsChartData, labels.length, valFmtCode)),
		],
		{ closePrefix: '  ' }
	)

	return el(`c:${chartType}Chart`, null, [
		raw(voidEl('c:varyColors', { val: 1 }, { openPrefix: '  ' })),
		raw(ser),
		raw(
			voidEl('c:firstSliceAng', { val: opts.firstSliceAng ? Math.round(opts.firstSliceAng) : 0 }, { openPrefix: '  ' })
		),
		chartType === ChartType.doughnut
			? raw(voidEl('c:holeSize', { val: typeof opts.holeSize === 'number' ? opts.holeSize : 50 }))
			: null,
	])
}
