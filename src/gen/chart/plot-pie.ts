/**
 * ts-pptx: Pie & Doughnut Plot Assembly
 *
 * Emits the `<c:pieChart>` / `<c:doughnutChart>` plot elements. These are the only
 * families with no axes at all -- a single series, one `<c:dPt>` per slice carrying its
 * own fill, and optional leader lines -- so the builder takes no axis ids. Reached
 * through {@link ./chart-xml}'s `makeChartType` dispatch.
 */

import { ChartType } from '../../enums.js'
import { DEF_SHAPE_SHADOW } from '../../constants-internal.js'
import type { ChartDataPointStyle } from '../../types/chart.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { dataValues, firstLabelGroup } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	createChartBorderLine,
	createDataBorderLine,
	createLeaderLinesElement,
	dLblShowFlags,
	labelFontAttrs,
	labelFontChildren,
	paletteColor,
	resolveChartPalette,
	strRefBlock,
} from './chart-parts.js'

/**
 * The label run properties both `<c:txPr>` spellings share: {@link labelFontAttrs} in the
 * `sz, b, i, u, strike` ordering, over {@link labelFontChildren}'s colour and typeface.
 */
function labelDefRPr(opts: ChartOptsInternal): string {
	return el('a:defRPr', labelFontAttrs(opts), labelFontChildren(opts))
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
			? createDataBorderLine(opts.dataBorder, 'flat')
			: ''
	const spPr = el('c:spPr', null, [
		raw(el('a:solidFill', null, raw(createColorElement(ptStyle?.fill || fallbackColor)))),
		raw(border),
		raw(createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)),
	])
	return el('c:dPt', null, [raw(voidEl('c:idx', { val: idx })), raw(voidEl('c:bubble3D', { val: 0 })), raw(spPr)])
}

/**
 * One `<c:dLbl>`: a pie labels its points, not its series, so every slice carries its own copy of
 * the number format, text style and show flags. A `customLabels` entry replaces the value with
 * literal rich text, which is why it also forces `<c:showVal>` off.
 */
function pieDataLabel(idx: number, customLbl: string | undefined, opts: ChartOptsInternal, chartType: ChartType) {
	const numFmt = voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })
	const txPr = el('c:txPr', null, [
		raw(voidEl('a:bodyPr', null)),
		raw(voidEl('a:lstStyle')),
		raw(el('a:p', null, raw(el('a:pPr', null, raw(labelDefRPr(opts)))))),
	])
	return el('c:dLbl', null, [
		raw(voidEl('c:idx', { val: idx })),
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
		raw(voidEl('c:spPr', null)),
		raw(txPr),
		chartType === ChartType.pie && opts.dataLabelPosition
			? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition }))
			: null,
		...dLblShowFlags({
			val: customLbl ? 0 : opts.showValue ? 1 : 0,
			catName: opts.showLabel ? 1 : 0,
			serName: opts.showSerName ? 1 : 0,
			percent: opts.showPercent ? 1 : 0,
		}),
	])
}

/** The `<c:cat>` slice-name reference and the `<c:val>` cache, both keyed on the label count. */
function pieCategories(labels: string[]): string {
	const strCache = el('c:strCache', null, [
		raw(voidEl('c:ptCount', { val: labels.length })),
		raw(labels.map((label, idx) => el('c:pt', { idx }, raw(el('c:v', null, label)))).join('')),
	])
	const strRef = el('c:strRef', null, [raw(el('c:f', null, `Sheet1!$A$2:$A$${labels.length + 1}`)), raw(strCache)])
	return el('c:cat', null, raw(strRef))
}

/** The `<c:val>` numeric cache. A missing value keeps its `<c:pt>` with an empty `<c:v>`. */
function pieValues(obj: OptsChartDataInternal, count: number, valFmtCode: string): string {
	const points = dataValues(obj)
		.map((value, idx) => el('c:pt', { idx }, raw(el('c:v', null, value || value === 0 ? value : ''))))
		.join('')
	const numCache = el('c:numCache', null, [
		raw(el('c:formatCode', null, valFmtCode)),
		raw(voidEl('c:ptCount', { val: count })),
		raw(points),
	])
	const numRef = el('c:numRef', null, [raw(el('c:f', null, `Sheet1!$B$2:$B$${count + 1}`)), raw(numCache)])
	return el('c:val', null, raw(numRef))
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
	const spPr = el('c:spPr', null, [
		raw(el('a:solidFill', null, raw(voidEl('a:schemeClr', { val: 'accent1' })))),
		raw(
			el('a:ln', { w: 9525, cap: 'flat' }, [
				raw(el('a:solidFill', null, raw(voidEl('a:srgbClr', { val: 'F9F9F9' })))),
				raw(voidEl('a:prstDash', { val: 'solid' })),
				raw(voidEl('a:round')),
			])
		),
		raw(opts.dataNoEffects ? voidEl('a:effectLst') : createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW)),
	])

	// The plot-level `<c:dLbls>` carries the defaults; the per-point ones above it carry the overrides.
	const dLbls = el('c:dLbls', null, [
		raw(labels.map((_label, idx) => pieDataLabel(idx, optsChartData.customLabels?.[idx], opts, chartType)).join('')),
		raw(voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })),
		raw(
			el('c:txPr', null, [
				raw(voidEl('a:bodyPr', null)),
				raw(voidEl('a:lstStyle', null)),
				raw(el('a:p', null, raw(el('a:pPr', null, raw(labelDefRPr(opts)))))),
			])
		),
		chartType === ChartType.pie ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition || 'ctr' })) : null,
		...dLblShowFlags({ catName: 1, percent: 1 }),
		raw(voidEl('c:showLeaderLines', { val: opts.showLeaderLines ? 1 : 0 })),
		raw(createLeaderLinesElement(opts)),
	])

	const ser = el('c:ser', null, [
		raw(voidEl('c:idx', { val: 0 })),
		raw(voidEl('c:order', { val: 0 })),
		raw(strRefBlock('Sheet1!$B$1', optsChartData.name ?? '')),
		raw(spPr),
		raw(
			labels
				.map((_label, idx) => pieDataPoint(idx, optsChartData.pointStyles?.[idx], opts, paletteColor(chartColors, idx)))
				.join('')
		),
		raw(dLbls),
		raw(pieCategories(labels)),
		raw(pieValues(optsChartData, labels.length, valFmtCode)),
	])

	return el(`c:${chartType}Chart`, null, [
		raw(voidEl('c:varyColors', { val: 1 })),
		raw(ser),
		raw(voidEl('c:firstSliceAng', { val: opts.firstSliceAng ? Math.round(opts.firstSliceAng) : 0 })),
		chartType === ChartType.doughnut
			? raw(voidEl('c:holeSize', { val: typeof opts.holeSize === 'number' ? opts.holeSize : 50 }))
			: null,
	])
}
