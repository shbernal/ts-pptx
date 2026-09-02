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
import { warn } from '../../diagnostics.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { categoryRange, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, raw, voidEl, type XmlChild } from '../oxml/el.js'
import {
	catRefBlock,
	createChartBorderLine,
	createDataBorderLine,
	createLeaderLinesElement,
	dLblShowFlags,
	labelFontAttrs,
	labelFontChildren,
	labelTextProps,
	numRefBlock,
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
	const txPr = labelTextProps(labelDefRPr(opts))
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
		...pieLabelFlags(opts, customLbl),
	])
}

/**
 * The four `<c:show*>` flags a pie's data labels carry.
 *
 * The per-point `<c:dLbl>` read the caller's options; the plot-level `<c:dLbls>` wrote
 * `catName: 1, percent: 1` as constants, so a caller's `showLabel: false` and
 * `showPercent: false` both came back inverted. The constants were masked while every point
 * carried its own `<c:dLbl>` to override them, which holds only while the pie has labels.
 * @param opts - the chart's normalized options
 * @param customLbl - this point's literal label text, which replaces the value and so forces
 *   `<c:showVal>` off. Absent at the plot level, which states no text of its own.
 */
function pieLabelFlags(opts: ChartOptsInternal, customLbl?: string): XmlChild[] {
	return dLblShowFlags({
		val: customLbl ? 0 : opts.showValue ? 1 : 0,
		catName: opts.showLabel ? 1 : 0,
		serName: opts.showSerName ? 1 : 0,
		percent: opts.showPercent ? 1 : 0,
	})
}

/** The `<c:cat>` slice-name reference, keyed on the label count. */
function pieCategories(labels: string[]): string {
	return el('c:cat', null, raw(catRefBlock('str', categoryRange(labels.length), labels)))
}

/**
 * The `<c:val>` numeric cache, keyed on the slice count.
 *
 * This was the last hand-built copy of {@link numRefBlock}, and it was kept out of that merge
 * because it disagreed about a gap rather than merely spelling one differently: it emitted a
 * `<c:pt>` with an empty `<c:v>` for a `null`/`undefined`/`NaN` and `<c:v>Infinity</c:v>` for an
 * infinity, where `numCachePt` leaves the point out. Measured against desktop PowerPoint,
 * the two disagreements are not alike:
 *
 * - The empty `<c:v>` was harmless. A pie with the point present-but-empty and the same pie with
 *   the point absent open without a repair prompt, resolve to the same object model
 *   (`SeriesCollection(1).Values` reads `Empty` at that index in both), and export to a
 *   byte-identical PNG. A sensitivity check confirms the cache is what paints: perturbing a
 *   *neighbouring* cached value moves both the read-back and the pixels.
 * - `<c:v>Infinity</c:v>` was not. PowerPoint refuses the package with 0x80070570 — the
 *   corrupt-file error — and so does `NaN` or `INF` in that position. A pie was therefore the one
 *   family where a non-finite value reached the deck instead of being warned about and dropped.
 */
function pieValues(obj: OptsChartDataInternal, count: number, valFmtCode: string): string {
	return numRefBlock('c:val', sheetRangeRef(2, 2, 2, count + 1), valFmtCode, dataValues(obj), count)
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
	// A pie slices its values, so the values are what decide how many slices there are. Keying
	// everything on `labels.length` meant an unlabelled pie emitted no `<c:dPt>` at all and
	// reversed both sheet ranges (`Sheet1!$A$2:$A$1`), because the end row is `count + 1`.
	const sliceCount = labels.length || dataValues(optsChartData).length
	if (sliceCount === 0) {
		warn('chart/point-count-mismatch', 'addChart: a pie/doughnut series with no values plots nothing; skipping it.')
		return ''
	}

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
		raw(
			Array.from({ length: sliceCount }, (_unused, idx) =>
				pieDataLabel(idx, optsChartData.customLabels?.[idx], opts, chartType)
			).join('')
		),
		raw(voidEl('c:numFmt', { formatCode: (opts.dataLabelFormatCode ?? '') || 'General', sourceLinked: 0 })),
		raw(labelTextProps(labelDefRPr(opts))),
		chartType === ChartType.pie ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition || 'ctr' })) : null,
		...pieLabelFlags(opts),
		raw(voidEl('c:showLeaderLines', { val: opts.showLeaderLines ? 1 : 0 })),
		raw(createLeaderLinesElement(opts)),
	])

	const ser = el('c:ser', null, [
		raw(voidEl('c:idx', { val: 0 })),
		raw(voidEl('c:order', { val: 0 })),
		raw(strRefBlock(sheetCellRef(2, 1), optsChartData.name ?? '')),
		raw(spPr),
		raw(
			Array.from({ length: sliceCount }, (_unused, idx) =>
				pieDataPoint(idx, optsChartData.pointStyles?.[idx], opts, paletteColor(chartColors, idx))
			).join('')
		),
		raw(dLbls),
		// `<c:cat>` is optional on a `CT_PieSer`; an unlabelled pie states no category names
		// rather than referencing an empty range.
		labels.length > 0 ? raw(pieCategories(labels)) : null,
		raw(pieValues(optsChartData, sliceCount, valFmtCode)),
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
