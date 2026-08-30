/**
 * ts-pptx: Scatter Plot Assembly
 *
 * Emits the `<c:scatterChart>` plot element. Scatter is the one family whose first data
 * row supplies X *values* rather than categories, so each `<c:ser>` carries an
 * `<c:xVal>`/`<c:yVal>` pair instead of `<c:cat>`/`<c:val>` -- which is why it does not
 * share the category-axis builder. Reached through {@link ./chart-xml}'s `makeChartType`
 * dispatch.
 */

import { BARCHART_COLORS, DEF_SHAPE_SHADOW } from '../../constants-internal.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { getUuid } from '../utils.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { createLineCap } from '../drawingml/line.js'
import { percentToFixedPercent, ptsToEmuLenient } from '../../units-internal.js'
import { dataLabels, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import {
	chartColorLineFill,
	chartDataLabels,
	dLblShowFlags,
	labelFontAttrs,
	labelFontChildren,
	makeChartErrorBarsXml,
	makeSeriesDataPointsXml,
	numRefBlock,
	paletteColor,
	resolveChartPalette,
	strRefBlock,
	type PlotBuilder,
} from './chart-parts.js'

/**
 * The `(x, y)` runs appended to a `customXY` label: a literal `" ("`, an `XVALUE` field, a
 * `", "`, a `YVALUE` field and a `")"`.
 *
 * Each `<a:fld>` id is minted per build. A field id has to be unique, which is a property a
 * derived id would have to reproduce without an oracle for how far that uniqueness has to
 * reach — so the ids stay random and the *comparison* gives: `NORMALIZERS` in
 * `scripts/pptx-parts.mjs` erases exactly these two field types before the byte-identity gate
 * diffs a part. Do not "fix" the nondeterminism here.
 */
function customXYRuns(obj: OptsChartDataInternal, opts: ChartOptsInternal): string {
	const lang = opts.lang || 'en-US'
	const literal = (text: string): string =>
		el('a:r', null, [raw(voidEl('a:rPr', { lang, baseline: 0, dirty: 0 })), raw(el('a:t', null, text))])
	const field = (type: 'XVALUE' | 'YVALUE', text: string): string =>
		el('a:fld', { id: `{${getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')}}`, type }, [
			raw(voidEl('a:rPr', { lang, baseline: 0 })),
			raw(el('a:pPr', null, raw(voidEl('a:defRPr', null)))),
			raw(el('a:t', null, text)),
		])
	return (
		literal(' (') +
		field('XVALUE', '[' + (obj.name ?? '')) +
		literal(', ') +
		field('YVALUE', '[' + (obj.name ?? '') + ']') +
		literal(')') +
		voidEl('a:endParaRPr', { lang, dirty: 0 })
	)
}

/**
 * One per-point `<c:dLbl>` for the `custom` / `customXY` label formats: rich text carrying the
 * caller's label, optionally followed by the X/Y value fields.
 */
function scatterCustomLabel(
	obj: OptsChartDataInternal,
	opts: ChartOptsInternal,
	label: string,
	idx: number,
	chartUuid: string
): string {
	const rich = el('c:rich', null, [
		raw(el('a:bodyPr', null, raw(voidEl('a:spAutoFit', null)))),
		raw(voidEl('a:lstStyle', null)),
		raw(
			el('a:p', null, [
				raw(el('a:pPr', null, raw(el('a:defRPr', labelFontAttrs(opts), labelFontChildren(opts))))),
				raw(
					el('a:r', null, [
						raw(
							el('a:rPr', { lang: opts.lang || 'en-US', ...labelFontAttrs(opts), dirty: 0 }, labelFontChildren(opts))
						),
						raw(el('a:t', null, label)),
					])
				),
				// The X/Y values are appended only for a label that is not blank or all spaces,
				// which is what lets a caller label a subset of the points.
				opts.dataLabelFormatScatter === 'customXY' && !/^ *$/.test(label) ? raw(customXYRuns(obj, opts)) : null,
			])
		),
	])
	const spPr = el('c:spPr', null, [
		raw(voidEl('a:noFill', null)),
		raw(el('a:ln', null, raw(voidEl('a:noFill', null)))),
		raw(voidEl('a:effectLst', null)),
	])
	const extLst = el('c:extLst', null, [
		raw(
			voidEl('c:ext', {
				uri: '{CE6537A1-D6FC-4f65-9D91-7224C49458BB}',
				'xmlns:c15': 'http://schemas.microsoft.com/office/drawing/2012/chart',
			})
		),
		raw(
			el(
				'c:ext',
				{
					uri: '{C3380CC4-5D6E-409C-BE32-E72D297353CC}',
					'xmlns:c16': 'http://schemas.microsoft.com/office/drawing/2014/chart',
				},
				raw(voidEl('c16:uniqueId', { val: `{${String(idx + 1).padStart(8, '0')}${chartUuid}}` }))
			)
		),
	])
	return el('c:dLbl', null, [
		raw(voidEl('c:idx', { val: idx })),
		raw(el('c:tx', null, raw(rich))),
		raw(spPr),
		opts.dataLabelPosition ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition })) : null,
		...dLblShowFlags({}),
		raw(voidEl('c:showLeaderLines', { val: 1 })),
		raw(extLst),
	])
}

/** The single chart-level `<c:dLbls>` of the `XY` label format: PowerPoint composes the text. */
function scatterXYLabels(opts: ChartOptsInternal): string {
	const spPr = el('c:spPr', null, [
		raw(voidEl('a:noFill', null)),
		raw(el('a:ln', null, raw(voidEl('a:noFill', null)))),
		raw(voidEl('a:effectLst', null)),
	])
	const txPr = el('c:txPr', null, [
		raw(el('a:bodyPr', null, raw(voidEl('a:spAutoFit', null)))),
		raw(voidEl('a:lstStyle', null)),
		raw(
			el('a:p', null, [
				raw(el('a:pPr', null, raw(el('a:defRPr', labelFontAttrs(opts), labelFontChildren(opts))))),
				raw(voidEl('a:endParaRPr', { lang: opts.lang || 'en-US' })),
			])
		),
	])
	const extLst = el(
		'c:extLst',
		null,
		raw(
			el(
				'c:ext',
				{
					uri: '{CE6537A1-D6FC-4f65-9D91-7224C49458BB}',
					'xmlns:c15': 'http://schemas.microsoft.com/office/drawing/2012/chart',
				},
				raw(voidEl('c15:showLeaderLines', { val: 1 }))
			)
		)
	)
	return el('c:dLbls', null, [
		raw(spPr),
		raw(txPr),
		opts.dataLabelPosition ? raw(voidEl('c:dLblPos', { val: opts.dataLabelPosition })) : null,
		...dLblShowFlags({
			val: opts.showLabel ? 1 : 0,
			catName: opts.showLabel ? 1 : 0,
			serName: opts.showSerName ? 1 : 0,
		}),
		raw(extLst),
	])
}

/** Fill, outline and shadow for one scatter series. */
function scatterSerShapeProps(opts: ChartOptsInternal, serColor: string, serIndex: number): string {
	const fill =
		serColor === 'transparent'
			? voidEl('a:noFill')
			: opts.chartColorsOpacity
				? el(
						'a:solidFill',
						null,
						raw(
							createColorElement(
								serColor,
								voidEl('a:alpha', {
									val: percentToFixedPercent(
										opts.chartColorsOpacity,
										'chart/option-out-of-range',
										'chartColorsOpacity'
									),
								})
							)
						)
					)
				: genXmlColorSelection(serColor)
	const line =
		opts.lineSize === 0
			? el('a:ln', null, raw(voidEl('a:noFill')))
			: el('a:ln', { w: ptsToEmuLenient(opts.lineSize ?? 2), cap: createLineCap(opts.lineCap) }, [
					raw(chartColorLineFill(serColor)),
					raw(voidEl('a:prstDash', { val: opts.lineDashValues?.[serIndex] ?? opts.lineDash ?? 'solid' })),
					raw(voidEl('a:round')),
				])
	return el('c:spPr', null, [raw(fill), raw(line), raw(createShadowEffectLst(opts.shadow, DEF_SHAPE_SHADOW))])
}

/** The point marker: symbol, optional size, and its own fill and outline. */
function scatterMarker(opts: ChartOptsInternal, markerColor: string): string {
	const spPr = el('c:spPr', null, [
		raw(markerColor === 'transparent' ? voidEl('a:noFill') : genXmlColorSelection(markerColor)),
		raw(
			el('a:ln', { w: opts.lineDataSymbolLineSize, cap: 'flat' }, [
				raw(chartColorLineFill(opts.lineDataSymbolLineColor || markerColor)),
				raw(voidEl('a:prstDash', { val: 'solid' })),
				raw(voidEl('a:round')),
			])
		),
		raw(voidEl('a:effectLst')),
	])
	return el('c:marker', null, [
		raw(voidEl('c:symbol', { val: opts.lineDataSymbol })),
		// Defaults to "auto" otherwise (but this is usually too small, so there is a default).
		opts.lineDataSymbolSize ? raw(voidEl('c:size', { val: opts.lineDataSymbolSize })) : null,
		raw(spPr),
	])
}

/**
 * Plot an XY scatter chart into `<c:scatterChart>` (paired X/Y numeric series).
 */
export const makeScatterPlot: PlotBuilder = (chartType, data, opts, valAxisId, catAxisId, valFmtCode) => {
	/*
				`data` = [
					{ name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Value 1', values:[13, 20, 21, 25] },
					{ name:'Y-Value 2', values:[ 1,  2,  5,  9] }
				];
            */
	const chartColors = resolveChartPalette(opts)
	// X values come from the first row; each later row is one Y series.
	const xValues = dataValues(data[0])
	// Legacy single-series colour-vary, and it can never fire: a scatter `data` array is one row of
	// X values plus one row per Y series, the loop below runs over `data.slice(1)`, so by the time
	// `data.length === 1` there are no series left to colour. The intent was presumably "one Y
	// series" (`=== 2`). Fixing it would start emitting `<c:dPt>` colour-vary for single-series
	// scatter charts, which is an output change and a question about what such a chart should look
	// like — so the condition is left exactly as it has always been, said out loud rather than
	// quietly corrected.
	const scatterVaryColors =
		data.length === 1 && opts.chartColors !== BARCHART_COLORS ? opts.chartColors || BARCHART_COLORS : null

	const sers = data
		.slice(1)
		.map((obj, idx) => {
			const serColor = paletteColor(chartColors, idx)
			const yValues = dataValues(obj)
			// The Y series is cached against the X series' length, so a caller who supplied fewer Y
			// values than X leaves gaps rather than a short cache.
			const values =
				numRefBlock('c:xVal', `Sheet1!$A$2:$A$${xValues.length + 1}`, valFmtCode, xValues) +
				numRefBlock(
					'c:yVal',
					sheetRangeRef(idx + 2, 2, idx + 2, xValues.length + 1),
					valFmtCode,
					xValues.map((_value, i) => yValues[i])
				)

			// Scatter data point labels. `chartUuid` tails each point's `c16:uniqueId` and is minted
			// per build, for the reason {@link customXYRuns} gives about the `a:fld` ids.
			let labels = ''
			if (opts.showLabel) {
				const chartUuid = getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
				const isCustom = opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY'
				if (dataLabels(obj)[0] && isCustom) {
					labels += el(
						'c:dLbls',
						null,
						raw(
							firstLabelGroup(obj)
								.map((label, pointIdx) => scatterCustomLabel(obj, opts, label, pointIdx, chartUuid))
								.join('')
						)
					)
				}
				if (opts.dataLabelFormatScatter === 'XY') labels += scatterXYLabels(opts)
			}

			return el('c:ser', null, [
				raw(voidEl('c:idx', { val: idx })),
				raw(voidEl('c:order', { val: idx })),
				raw(strRefBlock(sheetCellRef(idx + 2, 1), obj.name ?? '')),
				raw(scatterSerShapeProps(opts, serColor, idx)),
				raw(scatterMarker(opts, serColor)),
				// Per-point data points (`c:dPt`) MUST precede `c:dLbls` (CT_ScatterSer schema order).
				raw(makeSeriesDataPointsXml(chartType, obj, opts, scatterVaryColors)),
				raw(labels),
				// Error bars come after dLbls and before xVal/yVal in schema order.
				raw(makeChartErrorBarsXml(chartType, obj.errorBars, obj)),
				raw(values),
				raw(voidEl('c:smooth', { val: opts.lineSmooth ? 1 : 0 })),
			])
		})
		.join('')

	return el(`c:${chartType}Chart`, null, [
		raw(voidEl('c:scatterStyle', { val: 'lineMarker' })),
		raw(voidEl('c:varyColors', { val: 0 })),
		raw(sers),
		raw(chartDataLabels(opts, false)),
		// Axis id order matters: category comes first.
		raw(voidEl('c:axId', { val: catAxisId })),
		raw(voidEl('c:axId', { val: valAxisId })),
	])
}
