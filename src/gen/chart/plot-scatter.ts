/**
 * ts-pptx: Scatter Plot Assembly
 *
 * Emits the `<c:scatterChart>` plot element. Scatter is the one family whose first data
 * row supplies X *values* rather than categories, so each `<c:ser>` carries an
 * `<c:xVal>`/`<c:yVal>` pair instead of `<c:cat>`/`<c:val>` -- which is why it does not
 * share the category-axis builder. Reached through {@link ./chart-xml}'s `makeChartType`
 * dispatch.
 */

import { BARCHART_COLORS } from '../../constants-internal.js'
import type { ChartSeriesOpts } from '../../types/index.js'
import type { ChartOptsInternal, OptsChartDataInternal } from '../../types/internal.js'
import { getUuid } from '../utils.js'
import { dataLabels, firstLabelGroup } from './data-refs.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { xsdBool } from '../../ooxml/xsd-boolean.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import {
	c15LeaderLinesExt,
	chartDataLabels,
	chartLang,
	dLblPosEl,
	dLblsBlock,
	dLblShowFlags,
	labelDefRPr,
	labelRun,
	makeChartErrorBarsXml,
	makeSeriesDataPointsXml,
	type PlotBuilder,
	seriesNameRef,
	seriesShapeProps,
	seriesStroke,
	serMarker,
	xySeriesRefs,
} from './chart-parts.js'

/**
 * The `<c:spPr>` a scatter data label carries: no fill, no outline, no effects — the label is
 * text on the plot rather than a box. Both label builders in this file spelled it out.
 */
const TRANSPARENT_LABEL_SPPR = el('c:spPr', null, [
	raw(voidEl('a:noFill', null)),
	raw(el('a:ln', null, raw(voidEl('a:noFill', null)))),
	raw(voidEl('a:effectLst', null)),
])

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
	const lang = chartLang(opts)
	const literal = (text: string): string =>
		el('a:r', null, [raw(voidEl('a:rPr', { lang, baseline: 0, dirty: 0 })), raw(el('a:t', null, text))])
	// Upper-cased: `a:fld/@id` is `ST_Guid`, whose pattern is `[0-9A-F]` only, so a lower-case
	// nibble is a schema error the validator reports at every field. `getUuid` emits lower case.
	const field = (type: 'XVALUE' | 'YVALUE', text: string): string =>
		el('a:fld', { id: `{${getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx').toUpperCase()}}`, type }, [
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
	chartUuid: string,
	over?: ChartSeriesOpts
): string {
	const rich = el('c:rich', null, [
		raw(el('a:bodyPr', null, raw(voidEl('a:spAutoFit', null)))),
		raw(voidEl('a:lstStyle', null)),
		raw(
			el('a:p', null, [
				raw(el('a:pPr', null, raw(labelDefRPr(opts, over)))),
				raw(labelRun(opts, label, over)),
				// The X/Y values are appended only for a label that is not blank or all spaces,
				// which is what lets a caller label a subset of the points.
				opts.dataLabelFormatScatter === 'customXY' && !/^ *$/.test(label) ? raw(customXYRuns(obj, opts)) : null,
			])
		),
	])
	const extLst = el('c:extLst', null, [
		raw(
			voidEl('c:ext', {
				uri: '{CE6537A1-D6FC-4f65-9D91-7224C49458BB}',
				'xmlns:c15': OOXML_NS.c15,
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
	return dLblsBlock(
		{
			lead: voidEl('c:idx', { val: idx }) + el('c:tx', null, raw(rich)),
			spPr: TRANSPARENT_LABEL_SPPR,
			dLblPos: dLblPosEl(opts),
			flags: dLblShowFlags({}),
			// No `c:showLeaderLines` here. It used to be written hard-coded on, and `CT_DLbl` has no
			// such child -- it belongs to `CT_DLbls`, the plural container -- so every scatter with
			// `showLabel` and a `custom` or `customXY` format failed the schema validator at every
			// point. Desktop PowerPoint opened it anyway, and no schema case covered scatter custom
			// labels, so it went unseen.
			//
			// It was kept on the reasoning that a moved custom label with no leader line is a
			// different chart. A COM probe settles that: drag a scatter's data label away from its
			// point in PowerPoint and the `c:dLbl` it writes carries `c:idx`, a
			// `c:layout/c:manualLayout` holding the drag offset, `c:tx`, the six show flags and an
			// `extLst` of `c15:showDataLabelsRange` and `c16:uniqueId`. No leader line, in any
			// spelling, and none on the enclosing `c:dLbls` either. PowerPoint records nothing there,
			// so neither do we.
			extLst,
		},
		'c:dLbl'
	)
}

/** The per-series `<c:dLbls>` of the `XY` label format: PowerPoint composes the text. */
function scatterXYLabels(opts: ChartOptsInternal, over?: ChartSeriesOpts): string {
	const txPr = el('c:txPr', null, [
		raw(el('a:bodyPr', null, raw(voidEl('a:spAutoFit', null)))),
		raw(voidEl('a:lstStyle', null)),
		raw(
			el('a:p', null, [
				raw(el('a:pPr', null, raw(labelDefRPr(opts, over)))),
				raw(voidEl('a:endParaRPr', { lang: chartLang(opts) })),
			])
		),
	])
	const extLst = c15LeaderLinesExt(1)
	return dLblsBlock({
		spPr: TRANSPARENT_LABEL_SPPR,
		txPr,
		dLblPos: dLblPosEl(opts),
		flags: dLblShowFlags({
			val: xsdBool(opts.showLabel),
			catName: xsdBool(opts.showLabel),
			serName: xsdBool(opts.showSerName),
		}),
		extLst,
	})
}

/** Fill, outline and shadow for one scatter series. */
function scatterSerShapeProps(opts: ChartOptsInternal, serColor: string, serIndex: number, lineSize?: number): string {
	return seriesShapeProps(opts, serColor, seriesStroke(opts, serColor, serIndex, lineSize))
}

/**
 * Plot an XY scatter chart into `<c:scatterChart>` (paired X/Y numeric series).
 */
export const makeScatterPlot: PlotBuilder = (chartType, data, opts, valAxisId, catAxisId, valFmtCode, sheet) => {
	/*
				`data` = [
					{ name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Value 1', values:[13, 20, 21, 25] },
					{ name:'Y-Value 2', values:[ 1,  2,  5,  9] }
				];
            */
	// X values come from the first row; each later row is one Y series. Every column comes from the
	// worksheet layout: a standalone scatter's sheet puts X in column A, and a combo's sheet puts it
	// in the X row's own column after the label columns.
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
		.map((obj) => {
			// `seriesOptions[0]` styles the FIRST Y series, `data[1]`: `data[0]` is the shared X row and is
			// not a series. In a combo the index stays clear of every other subchart's, where a count
			// restarting at 0 gave a bar and a scatter the same `<c:idx>` and the same colour.
			const { idx, over, color: serColor, xVal, yVal } = xySeriesRefs(chartType, obj, data, opts, sheet, valFmtCode)

			// Scatter data point labels. `chartUuid` tails each point's `c16:uniqueId` and is minted
			// per build, for the reason {@link customXYRuns} gives about the `a:fld` ids.
			let labels = ''
			if (opts.showLabel) {
				// Upper-cased for the same reason as the `a:fld` ids above, and to match what
				// PowerPoint writes: `{00000001-10AD-4F37-A175-DD7655EB2B6B}`.
				const chartUuid = getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx').toUpperCase()
				const isCustom = opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY'
				if (dataLabels(obj)[0] && isCustom) {
					labels += el(
						'c:dLbls',
						null,
						raw(
							firstLabelGroup(obj)
								.map((label, pointIdx) => scatterCustomLabel(obj, opts, label, pointIdx, chartUuid, over))
								.join('')
						)
					)
				}
				if (opts.dataLabelFormatScatter === 'XY') labels += scatterXYLabels(opts, over)
			}

			return el('c:ser', null, [
				raw(voidEl('c:idx', { val: idx })),
				raw(voidEl('c:order', { val: idx })),
				raw(seriesNameRef(obj, sheet)),
				raw(scatterSerShapeProps(opts, serColor, idx, over?.lineSize)),
				raw(serMarker(opts, serColor, serColor)),
				// Per-point data points (`c:dPt`) MUST precede `c:dLbls` (CT_ScatterSer schema order).
				raw(makeSeriesDataPointsXml(chartType, obj, opts, scatterVaryColors)),
				raw(labels),
				// Error bars come after dLbls and before xVal/yVal in schema order.
				raw(makeChartErrorBarsXml(chartType, obj.errorBars, obj)),
				raw(xVal + yVal),
				raw(voidEl('c:smooth', { val: xsdBool(opts.lineSmooth) })),
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
