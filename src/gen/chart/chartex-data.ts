/**
 * PptxGenJS: chartEx (`cx:`) Chart-Data Builder
 *
 * Builds the `<cx:chartData>` block of a chartEx chart part (waterfall, and the other
 * Office-2016 layouts as they land). This is the chartEx analogue of the `<c:cat>`/`<c:val>`
 * caches on the classic side: it names the embedded-workbook cells via `<cx:f>` formulas and
 * mirrors their values into `<cx:lvl>` string/number caches so PowerPoint can render without
 * opening the workbook. The cell layout is identical to the classic single-series cat chart —
 * categories in column A (rows 2..n+1), the one value series in column B (name in B1) — so the
 * references here line up 1:1 with the workbook written by {@link ./embed-xlsx}.
 */

import type { SlideRelChart } from '../../types/internal.js'
import { warn } from '../../log.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { dataValues, firstLabelGroup } from './data-refs.js'

/**
 * Build the `<cx:chartData>` for a single-series chartEx chart.
 *
 * `<cx:data id="0">` carries a `<cx:strDim type="cat">` (category labels) and a
 * `<cx:numDim type="val">` (series values); `<cx:externalData r:id="rId1">` points at the
 * embedded workbook (same `rId1` the chart part's `.rels` maps to the `.xlsx`).
 * @param {SlideRelChart} rel - the registered chart
 * @return {string} `<cx:chartData>…</cx:chartData>` XML
 */
export function makeChartExData(rel: SlideRelChart): string {
	const series = rel.data[0]
	const cats = firstLabelGroup(series)
	const vals = dataValues(series)
	// Row span is driven by whichever axis has more points; the workbook writes both to the same rows.
	const ptCount = Math.max(cats.length, vals.length)
	const lastRow = ptCount + 1

	// Category (string) dimension — column A.
	const catPts = cats.map((label, idx) => el('cx:pt', { idx }, label ?? '')).join('')
	const catDim = el('cx:strDim', { type: 'cat' }, [
		raw(el('cx:f', null, `Sheet1!$A$2:$A$${lastRow}`)),
		raw(el('cx:lvl', { ptCount }, raw(catPts))),
	])

	// Value (number) dimension — column B. Skip null/non-finite points (a valid sparse cache);
	// warn on non-finite, matching the classic side's `numCachePt` policy.
	let valPts = ''
	vals.forEach((value, idx) => {
		if (value == null) return
		if (!Number.isFinite(value)) {
			warn(`chartEx value "${value}" at index ${idx} is not a finite number; data point omitted.`)
			return
		}
		valPts += el('cx:pt', { idx }, value)
	})
	const valDim = el('cx:numDim', { type: 'val' }, [
		raw(el('cx:f', null, `Sheet1!$B$2:$B$${lastRow}`)),
		raw(el('cx:lvl', { ptCount, formatCode: 'General' }, raw(valPts))),
	])

	const data = el('cx:data', { id: 0 }, [raw(catDim), raw(valDim)])
	// CT_ChartData order is externalData? → data+. externalData is a LEAF whose only attribute is
	// `id` (r:id → the embedded workbook); it carries no `autoUpdate` and no child elements.
	const externalData = voidEl('cx:externalData', { 'r:id': 'rId1' })
	return el('cx:chartData', null, [raw(externalData), raw(data)])
}
