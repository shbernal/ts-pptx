/**
 * ts-pptx: chartEx (`cx:`) Chart-Data Builder
 *
 * Builds the `<cx:chartData>` block of a chartEx chart part (waterfall, and the other
 * Office-2016 layouts as they land). This is the chartEx analogue of the `<c:cat>`/`<c:val>`
 * caches on the classic side: it names the embedded-workbook cells via `<cx:f>` formulas and
 * mirrors their values into `<cx:lvl>` string/number caches so PowerPoint can render without
 * opening the workbook. The category label levels occupy the leading columns (0 of them for a
 * category-less histogram, 1 for a flat chart, N for a hierarchical treemap/sunburst) and the one
 * value series sits in the next column — so the references here line up 1:1 with the workbook
 * written by {@link ./embed-xlsx}.
 */

import { ChartType } from '../../core-enums.js'
import type { SlideRelChart } from '../../types/internal.js'
import { warn } from '../../log.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { dataLabels, dataValues, firstLabelGroup, sheetCellRef, sheetRangeRef } from './data-refs.js'

/**
 * The numeric-dimension tag PowerPoint expects varies by layout. The flat layouts (waterfall,
 * funnel, histogram, pareto, boxWhisker) use `type="val"`; the hierarchical treemap/sunburst use
 * `type="size"`; a `regionMap` uses `type="colorVal"` (the value drives each region's fill color).
 * The value always comes from the series' `values`; only the dimension tag differs.
 */
function valueDimType(type: ChartType): 'size' | 'val' | 'colorVal' {
	if (type === ChartType.treemap || type === ChartType.sunburst) return 'size'
	if (type === ChartType.regionMap) return 'colorVal'
	return 'val'
}

/**
 * Build the `<cx:chartData>` for a single-series chartEx chart.
 *
 * `<cx:data id="0">` carries a `<cx:strDim type="cat">` (category labels) and a numeric dimension
 * (`type="val"` for flat layouts, `type="size"` for the hierarchical treemap/sunburst);
 * `<cx:externalData r:id="rId1">` points at the embedded workbook (same `rId1` the chart part's
 * `.rels` maps to the `.xlsx`).
 *
 * Category labels are `string[][]` — one inner array per hierarchy level (leaf first). A flat chart
 * has a single level; treemap/sunburst carry several. Each level becomes one `<cx:lvl>`, emitted
 * leaf-first to match PowerPoint. The workbook ({@link ./embed-xlsx}) lays the levels out in
 * columns A..N (outermost in A, leaf in the last label column) with the value series in the next
 * column; the single `<cx:f>` range on each dimension spans exactly those columns.
 * @param {SlideRelChart} rel - the registered chart
 * @return {string} `<cx:chartData>…</cx:chartData>` XML
 */
export function makeChartExData(rel: SlideRelChart): string {
	const series = rel.data[0]
	const type = rel.opts._type as ChartType
	const levels = dataLabels(series)
	const totLvl = levels.length // 0 for a category-less layout (histogram), 1 flat, N hierarchical
	const vals = dataValues(series)
	// Row span is driven by whichever dimension has more points; the workbook writes both to the same rows.
	const ptCount = Math.max(firstLabelGroup(series).length, vals.length)
	const lastRow = ptCount + 1
	const valueCol = totLvl + 1 // value series sits in the column right after all label columns

	// Category (string) dimension — one <cx:lvl> per hierarchy level, leaf-first, over columns A..totLvl.
	// A category-less layout (histogram bins raw observations itself) emits no strDim at all.
	let catDim = ''
	if (totLvl > 0) {
		const catLvls = levels
			.map((group) => {
				const pts = group.map((label, idx) => el('cx:pt', { idx }, label ?? '')).join('')
				return el('cx:lvl', { ptCount: group.length }, raw(pts))
			})
			.join('')
		catDim = el('cx:strDim', { type: 'cat' }, [
			raw(el('cx:f', null, sheetRangeRef(1, 2, totLvl, lastRow))),
			raw(catLvls),
		])
	}

	// Numeric dimension — the value column. Skip null/non-finite points (a valid sparse cache);
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
	const valDim = el('cx:numDim', { type: valueDimType(type) }, [
		raw(el('cx:f', null, sheetRangeRef(valueCol, 2, valueCol, lastRow))),
		raw(el('cx:lvl', { ptCount, formatCode: 'General' }, raw(valPts))),
	])

	const data = el('cx:data', { id: 0 }, [raw(catDim), raw(valDim)])
	// CT_ChartData order is externalData? → data+. externalData is a LEAF whose only attribute is
	// `id` (r:id → the embedded workbook); it carries no `autoUpdate` and no child elements.
	const externalData = voidEl('cx:externalData', { 'r:id': 'rId1' })
	return el('cx:chartData', null, [raw(externalData), raw(data)])
}

/**
 * The `Sheet1` cell holding the series name (the numeric dimension's header) — column immediately
 * after all label columns, row 1. A category-less histogram lands on `$A$1`; a flat chart on
 * `$B$1`; a 3-level treemap on `$D$1`.
 * @param {SlideRelChart} rel - the registered chart
 * @return {string} an absolute single-cell reference, e.g. `Sheet1!$B$1`
 */
export function chartExSeriesNameRef(rel: SlideRelChart): string {
	return sheetCellRef(dataLabels(rel.data[0]).length + 1, 1)
}
