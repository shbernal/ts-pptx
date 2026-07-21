/**
 * PptxGenJS: chartEx (`cx:`) Chart Assembly
 *
 * Builds a chartEx chart part `ppt/charts/chartEx{N}.xml` — the `<cx:chartSpace>` for the
 * Office-2016 chart family (currently `waterfall`). This is the chartEx analogue of
 * {@link ./chart-xml} (`makeXmlCharts`): a pure string builder, no I/O, no model mutation.
 *
 * chartEx differs from the classic `<c:chartSpace>` in three ways the rest of the package must
 * account for (all handled at the call sites, not here):
 * - it is a **separate part** with content type `application/vnd.ms-office.chartex+xml`,
 * - it is referenced from the slide through `<mc:AlternateContent>` (see `gen/slide/object.ts`),
 * - it renders **only in PowerPoint 2016+/Microsoft 365**; other consumers show the fallback shape.
 *
 * The `<cx:f>` formulas point back at the same embedded workbook the classic charts use
 * ({@link ./embed-xlsx}); the cell mapping lives in {@link ./chartex-data} / {@link ./data-refs}.
 */

import { ChartType } from '../../core-enums.js'
import { DEF_FONT_COLOR, XML_DECL } from '../../core-enums-internal.js'
import type { SlideRelChart } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { createChartTextFonts } from './chart-parts.js'
import { makeChartExData } from './chartex-data.js'

const CX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/**
 * Map a chartEx {@link ChartType} to its `<cx:series layoutId>` token (the CT_SeriesLayout value
 * PowerPoint keys the chart geometry on). Returns `''` for a non-chartEx type.
 */
export function chartExLayoutId(type: ChartType): string {
	switch (type) {
		case ChartType.waterfall:
			return 'waterfall'
		default:
			return ''
	}
}

/**
 * A deterministic `uniqueId` GUID for a `<cx:series>`, derived from the chart's package-global id.
 * PowerPoint stamps a random GUID here; deriving it keeps output byte-stable (the project pins
 * chart determinism — see `docs/backlog.yml` `fork-chart-counter-nondeterminism`).
 */
function chartExUniqueId(globalId: number): string {
	return `{00000000-0000-0000-0000-${String(globalId).padStart(12, '0')}}`
}

/** chartEx legend positions are `t|b|l|r`; clamp the classic `'tr'` (and anything else) to `'t'`. */
function chartExLegendPos(pos: string | undefined): string {
	return pos === 'b' || pos === 'l' || pos === 'r' ? pos : 't'
}

/** Build the `<cx:series>` (title cell, data labels, dataId, and layout-specific `<cx:layoutPr>`). */
function makeChartExSeries(rel: SlideRelChart): string {
	const opts = rel.opts
	const type = opts._type as ChartType
	const showValue = !!(opts.showValue || opts.showLabel)

	const tx = el(
		'cx:tx',
		null,
		raw(el('cx:txData', null, [raw(el('cx:f', null, 'Sheet1!$B$1')), raw(el('cx:v', null, rel.data[0]?.name ?? ''))]))
	)

	// Data labels: chartEx toggles each field via <cx:visibility>. Only emitted when the caller
	// asked to show values, matching the classic default of no data labels.
	const dataLabels = showValue
		? el('cx:dataLabels', { pos: 'outEnd' }, raw(voidEl('cx:visibility', { seriesName: 0, categoryName: 0, value: 1 })))
		: ''

	// Layout-specific series props. waterfall: subtotal/total column indices.
	let layoutPr = ''
	if (type === ChartType.waterfall && Array.isArray(opts.subtotals) && opts.subtotals.length > 0) {
		const idxs = opts.subtotals.map((idx) => voidEl('cx:idx', { val: idx })).join('')
		layoutPr = el('cx:layoutPr', null, raw(el('cx:subtotals', null, raw(idxs))))
	}

	return el('cx:series', { layoutId: chartExLayoutId(type), uniqueId: chartExUniqueId(rel.globalId) }, [
		raw(tx),
		raw(dataLabels),
		raw(voidEl('cx:dataId', { val: 0 })),
		raw(layoutPr),
	])
}

/** Build the `<cx:axis>` list for a layout. waterfall has a category (id 0) + value (id 1) axis. */
function makeChartExAxes(type: ChartType): string {
	if (type !== ChartType.waterfall) return ''
	const catAxis = el('cx:axis', { id: 0 }, [
		// chartEx catScaling gapWidth is a fraction (1.0 = 100%), NOT the classic integer percent.
		raw(voidEl('cx:catScaling', { gapWidth: '0.5' })),
		raw(voidEl('cx:tickLabels')),
	])
	const valAxis = el('cx:axis', { id: 1 }, [
		raw(voidEl('cx:valScaling')),
		raw(voidEl('cx:majorGridlines')),
		raw(voidEl('cx:tickLabels')),
	])
	return catAxis + valAxis
}

/** Build a minimal `<cx:title>` from the chart title options (only when `showTitle`). */
function makeChartExTitle(rel: SlideRelChart): string {
	if (!rel.opts.showTitle) return ''
	const color = rel.opts.titleColor || DEF_FONT_COLOR
	const face = rel.opts.titleFontFace || 'Calibri'
	const rPr = el('a:defRPr', null, [raw(genXmlColorSelection(color)), raw(createChartTextFonts(face))])
	const rich = el('cx:rich', null, [
		raw(voidEl('a:bodyPr')),
		raw(voidEl('a:lstStyle')),
		raw(
			el('a:p', null, [
				raw(el('a:pPr', null, raw(rPr))),
				raw(
					el('a:r', null, [
						raw(voidEl('a:rPr', { lang: rel.opts.lang || 'en-US' })),
						raw(el('a:t', null, rel.opts.title || 'Chart Title')),
					])
				),
			])
		),
	])
	return el('cx:title', { pos: 't', align: 'ctr', overlay: 0 }, raw(el('cx:tx', null, raw(rich))))
}

/**
 * Assemble the full `<cx:chartSpace>` for a chartEx chart part.
 * Document order (CT_ChartSpace): `chartData` → `chart`; the `<cx:chart>` holds
 * `title? → plotArea → legend?`.
 * @param {SlideRelChart} rel - the registered chartEx chart
 * @return {string} the `ppt/charts/chartEx{N}.xml` body
 */
export function makeXmlChartEx(rel: SlideRelChart): string {
	const type = rel.opts._type as ChartType

	const plotAreaRegion = el('cx:plotAreaRegion', null, raw(makeChartExSeries(rel)))
	const plotArea = el('cx:plotArea', null, [raw(plotAreaRegion), raw(makeChartExAxes(type))])
	const legend = rel.opts.showLegend
		? voidEl('cx:legend', { pos: chartExLegendPos(rel.opts.legendPos), align: 'ctr', overlay: 0 })
		: ''
	const chart = el('cx:chart', null, [raw(makeChartExTitle(rel)), raw(plotArea), raw(legend)])

	return (
		XML_DECL +
		el('cx:chartSpace', { 'xmlns:cx': CX_NS, 'xmlns:a': A_NS, 'xmlns:r': R_NS }, [
			raw(makeChartExData(rel)),
			raw(chart),
		])
	)
}
