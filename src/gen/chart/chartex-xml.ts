/**
 * PptxGenJS: chartEx (`cx:`) Chart Assembly
 *
 * Builds a chartEx chart part `ppt/charts/chartEx{N}.xml` — the `<cx:chartSpace>` for the
 * Office-2016 chart family (currently `waterfall`, `funnel`, `treemap`, `sunburst`, `histogram`).
 * This is the chartEx analogue of
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
import type { ChartExBinning } from '../../types/chart.js'
import type { SlideRelChart } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { createChartTextFonts } from './chart-parts.js'
import { chartExSeriesNameRef, makeChartExData } from './chartex-data.js'

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
		case ChartType.funnel:
			return 'funnel'
		case ChartType.treemap:
			return 'treemap'
		case ChartType.sunburst:
			return 'sunburst'
		case ChartType.histogram:
			// A histogram is a clustered-column series whose bins PowerPoint computes from <cx:binning>.
			return 'clusteredColumn'
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

/**
 * Build the histogram `<cx:binning>` element from the caller's `binning` opt. PowerPoint bins the
 * raw observations automatically; the only control wired here is `intervalClosed` — whether each
 * bin interval is right-closed (`"r"`, PowerPoint's default) or left-closed (`"l"`).
 *
 * NOTE: explicit bin geometry (`binCount`/`binSize` child elements, `underflow`/`overflow`
 * attributes) is intentionally NOT emitted. Although `<cx:binCount>`/`<cx:binSize>` pass OOXML
 * schema validation (they are declared children of `cx:binning` in the OpenXML SDK), PowerPoint
 * desktop refuses to open a deck that contains them (0x80004005) — so shipping them would corrupt
 * the file. They stay deferred until their real requirements are pinned down against PowerPoint.
 */
function makeChartExBinning(binning: ChartExBinning | undefined): string {
	return voidEl('cx:binning', { intervalClosed: binning?.intervalClosed === 'l' ? 'l' : 'r' })
}

/** Build the `<cx:series>` (title cell, data labels, dataId, and layout-specific `<cx:layoutPr>`). */
function makeChartExSeries(rel: SlideRelChart): string {
	const opts = rel.opts
	const type = opts._type as ChartType
	const showValue = !!(opts.showValue || opts.showLabel)

	const tx = el(
		'cx:tx',
		null,
		raw(
			el('cx:txData', null, [
				raw(el('cx:f', null, chartExSeriesNameRef(rel))),
				raw(el('cx:v', null, rel.data[0]?.name ?? '')),
			])
		)
	)

	// Data labels: chartEx toggles each field via <cx:visibility>. Only emitted when the caller
	// asked to show values, matching the classic default of no data labels.
	const dataLabels = showValue
		? el('cx:dataLabels', { pos: 'outEnd' }, raw(voidEl('cx:visibility', { seriesName: 0, categoryName: 0, value: 1 })))
		: ''

	// Layout-specific series props.
	let layoutPr = ''
	if (type === ChartType.waterfall && Array.isArray(opts.subtotals) && opts.subtotals.length > 0) {
		// waterfall: subtotal/total column indices.
		const idxs = opts.subtotals.map((idx) => voidEl('cx:idx', { val: idx })).join('')
		layoutPr = el('cx:layoutPr', null, raw(el('cx:subtotals', null, raw(idxs))))
	} else if (type === ChartType.treemap) {
		// treemap: how parent-category banners are placed (PowerPoint's default is `overlapping`).
		layoutPr = el('cx:layoutPr', null, raw(voidEl('cx:parentLabelLayout', { val: 'overlapping' })))
	} else if (type === ChartType.histogram) {
		// histogram: how PowerPoint bins the raw observations. Default is auto binning with
		// right-closed intervals; the `binning` opt exposes explicit width/count/overflow control.
		layoutPr = el('cx:layoutPr', null, raw(makeChartExBinning(opts.binning)))
	}

	return el('cx:series', { layoutId: chartExLayoutId(type), uniqueId: chartExUniqueId(rel.globalId) }, [
		raw(tx),
		raw(dataLabels),
		raw(voidEl('cx:dataId', { val: 0 })),
		raw(layoutPr),
	])
}

/**
 * Build the `<cx:axis>` list for a layout (matched to what PowerPoint itself emits for each):
 * - **waterfall** — a category (id 0) + value (id 1) axis.
 * - **funnel** — a SINGLE category axis, which PowerPoint numbers id 1, with no value axis and no
 *   gridlines (the bars run horizontally off one category scale).
 * Every other layout returns no axes — the hierarchical treemap/sunburst are genuinely axis-free
 * (categories are encoded by the nested tiles/rings, not an axis scale).
 */
function makeChartExAxes(type: ChartType): string {
	// chartEx catScaling gapWidth is a fraction (1.0 = 100%), NOT the classic integer percent.
	if (type === ChartType.waterfall) {
		const catAxis = el('cx:axis', { id: 0 }, [
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
	if (type === ChartType.funnel) {
		return el('cx:axis', { id: 1 }, [raw(voidEl('cx:catScaling', { gapWidth: '2.19' })), raw(voidEl('cx:tickLabels'))])
	}
	if (type === ChartType.histogram) {
		// Histogram bins abut, so the category axis uses gapWidth 0; the value axis carries gridlines.
		const catAxis = el('cx:axis', { id: 0 }, [
			raw(voidEl('cx:catScaling', { gapWidth: '0' })),
			raw(voidEl('cx:tickLabels')),
		])
		const valAxis = el('cx:axis', { id: 1 }, [
			raw(voidEl('cx:valScaling')),
			raw(voidEl('cx:majorGridlines')),
			raw(voidEl('cx:tickLabels')),
		])
		return catAxis + valAxis
	}
	return ''
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
