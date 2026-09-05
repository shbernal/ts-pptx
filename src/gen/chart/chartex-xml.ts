/**
 * ts-pptx: chartEx (`cx:`) Chart Assembly
 *
 * Builds a chartEx chart part `ppt/charts/chartEx{N}.xml` — the `<cx:chartSpace>` for the
 * Office-2016 chart family (currently `waterfall`, `funnel`, `treemap`, `sunburst`, `histogram`,
 * `pareto`, `boxWhisker`, `regionMap`).
 * This is the chartEx analogue of
 * {@link ./chart-xml} (`makeXmlCharts`): a pure string builder, no I/O, no model mutation.
 *
 * chartEx differs from the classic `<c:chartSpace>` in three ways the rest of the package must
 * account for (all handled at the call sites, not here):
 * - it is a **separate part** with content type `application/vnd.ms-office.chartex+xml`,
 * - it is referenced from the slide through `<mc:AlternateContent>` (see `gen/slide/objects/chart.ts`),
 * - it renders **only in PowerPoint 2016+/Microsoft 365**; other consumers show the fallback shape.
 *
 * The `<cx:f>` formulas point back at the same embedded workbook the classic charts use
 * ({@link ./embed-xlsx}); the cell mapping lives in {@link ./chartex-data} / {@link ./data-refs}.
 */

import { ChartType } from '../../enums.js'
import { InternalError } from '../../errors.js'
import { DEF_FONT_COLOR, XML_DECL } from '../../constants-internal.js'
import type { ChartExBinning, ChartExGeography, ChartExStatistics } from '../../types/chart.js'
import type { SlideRelChart } from '../../types/internal.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { createChartTextFonts } from './chart-parts.js'
import { chartExSeriesNameRef, makeChartExData } from './chartex-data.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import { type XsdBool, xsdBool } from '../../ooxml/xsd-boolean.js'

/**
 * Map a chartEx {@link ChartType} to its `<cx:series layoutId>` token (the CT_SeriesLayout value
 * PowerPoint keys the chart geometry on). Throws on anything it does not own — see the `default`
 * arm for which members those are and why reaching it is a routing bug.
 */
function chartExLayoutId(type: ChartType): string {
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
		case ChartType.boxWhisker:
			return 'boxWhisker'
		case ChartType.regionMap:
			return 'regionMap'
		default:
			// The unmatched members are the classic 2007 catalog — owned by `makeChartType` in
			// `./chart-xml` — plus `pareto`, which is multi-series and is diverted to `makeParetoSeries`
			// before any single layoutId is asked for. Both are unreachable through `isChartExType`
			// routing; `''` here used to emit `<cx:series layoutId="">`, which PowerPoint cannot render.
			throw new InternalError(
				'chart/type-not-routed',
				`chartExLayoutId: "${String(type)}" has no <cx:series> layoutId — classic chart types belong to makeXmlCharts, pareto is built by makeParetoSeries, and a newly added chartEx type needs an arm here`,
				{ detail: { chartType: type } }
			)
	}
}

/**
 * A deterministic `uniqueId` GUID for a `<cx:series>`, derived from the chart's package-global id.
 * PowerPoint stamps a random GUID here; deriving it keeps output byte-stable (the project pins
 * chart determinism — the package-wide id this derives from is assigned at write time in
 * `package/assemble.ts`, STEP 2). `seriesIdx`
 * distinguishes multiple series within one chart (pareto emits two), varying the 4th GUID group so
 * every series still gets a distinct id.
 */
function chartExUniqueId(globalId: number, seriesIdx = 0): string {
	return `{00000000-0000-0000-${String(seriesIdx).padStart(4, '0')}-${String(globalId).padStart(12, '0')}}`
}

/** chartEx legend positions are `t|b|l|r`; clamp the classic `'tr'` (and anything else) to `'t'`. */
function chartExLegendPos(pos: string | undefined): string {
	return pos === 'b' || pos === 'l' || pos === 'r' ? pos : 't'
}

/**
 * The `<cx:axisId>` binding on a `<cx:series>` (which value/category axis the series plots against).
 *
 * NOTE — a schema-vs-PowerPoint divergence (the exact mirror of the histogram `binCount` gotcha):
 * the OpenXML SDK models `cx:axisId` as an `OpenXmlLeafTextElement`, so the OOXML validator only
 * accepts the text-content form `<cx:axisId>1</cx:axisId>` and flags the `val` attribute as
 * undeclared. But PowerPoint desktop REFUSES to open a deck using the text form (0x80070570) — it
 * writes and requires the ATTRIBUTE form `<cx:axisId val="1"/>` (matching the classic `c:axId`).
 * PowerPoint is the authoritative oracle here, so the attribute form is emitted; the resulting
 * validator complaint on `cx:axisId/@val` is expected and whitelisted in the schema-case test.
 */
function makeChartExAxisId(id: number): string {
	return voidEl('cx:axisId', { val: id })
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

/**
 * Build the box-and-whisker `<cx:layoutPr>` body — a `<cx:visibility>` toggle set plus a
 * `<cx:statistics>` quartile-method choice, in that document order (what PowerPoint itself emits).
 * All flags default to PowerPoint's own defaults: exclusive quartiles, mean marker on, outlier
 * points on, the mean-line and the full non-outlier scatter off. Booleans map to the `0`/`1`
 * attribute form PowerPoint uses.
 */
function makeBoxWhiskerLayoutPr(stats: ChartExStatistics | undefined): string {
	const bit = (v: boolean | undefined, dflt: boolean): XsdBool => xsdBool(v ?? dflt)
	const visibility = voidEl('cx:visibility', {
		meanLine: bit(stats?.meanLine, false),
		meanMarker: bit(stats?.meanMarker, true),
		nonoutliers: bit(stats?.nonoutliers, false),
		outliers: bit(stats?.outliers, true),
	})
	const statistics = voidEl('cx:statistics', {
		quartileMethod: stats?.quartileMethod === 'inclusive' ? 'inclusive' : 'exclusive',
	})
	return el('cx:layoutPr', null, [raw(visibility), raw(statistics)])
}

/**
 * Build the region-map `<cx:layoutPr>` body — a single `<cx:geography>` hint telling PowerPoint how
 * to interpret the category labels as geographic region names. `cultureLanguage`/`cultureRegion`
 * (both schema-REQUIRED) resolve ambiguous names (e.g. "Georgia" the country vs. the US state);
 * `attribution` is Bing's fixed acknowledgement string.
 *
 * NOTE — PowerPoint itself also nests a `<cx:geoCache>` here holding a base64 binary blob: its
 * cached Bing geometry for the resolved regions. That blob is produced by an online Bing lookup and
 * cannot be reproduced offline, so it is intentionally OMITTED. PowerPoint re-resolves the geography
 * from the region names when the deck is opened (verified over COM: the deck opens and reads back as
 * a regionMap either way) — this is why a region map is documented as best-effort/write-only and
 * renders blank if the names don't match PowerPoint's geography database.
 */
function makeGeographyLayoutPr(geo: ChartExGeography | undefined): string {
	const geography = voidEl('cx:geography', {
		cultureLanguage: geo?.cultureLanguage || 'en-US',
		cultureRegion: geo?.cultureRegion || 'US',
		attribution: 'Powered by Bing',
	})
	return el('cx:layoutPr', null, raw(geography))
}

/** The `<cx:tx>` series-name cell/value block (shared by every layout that names its series). */
function makeChartExSeriesName(rel: SlideRelChart): string {
	return el(
		'cx:tx',
		null,
		raw(
			el('cx:txData', null, [
				raw(el('cx:f', null, chartExSeriesNameRef(rel))),
				raw(el('cx:v', null, rel.data[0]?.name ?? '')),
			])
		)
	)
}

/**
 * Data labels: chartEx toggles each field via `<cx:visibility>`. Only emitted when the caller asked
 * to show values, matching the classic default of no data labels.
 */
function makeChartExDataLabels(rel: SlideRelChart): string {
	const showValue = !!(rel.opts.showValue || rel.opts.showLabel)
	return showValue
		? el('cx:dataLabels', { pos: 'outEnd' }, raw(voidEl('cx:visibility', { seriesName: 0, categoryName: 0, value: 1 })))
		: ''
}

/**
 * Build the pareto `<cx:series>` PAIR — pareto is the first chartEx layout that emits more than one
 * series. Series 0 is a `clusteredColumn` bound to value axis 1 whose `<cx:aggregation/>` tells
 * PowerPoint to sum the values per category, sort the bars descending, and drive the cumulative
 * line. Series 1 is the `paretoLine` itself: it carries `ownerIdx="0"` (it derives its data from
 * series 0, so it has NO `<cx:tx>` or `<cx:dataId>`) and binds to the secondary percentage axis 2.
 */
function makeParetoSeries(rel: SlideRelChart): string {
	const bar = el('cx:series', { layoutId: 'clusteredColumn', uniqueId: chartExUniqueId(rel.globalId, 0) }, [
		raw(makeChartExSeriesName(rel)),
		raw(makeChartExDataLabels(rel)),
		raw(voidEl('cx:dataId', { val: 0 })),
		raw(el('cx:layoutPr', null, raw(voidEl('cx:aggregation')))),
		raw(makeChartExAxisId(1)),
	])
	const line = el('cx:series', { layoutId: 'paretoLine', ownerIdx: 0, uniqueId: chartExUniqueId(rel.globalId, 1) }, [
		raw(makeChartExAxisId(2)),
	])
	return bar + line
}

/** Build the `<cx:series>` (title cell, data labels, dataId, and layout-specific `<cx:layoutPr>`). */
function makeChartExSeries(rel: SlideRelChart): string {
	const opts = rel.opts
	const type = opts._type as ChartType

	// Pareto is multi-series (a column series + a cumulative line on a secondary axis).
	if (type === ChartType.pareto) return makeParetoSeries(rel)

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
	} else if (type === ChartType.boxWhisker) {
		// box-and-whisker: quartile method + which box adornments (mean line/marker, outliers) show.
		layoutPr = makeBoxWhiskerLayoutPr(opts.statistics)
	} else if (type === ChartType.regionMap) {
		// region map: the <cx:geography> hint PowerPoint resolves the region names against.
		layoutPr = makeGeographyLayoutPr(opts.geography)
	}

	return el('cx:series', { layoutId: chartExLayoutId(type), uniqueId: chartExUniqueId(rel.globalId) }, [
		raw(makeChartExSeriesName(rel)),
		raw(makeChartExDataLabels(rel)),
		raw(voidEl('cx:dataId', { val: 0 })),
		raw(layoutPr),
	])
}

/** The chartEx category axis: a gap width and tick labels. `id` is 0 except on the funnel. */
function catScalingAxis(gapWidth: string, id = 0): string {
	return el('cx:axis', { id }, [raw(voidEl('cx:catScaling', { gapWidth })), raw(voidEl('cx:tickLabels'))])
}

/** The chartEx primary value axis (id 1): a linear scale, gridlines and tick labels. */
function gridlinedValAxis(): string {
	return el('cx:axis', { id: 1 }, [
		raw(voidEl('cx:valScaling')),
		raw(voidEl('cx:majorGridlines')),
		raw(voidEl('cx:tickLabels')),
	])
}

/**
 * The category-axis gap width each axis-bearing chartEx layout uses. A chartEx `catScaling`
 * gapWidth is a **fraction** (1.0 = 100%), NOT the classic integer percent.
 *
 * Histogram bins abut, so they take 0; the pareto's aggregated columns are histogram columns
 * and take the same. A layout absent from this table has no axes at all: the hierarchical
 * treemap and sunburst encode categories in nested tiles and rings rather than an axis scale,
 * and a `regionMap` plots on geography.
 */
const CHARTEX_CAT_GAP_WIDTH: Partial<Record<ChartType, string>> = {
	[ChartType.waterfall]: '0.5',
	[ChartType.histogram]: '0',
	[ChartType.pareto]: '0',
	[ChartType.boxWhisker]: '1',
}

/**
 * Build the `<cx:axis>` list for a layout, matched to what PowerPoint itself emits for each.
 *
 * Four of the five axis-bearing layouts are the same pair — a category axis (id 0) and a
 * gridlined value axis (id 1) — differing only in gap width, so they are a table rather than
 * four arms. The two that are not: the **funnel** has a SINGLE category axis, which PowerPoint
 * numbers id 1, with no value axis and no gridlines (its bars run horizontally off one
 * category scale); and the **pareto** adds a third axis for its cumulative line.
 */
function makeChartExAxes(type: ChartType): string {
	// The funnel is the one asymmetric layout: a single axis, and it is the *value* id the
	// category scale hangs off, so it is not a row in the table above.
	if (type === ChartType.funnel) return catScalingAxis('2.19', 1)

	const gapWidth = CHARTEX_CAT_GAP_WIDTH[type]
	if (gapWidth === undefined) return ''
	const axes = catScalingAxis(gapWidth) + gridlinedValAxis()
	// The pareto adds a SECONDARY value axis (id 2) for the cumulative paretoLine: a 0..1
	// percentage scale, no gridlines, since the primary value axis already carries them.
	if (type !== ChartType.pareto) return axes
	return (
		axes +
		el('cx:axis', { id: 2 }, [
			raw(voidEl('cx:valScaling', { max: '1', min: '0' })),
			raw(voidEl('cx:units', { unit: 'percentage' })),
			raw(voidEl('cx:tickLabels')),
		])
	)
}

/** Build a minimal `<cx:title>` from the chart title options (only when `showTitle`). */
function makeChartExTitle(rel: SlideRelChart): string {
	if (!rel.opts.showTitle) return ''
	const color = rel.opts.titleColor ?? DEF_FONT_COLOR
	const face = rel.opts.titleFontFace ?? 'Calibri'
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
		el('cx:chartSpace', { 'xmlns:cx': OOXML_NS.cx, 'xmlns:a': OOXML_NS.a, 'xmlns:r': OOXML_NS.r }, [
			raw(makeChartExData(rel)),
			raw(chart),
		])
	)
}
