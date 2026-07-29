/**
 * ts-pptx: chart slide-object serialization
 *
 * Emits a `chart` slide object as the `<p:graphicFrame>` that references the chart part.
 * Classic charts emit that frame bare; chartEx charts wrap it in `<mc:AlternateContent>` with a
 * plain-shape `<mc:Fallback>` for consumers older than PowerPoint 2016.
 */

import { ChartType, isChartExType } from '../../../core-enums.js'
import type { ObjectOptions } from '../../../core-interfaces.js'
import type { SlideObject } from '../../../types/internal.js'
import { genXmlPlaceholder } from '../../drawingml/text-body.js'
import { el, raw, voidEl } from '../../oxml/el.js'
import { cNvPrOpen, MC_NS } from './shared.js'

const DML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
/** graphicData URI + child namespace for chartEx charts (referenced via `<mc:AlternateContent>`). */
const CHARTEX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
/**
 * chartEx feature-version namespace declared on `<mc:Choice Requires>`. Each 2016 chart wave
 * introduced a feature level a consumer must "understand" to render the chart; a consumer that
 * doesn't falls through to `<mc:Fallback>`. Keyed by `ChartType`.
 */
const CHARTEX_FEATURE_NS: Partial<Record<ChartType, { prefix: string; uri: string }>> = {
	[ChartType.waterfall]: { prefix: 'cx1', uri: 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex' },
	[ChartType.funnel]: { prefix: 'cx2', uri: 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex' },
	[ChartType.regionMap]: { prefix: 'cx4', uri: 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex' },
}

/**
 * Render a `chart` slide object to its `<p:graphicFrame>` XML referencing the chart part.
 *
 * Classic charts emit a bare `<p:graphicFrame>` pointing at a `<c:chart>`. chartEx charts
 * (waterfall, …) instead wrap that graphicFrame in `<mc:AlternateContent>`: an `<mc:Choice>`
 * carrying the `<cx:chart>` reference (rendered by PowerPoint 2016+/Microsoft 365) and an
 * `<mc:Fallback>` placeholder shape shown by every other consumer.
 */
export function renderChartObject(
	slideItemObj: SlideObject,
	idx: number,
	x: number,
	y: number,
	cx: number,
	cy: number,
	placeholderObj: SlideObject | null
): string {
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	const opts = slideItemObj.options
	// A chart's options are really `ChartOptsInternal`; `_type` is not on the broad `ObjectOptions`.
	const chartType = (opts as { _type?: ChartType })._type
	const isChartEx = isChartExType(chartType)

	const graphicDataUri = isChartEx ? CHARTEX_NS : CHART_NS
	const chartChild = isChartEx
		? voidEl('cx:chart', { 'xmlns:cx': CHARTEX_NS, 'r:id': `rId${slideItemObj.chartRid}` }, { openPrefix: '   ' })
		: voidEl('c:chart', { 'r:id': `rId${slideItemObj.chartRid}`, 'xmlns:c': CHART_NS }, { openPrefix: '   ' })

	const graphicFrame = el('p:graphicFrame', null, [
		raw(
			el(
				'p:nvGraphicFramePr',
				null,
				[
					raw(cNvPrOpen(idx + 2, opts.objectName, opts.altText || '', '   ') + '/>'),
					raw(voidEl('p:cNvGraphicFramePr', null, { openPrefix: '   ' })),
					raw(el('p:nvPr', null, raw(genXmlPlaceholder(placeholderObj)), { openPrefix: '   ' })),
				],
				{ openPrefix: ' ', closePrefix: ' ' }
			)
		),
		raw(el('p:xfrm', null, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))], { openPrefix: ' ' })),
		raw(
			el(
				'a:graphic',
				{ 'xmlns:a': DML_NS },
				raw(el('a:graphicData', { uri: graphicDataUri }, raw(chartChild), { openPrefix: '  ', closePrefix: '  ' })),
				{ openPrefix: ' ', closePrefix: ' ' }
			)
		),
	])

	if (!isChartEx) return graphicFrame

	// chartEx: wrap the graphicFrame in <mc:AlternateContent>. The Choice declares the feature-level
	// namespace it Requires; the Fallback is a plain shape so non-2016 consumers show something.
	const feature = (chartType && CHARTEX_FEATURE_NS[chartType]) ?? CHARTEX_FEATURE_NS[ChartType.waterfall]
	const choice = el(
		'mc:Choice',
		{ [`xmlns:${feature?.prefix ?? 'cx1'}`]: feature?.uri, Requires: feature?.prefix ?? 'cx1' },
		raw(graphicFrame)
	)
	const fallback = el('mc:Fallback', null, raw(renderChartExFallback(idx, opts, x, y, cx, cy)))
	return el('mc:AlternateContent', { 'xmlns:mc': MC_NS }, [raw(choice), raw(fallback)])
}

/**
 * Build the `<mc:Fallback>` placeholder shape shown when a chartEx chart cannot be rendered
 * (any consumer older than PowerPoint 2016 / Microsoft 365). A light-grey rectangle at the chart's
 * position carrying a short explanatory note — enough that the slide reads sensibly rather than
 * showing a void where the chart would be.
 */
function renderChartExFallback(idx: number, opts: ObjectOptions, x: number, y: number, cx: number, cy: number): string {
	return el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(cNvPrOpen(idx + 2, opts.objectName, opts.altText || '') + '/>'),
				raw(voidEl('p:cNvSpPr')),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(
			el('p:spPr', null, [
				raw(el('a:xfrm', null, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])),
				raw(el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')))),
				raw(el('a:solidFill', null, raw(voidEl('a:srgbClr', { val: 'F2F2F2' })))),
				raw(el('a:ln', null, raw(el('a:solidFill', null, raw(voidEl('a:srgbClr', { val: 'BFBFBF' })))))),
			])
		),
		raw(
			el('p:txBody', null, [
				raw(voidEl('a:bodyPr')),
				raw(voidEl('a:lstStyle')),
				raw(
					el('a:p', null, [
						raw(
							el('a:r', null, [
								raw(voidEl('a:rPr', { lang: 'en-US' })),
								raw(el('a:t', null, 'This chart requires PowerPoint 2016 or newer to display.')),
							])
						),
					])
				),
			])
		),
	])
}
