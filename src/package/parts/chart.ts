/**
 * ts-pptx: the chart contribution to a written package
 *
 * The expensive one. A chart rel pulls in its `chart.xml`, the embedded workbook PowerPoint opens
 * when the user edits the data, the chart part's own rels, and — for a chartEx chart — the
 * style/colors sidecars without which PowerPoint reports the deck as corrupt. `createExcelWorksheet`
 * reaches every plot module, axis and sidecar builder under `gen/chart/`, which is why the packager
 * must not name it: doing so put ~22 kB gzip of chart code in the graph of a text-only deck.
 *
 * What is *not* here: the pass that assigns each chart rel its `globalId`, filename and Target.
 * That is package-wide ordering rather than chart knowledge — it numbers across slides, layouts and
 * the master before anything is written, and `gen/chart/chartex-xml.ts` derives series GUIDs from
 * the id it hands out — so it stays on the core path in `package/assemble.ts`.
 */

import type { ZipWriter } from '../../zip.js'
import type { ContentTypeDefault, ContentTypeOverride } from '../../gen/opc/content-types.js'
import type {
	PresentationPropsInternal,
	PresSlideInternal,
	SlideLayoutInternal,
	SlideRelChart,
} from '../../types/internal.js'
import { createExcelWorksheet } from '../../gen/chart/embed-xlsx.js'
import type { PartContributor, PartTarget } from './shared.js'

const CT_CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
// chartEx (cx:) parts use Microsoft content types, NOT the openxmlformats prefix. Each chartEx
// chart part also requires a chart-style + color-style sidecar part.
const CT_CHARTEX = 'application/vnd.ms-office.chartex+xml'
const CT_CHARTEX_STYLE = 'application/vnd.ms-office.chartstyle+xml'
const CT_CHARTEX_COLORS = 'application/vnd.ms-office.chartcolorstyle+xml'
/** The embedded workbook every chart carries; one Default covers all of them. */
const XLSX_DEFAULT: ContentTypeDefault = {
	extension: 'xlsx',
	contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}

/**
 * Content-type Override(s) for a chart rel. A classic chart is one Override; a chartEx chart is
 * three: the `chartEx{N}.xml` part plus its mandatory `style{N}.xml` and `colors{N}.xml` sidecars
 * (keyed to the same `globalId`).
 */
function chartOverrides(rel: SlideRelChart, leadingSpace = false): ContentTypeOverride[] {
	if (!rel.isChartEx) return [{ partName: rel.Target, contentType: CT_CHART, leadingSpace }]
	return [
		{ partName: rel.Target, contentType: CT_CHARTEX, leadingSpace },
		{ partName: `/ppt/charts/style${rel.globalId}.xml`, contentType: CT_CHARTEX_STYLE, leadingSpace },
		{ partName: `/ppt/charts/colors${rel.globalId}.xml`, contentType: CT_CHARTEX_COLORS, leadingSpace },
	]
}

export const chartContributor: PartContributor = {
	order: 30,
	parts: {
		fromTargetRels(target: PartTarget, zip: ZipWriter): Promise<unknown>[] {
			return (target._relsChart || []).map((rel) => createExcelWorksheet(rel, zip))
		},
	},
	contentTypes: {
		defaults(pres: PresentationPropsInternal): ContentTypeDefault[] {
			const targets: Array<PartTarget | undefined> = [...pres.slides, ...pres.slideLayouts, pres.masterSlide]
			return targets.some((target) => (target?._relsChart || []).length > 0) ? [XLSX_DEFAULT] : []
		},
		perSlide(slide: PresSlideInternal): ContentTypeOverride[] {
			return slide._relsChart.flatMap((rel) => chartOverrides(rel))
		},
		perLayout(layout: SlideLayoutInternal): ContentTypeOverride[] {
			return (layout._relsChart || []).flatMap((rel) => chartOverrides(rel, true))
		},
		trailing(pres: PresentationPropsInternal): ContentTypeOverride[] {
			return (pres.masterSlide?._relsChart || []).flatMap((rel) => chartOverrides(rel, true))
		},
	},
}
