/**
 * ts-pptx: the chart half of a slide extraction
 *
 * Apart from `gen/extract-slides.ts` for one reason: it names the chart emitters, and extraction
 * is reachable from every program that can build a deck. A module of its own is what lets the
 * extraction load it on a deck that has a chart and never mention it on a deck that has none.
 */

import type { ExtractedSlide } from '../read/api/presentation-types.js'
import type { PresSlideInternal } from '../types/internal.js'
import { makeXmlCharts } from './chart/chart-xml.js'
import { makeXmlChartEx } from './chart/chartex-xml.js'
import { makeChartExColorsXml, makeChartExStyleXml } from './chart/chartex-style.js'
import { buildEmbeddedWorksheet } from './chart/embed-xlsx.js'

/**
 * Charts: the chart part XML plus its embedded workbook bytes. The chart part's own `.rels`
 * (workbook reference) is rebuilt on injection.
 *
 * chartEx charts are a different part (`makeXmlChartEx`) behind a different rel type and
 * content type, and PowerPoint reports one as corrupt without its style/colors sidecars
 * (see gen/chart/chartex-style.ts). Both ride in the descriptor's `chartEx` slot, which
 * is what tells `appendSlides` the two shapes apart: it cannot be inferred from the XML,
 * and building one as a classic chart is what used to produce a `<c:chartSpace>` with
 * axes and no plot behind a slide still pointing at it through `<cx:chart>`.
 */
export function chartsOf(slide: PresSlideInternal): ExtractedSlide['charts'] {
	return (slide._relsChart || []).map((rel) => {
		const base = { rId: rel.rId, embeddingBytes: buildEmbeddedWorksheet(rel) }
		return rel.isChartEx
			? {
					...base,
					chartXml: makeXmlChartEx(rel),
					chartEx: { styleXml: makeChartExStyleXml(), colorsXml: makeChartExColorsXml() },
				}
			: { ...base, chartXml: makeXmlCharts(rel) }
	})
}
