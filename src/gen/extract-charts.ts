/**
 * ts-pptx: the chart half of a slide extraction
 *
 * Apart from `gen/extract-slides.ts` for one reason: it names the chart emitters, and extraction
 * is reachable from every program that can build a deck. A module of its own is what lets the
 * extraction load it on a deck that has a chart and never mention it on a deck that has none.
 */

import type { ExtractedSlide } from '../read/api/presentation-types.js'
import type { PresSlideInternal } from '../types/internal.js'
import { buildEmbeddedWorksheet, chartPartBodies } from './chart/embed-xlsx.js'

/**
 * Charts: the chart part XML plus its embedded workbook bytes. The chart part's own `.rels`
 * (workbook reference) is rebuilt on injection.
 *
 * chartEx charts are a different part behind a different rel type and content type, and
 * PowerPoint reports one as corrupt without its style/colors sidecars (see
 * gen/chart/chartex-style.ts). Both ride in the descriptor's `chartEx` slot, which is what tells
 * `appendSlides` the two shapes apart: it cannot be inferred from the XML. `chartPartBodies` makes
 * the classic-or-chartEx choice for the package writer too.
 */
export function chartsOf(slide: PresSlideInternal): ExtractedSlide['charts'] {
	return (slide._relsChart || []).map((rel) => ({
		rId: rel.rId,
		embeddingBytes: buildEmbeddedWorksheet(rel),
		...chartPartBodies(rel),
	}))
}
