/**
 * ts-pptx: chart-kind predicates
 *
 * The classification questions the chart emitters ask about a chart's `_type`, in one place.
 * There is nothing clever here — each is the comparison the call sites were already writing —
 * but writing it once is what keeps `bubble3d` from being forgotten at the eleventh site the
 * way it has been at a few of the first ten.
 *
 * "XY" is the distinction that actually drives markup rather than a name: a scatter or bubble
 * chart plots numbers on both axes, so its *category* axis is a second `<c:valAx>`, its series
 * carry `xVal`/`yVal` rather than `cat`/`val`, and its embedded worksheet is laid out to match.
 *
 * `ChartOptsInternal._type` widens to `ChartMulti[]` for a combo chart, which is not any single
 * kind, so every predicate takes the union and answers `false` for the array.
 */

import { ChartType } from '../../enums.js'
import type { ChartMulti } from '../../types/index.js'

/** A `ChartOptsInternal._type`: one chart kind, or a combo chart's list of subcharts. */
type ChartTypeOrCombo = ChartType | ChartMulti[] | undefined

/** Both bubble variants (`<c:bubbleChart>`), which differ only by `<c:bubble3D>`. */
export function isBubbleChart(type: ChartTypeOrCombo): boolean {
	return type === ChartType.bubble || type === ChartType.bubble3d
}

/** The scatter chart (`<c:scatterChart>`). */
export function isScatterChart(type: ChartTypeOrCombo): boolean {
	return type === ChartType.scatter
}

/**
 * A chart whose categories are numbers rather than labels: scatter and both bubbles.
 * These are the ones whose category axis is emitted as a `<c:valAx>`.
 */
export function isXyChart(type: ChartTypeOrCombo): boolean {
	return isScatterChart(type) || isBubbleChart(type)
}
