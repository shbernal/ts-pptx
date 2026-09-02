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
 * `ChartOptsInternal._type` widens to `ChartMultiInternal[]` for a combo chart, which is not any single
 * kind, so every predicate takes the union and answers `false` for the array.
 */

import { ChartType } from '../../enums.js'
import type { ChartPropsChartStock } from '../../types/index.js'
import type { ChartMultiInternal } from '../../types/internal.js'

/** A `ChartOptsInternal._type`: one chart kind, or a combo chart's list of subcharts. */
type ChartTypeOrCombo = ChartType | ChartMultiInternal[] | undefined

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

/**
 * The stock-chart styles, keyed off the public option so the type is the source rather than a
 * fourth transcription of the same four names.
 */
export type StockStyle = NonNullable<ChartPropsChartStock['stockStyle']>

/**
 * Per-style stock chart geometry: how many value series the style expects, whether the first
 * series is a Volume column drawn as a bar (on its own axis pair), and whether the open-close
 * `<c:upDownBars>` are drawn. HLC/VHLC are three-value (no open) and instead mark the close
 * with a dot; OHLC/VOHLC are four-value and use up/down bars for the open-close body.
 *
 * Here rather than in the plot builder because `addChartDefinition` needs the series counts to
 * warn on a mismatched series list, and had a second table of its own for exactly that.
 */
export const STOCK_STYLE_SPEC: Record<StockStyle, { seriesCount: number; volume: boolean; upDownBars: boolean }> = {
	hlc: { seriesCount: 3, volume: false, upDownBars: false },
	ohlc: { seriesCount: 4, volume: false, upDownBars: true },
	vhlc: { seriesCount: 4, volume: true, upDownBars: false },
	vohlc: { seriesCount: 5, volume: true, upDownBars: true },
}
