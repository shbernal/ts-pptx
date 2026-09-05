/**
 * ts-pptx: the chart construct family
 *
 * The expensive one: `addChart` reaches every plot module, axis and sidecar builder under
 * `gen/chart/`, which is what a program that never draws a chart is paying for when a slide method
 * names it by import.
 */

import { SlideObjectType, type CHART_NAME } from '../enums.js'
import { renderChartObject } from '../gen/slide/objects/chart.js'

import { InvalidOptionError } from '../errors.js'
import type { ChartMulti, OptsChartData } from '../types/index.js'
import { addChartDefinition } from '../gen/define/chart.js'
import type { ConstructFamily } from './shared.js'

/** Distinguish a multi-type (combo) chart array (`ChartMulti[]`) from a single chart's data (`OptsChartData[]`). */
function isMultiChart(arg: OptsChartData[] | ChartMulti[]): arg is ChartMulti[] {
	const first = arg[0] as Partial<ChartMulti> | undefined
	return !!first && typeof first === 'object' && 'type' in first && 'data' in first
}

export const chartFamily: ConstructFamily = {
	name: 'chart',
	authors: {
		addChart(slide, arg1, arg2) {
			let type: CHART_NAME | ChartMulti[]
			let data: OptsChartData[]

			const options = arg2 ?? {}
			if (Array.isArray(arg1) && isMultiChart(arg1)) {
				// Multi-type (combo) chart: addChart(ChartMulti[], options?)
				type = arg1
				data = []
			} else {
				// Canonical single-type form: addChart(data, { type, ...options })
				data = arg1 ?? []
				const optType = options.type
				if (!optType) {
					throw new InvalidOptionError(
						'chart/missing-type',
						'addChart: a chart `type` is required on the options object, e.g. addChart(data, { type: ChartType.bar }).'
					)
				}
				type = optType
			}

			// `_type` is set by `addChartDefinition` on its own copy of the options -- stamping it here
			// too would write an internal field onto the caller's object for no gain.
			// addChartDefinition's multi-type branch reads the shared options from its `data` slot
			if (Array.isArray(type)) addChartDefinition(slide, type, options, undefined)
			else addChartDefinition(slide, type, data, options)
		},
	},
	children: {
		chart(target, child) {
			addChartDefinition(target, child.type, child.data, child.options || {})
		},
	},
	renderers: {
		[SlideObjectType.chart]: renderChartObject,
	},
}
