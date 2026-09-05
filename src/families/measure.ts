/**
 * ts-pptx: the measurement construct family
 *
 * The three presentation methods that measure without authoring anything: `measureText`,
 * `overflowsBox` and `tableLayout`. They are a convenience over what the `ts-pptx/measure` subpath
 * already exports as free functions, and they are the only reason `measure/fit.ts` and
 * `measure/table-fit.ts` are reachable from a deck that never asks for a measurement.
 *
 * The export-time autofit bake is NOT here. `applyMeasuredFit` (`gen/prepare.ts`) runs on the
 * ordinary write path for `fit:'shrink'` text, so it belongs to the core rather than to a family a
 * program can leave out.
 */

import { measureText } from '../measure/fit.js'
import { computeTableLayout } from '../measure/table-fit.js'
import type { ConstructFamily } from './shared.js'

export const measureFamily = {
	name: 'measure',
	presentationAuthors: {
		measureText(ctx, text, opts) {
			return measureText(ctx.fontMetrics, text, opts)
		},
		overflowsBox(ctx, text, opts) {
			const m = measureText(ctx.fontMetrics, text, opts)
			return m.measurable && !m.fitsBox(opts.hIn)
		},
		tableLayout(ctx, rows, opts) {
			return computeTableLayout(rows, opts, ctx.pres.presLayout, ctx.fontMetrics)
		},
	},
} satisfies ConstructFamily
