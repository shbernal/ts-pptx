/**
 * ts-pptx: the measurement construct family
 *
 * The three presentation methods that measure without authoring anything: `measureText`,
 * `overflowsBox` and `tableLayout`. They measure through `measure/fit.ts` and
 * `measure/table-fit.ts`, which the core reaches anyway through the export-time bake below.
 *
 * Only one of the three is a second spelling of something the `pptx-ts/measure` subpath also
 * exports, and even that one differs where it matters: the subpath's `measureText` takes a
 * `FontMetricsRegistry`, this binds the presentation's own. `overflowsBox` and `tableLayout` have
 * no subpath equivalent at all -- the first composes two of the subpath's primitives, and the
 * second needs `pres.presLayout`, which a free function has no way to reach. So the pair is a
 * primitive and its bound convenience, not one capability written twice.
 *
 * The export-time autofit bake is NOT here. `applyMeasuredFit` (`gen/prepare.ts`) runs on the
 * ordinary write path for `fit:'shrink'` text, so it belongs to the core rather than to a family a
 * program can leave out.
 */

import { measureText, requireBoxHeight } from '../measure/fit.js'
import { computeTableLayout } from '../measure/table-fit.js'
import type { ConstructFamily } from './shared.js'

export const measureFamily = {
	name: 'measure',
	presentationAuthors: {
		measureText(ctx, text, opts) {
			return measureText(ctx.fontMetrics, text, opts)
		},
		overflowsBox(ctx, text, opts) {
			// Checked first: text that cannot be measured never reached the height check.
			requireBoxHeight(opts.hIn, 'overflowsBox')
			const m = measureText(ctx.fontMetrics, text, opts)
			return m.measurable && !m.fitsBox(opts.hIn)
		},
		tableLayout(ctx, rows, opts) {
			return computeTableLayout(rows, opts, ctx.pres.presLayout, ctx.fontMetrics)
		},
	},
} satisfies ConstructFamily
