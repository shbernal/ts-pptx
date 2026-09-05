/**
 * ts-pptx: the table family's live-DOM half
 *
 * `tableToSlides` reproduces a rendered HTML `<table>` as a PowerPoint table, resolving the element
 * id against the global `document`. It is the table family's, but it is composed separately because
 * it is the browser entry's alone: wiring it into `families/table.ts` would put `gen/table/html-dom.ts`
 * in the graph of every Node program that draws a table, and wiring it into the entry directly would
 * make it the one thing a browser tier could not drop.
 */

import { genTableToSlides } from '../gen/table/html-dom.js'
import type { ConstructFamily } from './shared.js'

export const tableDomFamily: ConstructFamily = {
	name: 'table',
	presentationAuthors: {
		tableToSlides(ctx, eleId, options) {
			// @note `options.verbose` (a documented dev-only flag on TableToSlidesProps) is read
			// inside genTableToSlides to trace the auto-paging layout process.
			genTableToSlides(ctx.pres, eleId, options)
		},
	},
}
