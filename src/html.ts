/**
 * `ts-pptx/html` — reproduce an existing HTML `<table>` as a PowerPoint table, as a free
 * function that works anywhere there is a DOM.
 *
 * The same conversion is available as `TsPptx.prototype.tableToSlides(id, options)` on the
 * browser build. Reach for *that* when you are in a browser and have a table id; reach for
 * *this* when either of those is not true:
 *
 * - **You are not in a browser.** Node has no global `document`, so the method's entry point
 *   does not exist. Pass the element itself (from any DOM implementation) and no global is
 *   consulted at all; pass an id together with `options.document` and the id resolves against
 *   that document.
 * - **You already have the element.** Passing it skips the id lookup, and skips the class of
 *   bug where an id that is not a valid CSS identifier resolves to nothing.
 * - **Your presentation is not a `TsPptx`.** The `pptx` parameter is structural
 *   ({@link TableToSlidesHost}: `addSlide` + `presLayout`), so anything shaped like a
 *   presentation works, including the Node build.
 *
 * **Column widths need a layout engine.** In a browser the columns are sized from each cell's
 * rendered `offsetWidth`, which reproduces the table's real proportions. Nothing outside a
 * browser lays a table out, so `offsetWidth` is `0` there and the conversion degrades: it uses
 * the computed CSS `width`s when the stylesheet states them for every column in one unit, and
 * an equal split when it does not. To pin widths regardless, put `data-pptx-width` (exact
 * inches) or `data-pptx-min-width` (a floor) on the `<thead>` header cells — those win outright
 * on every path.
 *
 * Everything else — cell text (with `<br>` preserved as a line break), colspan/rowspan,
 * computed colors, weight, alignment, padding, borders, and auto-paging across as many slides
 * as the rows need — works the same wherever it runs.
 * @example
 * ```ts
 * import { TsPptx } from '@shbernal/ts-pptx'
 * import { tableToSlides } from '@shbernal/ts-pptx/html'
 * import { Window } from 'happy-dom'
 *
 * const win = new Window()
 * win.document.body.innerHTML = '<table id="report">…</table>'
 *
 * const pptx = new TsPptx()
 * tableToSlides(pptx, win.document.getElementById('report'), { autoPage: true })
 * await pptx.writeFile({ fileName: 'report.pptx' })
 * ```
 */

import { genTableToSlides, type TableToSlidesHost } from './gen/table/html-dom.js'
import type { TableToSlidesElement, TableToSlidesProps } from './core-interfaces.js'

/**
 * Reproduce an HTML `<table>` as a PowerPoint table, adding as many slides as its rows need.
 * @param {TableToSlidesHost} pptx - the presentation to add slides to
 * @param {TableToSlidesElement | string} table - the `<table>` element, or its id (see `options.document`)
 * @param {TableToSlidesProps} options - generation options
 */
export function tableToSlides(
	pptx: TableToSlidesHost,
	table: TableToSlidesElement | string,
	options: TableToSlidesProps = {}
): void {
	genTableToSlides(pptx, table, options)
}

export type { TableToSlidesHost }
export type { TableToSlidesDocument, TableToSlidesElement, TableToSlidesProps } from './core-interfaces.js'
