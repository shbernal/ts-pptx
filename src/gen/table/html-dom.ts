/**
 * ts-pptx: HTML-table → slides (browser DOM path)
 *
 * The browser-only tableToSlides() flow: reproduce a rendered HTML table as a
 * PowerPoint table across as many slides as needed. Out of active scope (see
 * AGENTS.md) — the DOM-independent parts are factored into pure helpers
 * (`resolveHtmlColWidth`, `htmlBorderToProps`) that are unit-tested directly.
 *
 * This module is imported only by the browser entry (`browser.ts`, which adds the
 * `tableToSlides` method), so it bundles into the browser/standalone chunks — never
 * the Node build or the shared core chunk. That is why the live-DOM code below needs
 * no `v8 ignore` fence: the whole file is coverage-excluded via the `dist/browser*.js`
 * chunk globs (see vitest.config.ts), while the pure helpers stay unit-tested from src.
 */

import { SlideObjectType } from '../../core-enums.js'
import type {
	AddSlideProps,
	BorderProps,
	PresLayout,
	TableCell,
	TableToSlidesProps,
	TableCellProps,
} from '../../core-interfaces.js'
import type { Slide } from '../../types/slide.js'
import type { SlideLayoutInternal } from '../../types/internal.js'
import { inch2Emu } from '../../units-internal.js'
import { warn } from '../../log.js'
import { EMU_PER_INCH } from '../../units.js'
import { getSlidesForTableRows } from './autopage.js'

type MarginTuple = [number, number, number, number]
type BorderTuple = [BorderProps, BorderProps, BorderProps, BorderProps]
type TableToSlidesHost = {
	addSlide: (options?: AddSlideProps) => Slide
	presLayout: PresLayout
}

/**
 * The whole DOM surface `genTableToSlides` needs: the table element, and a way to read a
 * cell's computed style.
 *
 * Resolving this once up front is what decouples the flow from `window`/`document`. Given the
 * element itself, both halves come from the element's own `ownerDocument`/`defaultView`, so no
 * global is read at all — which is what lets the same implementation run outside a browser.
 */
type DomContext = {
	table: Element
	getComputedStyle: (el: Element) => CSSStyleDeclaration
}

/**
 * Stand-in for `getComputedStyle` on a document that has no view — a document produced by
 * parsing rather than by a browsing context has a `null` `defaultView` and therefore no
 * cascade to compute against. Every property reads as `''`, which every consumer below
 * already treats as "not set", so styling degrades to defaults instead of throwing.
 */
const NO_COMPUTED_STYLE = {
	getPropertyValue: () => '',
} as unknown as CSSStyleDeclaration

/**
 * Bind a style reader to the document's own view, falling back to {@link NO_COMPUTED_STYLE}.
 * @param {Document | null} doc - the table's owner document
 * @returns {(el: Element) => CSSStyleDeclaration} computed-style reader
 */
function resolveStyleReader(doc: Document | null): (el: Element) => CSSStyleDeclaration {
	const view = doc?.defaultView
	return view ? view.getComputedStyle.bind(view) : () => NO_COMPUTED_STYLE
}

/**
 * Resolve the table element and its style reader from whatever the caller passed.
 *
 * An **element** needs no ambient DOM: the document and its view are reached through
 * `ownerDocument`/`defaultView`. A **string id** does need one, taken from `options.document`
 * when supplied and otherwise from `globalThis.document`; with neither, the error names both
 * remedies rather than surfacing a bare `document is not defined`.
 * @param {Element | string} target - the table element, or the id of one
 * @param {TableToSlidesProps} options - generation options (may carry an explicit `document`)
 * @returns {DomContext} resolved DOM context
 */
function resolveDomContext(target: Element | string, options: TableToSlidesProps): DomContext {
	if (typeof target !== 'string') return { table: target, getComputedStyle: resolveStyleReader(target.ownerDocument) }

	const doc = options.document ?? (globalThis as { document?: Document }).document
	if (!doc) {
		throw new Error(
			'tableToSlides: no DOM available to resolve the table id "' +
				target +
				'" — pass the <table> element itself, or supply a document via `options.document`.'
		)
	}
	const table = doc.getElementById(target)
	if (!table) throw new Error('tableToSlides: Table ID "' + target + '" does not exist!')
	return { table, getComputedStyle: resolveStyleReader(table.ownerDocument) }
}

// ===== DOM-table helpers (browser path) =====
// DOM-independent helpers factored out of the browser-only tableToSlides() flow
// so they can be unit-tested without a rendered page (see AGENTS.md scope note).

/**
 * Convert one 0-255 RGB component to its two-digit hex form.
 *
 * Lives here rather than beside the other color code in `gen/drawingml/color.ts`
 * because `getComputedStyle` colors are the only thing in the package that arrives
 * as RGB components — nothing on the Node path calls this. Keeping it in the shared
 * color module put two functions the Node chunk can never execute into that chunk's
 * coverage report; here they fall under the `dist/browser*.js` exclusion that already
 * covers this whole file, so no `v8 ignore` fence is needed.
 * @param {number} c - component color
 * @returns {string} hex string
 */
function componentToHex(c: number): string {
	const hex = c.toString(16)
	return hex.length === 1 ? '0' + hex : hex
}

/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
function rgbToHex(r: number, g: number, b: number): string {
	return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase()
}

/**
 * Convert a computed CSS border (width string + color string) from `getComputedStyle` into a
 * pptx `BorderProps`.
 *
 * Preserves *fractional* widths: a hairline CSS border such as `0.5px` must not be rounded to
 * `0pt` and silently vanish — the table serializer (`valToPts`) emits fractional points just
 * fine, so there is no reason to integer-round here. A
 * computed width of `0` (or a non-finite value) yields `{ type: 'none' }` so we never emit a
 * zero-width line.
 * @param {string} widthStr - computed `border-<side>-width`, e.g. `"0.5px"`
 * @param {string} colorStr - computed `border-<side>-color`, e.g. `"rgb(102, 102, 102)"`
 * @returns {BorderProps} border props for the cell side
 */
export function htmlBorderToProps(widthStr: string, colorStr: string): BorderProps {
	const pt = Number(String(widthStr).replace('px', ''))
	if (!isFinite(pt) || pt <= 0) return { type: 'none' }
	const arrRGB = String(colorStr)
		.replace(/\s+/gi, '')
		.replace('rgba(', '')
		.replace('rgb(', '')
		.replace(')', '')
		.split(',')
	return { width: pt, color: rgbToHex(Number(arrRGB[0]), Number(arrRGB[1]), Number(arrRGB[2])) }
}

/**
 * Resolve a single HTML-table column width for `tableToSlides`.
 *
 * Precedence: an explicit `data-pptx-width` wins outright; otherwise the proportional width
 * derived from the live table is used, raised to `data-pptx-min-width` when that floor is larger.
 *
 * Hidden tables report `offsetWidth` 0 for every cell, which makes `calcWidth` non-finite (a 0/0
 * proportional calc). Fall back to `0` there so an explicit `data-pptx-width` / `data-pptx-min-width`
 * override still drives the column instead of emitting a `NaN` width.
 * @param {number} calcWidth - proportional width derived from `offsetWidth` (may be `NaN` for hidden tables)
 * @param {number} setWidth - `data-pptx-width` override (`0`/`NaN` when absent or invalid)
 * @param {number} minWidth - `data-pptx-min-width` floor (`0`/`NaN` when absent or invalid)
 * @returns {number} resolved column width
 */
export function resolveHtmlColWidth(calcWidth: number, setWidth: number, minWidth: number): number {
	const safeCalc = isFinite(calcWidth) ? calcWidth : 0
	if (isFinite(setWidth) && setWidth > 0) return setWidth
	return isFinite(minWidth) && minWidth > safeCalc ? minWidth : safeCalc
}

/** A CSS value that is a bare number, or a number with a `px`/`%` unit. Nothing else parses. */
const CSS_LENGTH = /^\s*(-?(?:\d+\.?\d*|\.\d+))\s*(px|%)?\s*$/i

/**
 * Parse a single computed CSS length in **px** to its numeric magnitude.
 *
 * A bare number is read as px (computed styles do not emit one, but an inline `style` read
 * back by a non-browser DOM can). Everything else — `''`, `auto`, `3em`, and deliberately
 * `30%` — is `NaN`: a percentage is meaningful only relative to something, so it cannot be
 * turned into an absolute length here. (Percentages *are* usable as a proportional column
 * basis; that is {@link parseCssWidthBasis}'s job, not this one's.) Callers decide what a
 * `NaN` means for them.
 * @param {string} value - computed CSS value, e.g. `"1.5px"`
 * @returns {number} magnitude in px, or `NaN` when the value is not an absolute px length
 */
export function parseCssPx(value: string): number {
	const match = CSS_LENGTH.exec(String(value ?? ''))
	if (!match || match[2] === '%') return NaN
	return Number(match[1])
}

/**
 * Parse a whole row of computed CSS widths into a column basis, or reject the set.
 *
 * The result is only ever used *proportionally* (each column's share of the total), so the
 * unit does not have to be absolute — it only has to be the same for every column. Hence:
 *
 * - all `px` (or bare numbers) → their magnitudes
 * - all `%` → their magnitudes; `[25%, 25%, 50%]` is a perfectly good 1:1:2 basis
 * - anything else, or a **mix** of units, or a single unparseable entry (`auto`, `''`,
 *   `3em`) → `[]`, meaning "no usable basis"
 *
 * Rejecting the whole set on one bad entry is deliberate: a partial basis would silently
 * give the unparseable columns zero width, which is worse than falling back to an equal
 * split. Negative magnitudes reject for the same reason.
 * @param {readonly string[]} values - computed `width` per column, in column order
 * @returns {number[]} per-column basis, or `[]` when the set is not usable
 */
export function parseCssWidthBasis(values: readonly string[]): number[] {
	const basis: number[] = []
	let unit = ''
	for (const value of values) {
		const match = CSS_LENGTH.exec(String(value ?? ''))
		if (!match) return []
		const magnitude = Number(match[1])
		if (!isFinite(magnitude) || magnitude < 0) return []
		const valueUnit = (match[2] ?? 'px').toLowerCase()
		if (!unit) unit = valueUnit
		else if (unit !== valueUnit) return []
		basis.push(magnitude)
	}
	return basis
}

/**
 * Choose which vector the proportional column math runs on, degrading gracefully.
 *
 * `offsetWidth` is the ideal basis — it is the width the table actually rendered at — but it
 * requires a layout engine, and reports `0` for every cell both in a hidden table and on any
 * DOM outside a browser. So:
 *
 * 1. **measured** (`offsetWidth`) when it carries any width at all — the browser path, unchanged
 * 2. **CSS widths** when the stylesheet states them (see {@link parseCssWidthBasis})
 * 3. **equal split** otherwise — a basis of all-ones, which the existing proportional math
 *    turns into equal columns for free
 *
 * Only the basis changes; `data-pptx-width` / `data-pptx-min-width` overrides still apply
 * downstream via {@link resolveHtmlColWidth} and still win outright.
 * @param {readonly number[]} measured - per-column `offsetWidth`
 * @param {readonly number[]} cssWidths - per-column CSS width basis (`[]` when unusable)
 * @returns {number[]} the basis vector to run the proportional calc on
 */
export function pickColWidthBasis(measured: readonly number[], cssWidths: readonly number[]): number[] {
	const sum = (arr: readonly number[]): number => arr.reduce((acc, n) => acc + (isFinite(n) ? n : 0), 0)
	if (sum(measured) > 0) return [...measured]
	if (cssWidths.length === measured.length && sum(cssWidths) > 0) return [...cssWidths]
	return measured.map(() => 1)
}

/**
 * The minimum a cell has to look like for {@link readCellText}. Any DOM's element satisfies it;
 * stating it structurally is what lets the fallback walk be unit-tested without a DOM at all.
 */
type TextCell = {
	innerText?: string
	nodeType?: number
	nodeValue?: string | null
	nodeName?: string
	childNodes?: ArrayLike<TextCell>
}

/**
 * Read a cell's text, with `<br>` surviving as a newline.
 *
 * `innerText` is the right answer and is what the browser path uses — it is the *rendered*
 * text, so it already collapses whitespace and turns `<br>` into `"\n"`. But it is an
 * HTML-spec extra that not every DOM implements (jsdom notably does not), and on those the
 * property is simply absent, which would silently empty every cell.
 *
 * The fallback walks `childNodes`, concatenating text and mapping `<br>` to `"\n"`, then
 * collapses whitespace runs per line and trims — an approximation of `innerText` for a DOM
 * with no rendering to ask. It cannot know about `display: none` or `text-transform`; a
 * caller that needs those needs a real browser.
 * @param {TextCell} cell - the table cell element
 * @returns {string} cell text, `<br>`-separated lines joined by `"\n"`
 */
export function readCellText(cell: TextCell): string {
	if (typeof cell.innerText === 'string') return cell.innerText

	// Accumulate one bucket per line rather than one string with "\n" separators: the source's
	// own indentation contains newlines, and only a `<br>` may actually start a new line. A
	// separator character could not tell those two apart; separate buckets never have to.
	const lines: string[] = ['']
	const append = (text: string): void => {
		lines[lines.length - 1] = (lines[lines.length - 1] ?? '') + text
	}
	const walk = (node: TextCell): void => {
		if (node.nodeType === 3) {
			append(node.nodeValue ?? '')
			return
		}
		if (node.nodeType !== 1) return
		if ((node.nodeName ?? '').toUpperCase() === 'BR') {
			lines.push('')
			return
		}
		const children = node.childNodes
		for (let idx = 0; idx < (children?.length ?? 0); idx++) {
			const child = children?.[idx]
			if (child) walk(child)
		}
	}

	// Walk from the cell's children: the cell itself is an element, and can never be a BR.
	const children = cell.childNodes
	for (let idx = 0; idx < (children?.length ?? 0); idx++) {
		const child = children?.[idx]
		if (child) walk(child)
	}
	return lines.map((line) => line.replace(/\s+/g, ' ').trim()).join('\n')
}

// ===== Live-DOM tableToSlides() flow =====
// Browser-only entry (reads a rendered table via getComputedStyle/offsetWidth);
// out of active scope — see AGENTS.md. Internally signposted by `// STEP 1..5`.

/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {TableToSlidesHost} pptx - ts-pptx instance
 * @param {Element | string} target - the table element, or the HTMLElementID of the table
 * @param {TableToSlidesProps} options - array of options (e.g.: tabsize)
 * @param {SlideLayoutInternal} masterSlide - masterSlide
 */
export function genTableToSlides(
	pptx: TableToSlidesHost,
	target: Element | string,
	options: TableToSlidesProps = {},
	masterSlide?: SlideLayoutInternal
): void {
	const opts = options || {}
	opts.slideMargin = opts.slideMargin || opts.slideMargin === 0 ? opts.slideMargin : 0.5
	let emuSlideTabW = opts.w || pptx.presLayout.width
	const arrObjTabHeadRows: TableCell[][] = []
	const arrObjTabBodyRows: TableCell[][] = []
	const arrObjTabFootRows: TableCell[][] = []
	const arrColW: number[] = []
	const arrTabColW: number[] = []
	const arrTabColCssW: string[] = []
	let arrInchMargins: [number, number, number, number] = [0.5, 0.5, 0.5, 0.5] // TRBL-style
	let intTabW = 0

	// REALITY-CHECK: resolve the table (and its style reader) before anything else — this is
	// also where a missing/unresolvable table id fails.
	const ctx = resolveDomContext(target, opts)

	// STEP 1: Set margins
	if (masterSlide?._margin) {
		if (Array.isArray(masterSlide._margin)) arrInchMargins = masterSlide._margin
		else if (!isNaN(masterSlide._margin))
			arrInchMargins = [masterSlide._margin, masterSlide._margin, masterSlide._margin, masterSlide._margin]
		opts.slideMargin = arrInchMargins
	} else if (opts?.slideMargin) {
		if (Array.isArray(opts.slideMargin)) arrInchMargins = opts.slideMargin
		else if (!isNaN(opts.slideMargin))
			arrInchMargins = [opts.slideMargin, opts.slideMargin, opts.slideMargin, opts.slideMargin]
	}
	emuSlideTabW = (opts.w ? inch2Emu(opts.w) : pptx.presLayout.width) - inch2Emu(arrInchMargins[1] + arrInchMargins[3])

	if (opts.verbose) {
		console.log('[[VERBOSE MODE]]')
		console.log('|-- `tableToSlides` ----------------------------------------------------|')
		console.log(`| tableProps.h .................................... = ${opts.h}`)
		console.log(`| tableProps.w .................................... = ${opts.w}`)
		console.log(
			`| pptx.presLayout.width ........................... = ${(pptx.presLayout.width / EMU_PER_INCH).toFixed(1)}`
		)
		console.log(
			`| pptx.presLayout.height .......................... = ${(pptx.presLayout.height / EMU_PER_INCH).toFixed(1)}`
		)
		console.log(`| emuSlideTabW .................................... = ${(emuSlideTabW / EMU_PER_INCH).toFixed(1)}`)
	}

	// STEP 2: Grab table col widths - just find the first availble row, either thead/tbody/tfoot, others may have colspans, who cares, we only need col widths from 1
	// Two bases are collected in the same pass: the rendered `offsetWidth` (what a browser
	// reports) and the computed CSS `width` (the only width statement available where nothing
	// laid the table out). `pickColWidthBasis` decides which one STEP 3 runs on.
	let firstRowCells = ctx.table.querySelectorAll('tr:first-child th')
	if (firstRowCells.length === 0) firstRowCells = ctx.table.querySelectorAll('tr:first-child td')
	const arrCellSpans: number[] = []
	firstRowCells.forEach((cellEle: Element) => {
		const cell = cellEle as HTMLTableCellElement
		const offsetW = Number(cell.offsetWidth)
		const measured = isFinite(offsetW) ? offsetW : 0
		arrTabColCssW.push(ctx.getComputedStyle(cell).getPropertyValue('width'))
		if (cell.getAttribute('colspan')) {
			// Guesstimate (divide evenly) col widths
			// NOTE: both j$query and vanilla selectors return {0} when table is not visible)
			const span = Number(cell.getAttribute('colspan'))
			arrCellSpans.push(span)
			for (let idxc = 0; idxc < span; idxc++) {
				arrTabColW.push(Math.round(measured / span))
			}
		} else {
			arrCellSpans.push(1)
			arrTabColW.push(measured)
		}
	})

	// The CSS basis is parsed per *cell* (unit uniformity is a property of the row as authored),
	// then expanded across each cell's colspan the same way `offsetWidth` is.
	const parsedCellCssW = parseCssWidthBasis(arrTabColCssW)
	const arrCssColW: number[] = []
	if (parsedCellCssW.length === arrCellSpans.length) {
		parsedCellCssW.forEach((cssW, idxC) => {
			const span = arrCellSpans[idxC] ?? 1
			for (let idxc = 0; idxc < span; idxc++) arrCssColW.push(cssW / span)
		})
	}

	const arrBasisColW = pickColWidthBasis(arrTabColW, arrCssColW)
	arrBasisColW.forEach((colW) => {
		intTabW += colW
	})

	// STEP 3: Calc/Set column widths by using same column width percent from HTML table
	arrBasisColW.forEach((colW, idxW) => {
		const intCalcWidth = Number(((Number(emuSlideTabW) * ((colW / intTabW) * 100)) / 100 / EMU_PER_INCH).toFixed(2))
		const headCell = ctx.table.querySelector(`thead tr:first-child th:nth-child(${idxW + 1})`)
		const intSetWidth = headCell ? Number(headCell.getAttribute('data-pptx-width')) : 0
		const intMinWidth = headCell ? Number(headCell.getAttribute('data-pptx-min-width')) : 0
		arrColW.push(resolveHtmlColWidth(intCalcWidth, intSetWidth, intMinWidth))
	})
	if (opts.verbose) {
		console.log(`| arrColW ......................................... = [${arrColW.join(', ')}]`)
	}

	// STEP 4: Iterate over each table element and create data arrays (text and opts)
	// NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
	const tableParts = ['thead', 'tbody', 'tfoot']
	tableParts.forEach((part) => {
		ctx.table.querySelectorAll(`${part} tr`).forEach((row: Element) => {
			const htmlRow = row as HTMLTableRowElement
			const arrObjTabCells: TableCell[] = []
			Array.from(htmlRow.cells).forEach((cell) => {
				// The computed style is read a dozen times below; resolve it once per cell.
				const style = ctx.getComputedStyle(cell)

				// A: Get RGB text/bkgd colors
				const arrRGB1 = style
					.getPropertyValue('color')
					.replace(/\s+/gi, '')
					.replace('rgba(', '')
					.replace('rgb(', '')
					.replace(')', '')
					.split(',')
				let arrRGB2 = style
					.getPropertyValue('background-color')
					.replace(/\s+/gi, '')
					.replace('rgba(', '')
					.replace('rgb(', '')
					.replace(')', '')
					.split(',')
				if (
					// NOTE: Default for unstyled tables is black bkgd, so use white instead
					style.getPropertyValue('background-color') === 'rgba(0, 0, 0, 0)'
				) {
					arrRGB2 = ['255', '255', '255']
				}

				// B: Create option object
				const cellOpts: TableCellProps = {
					bold: !!(
						style.getPropertyValue('font-weight') === 'bold' || Number(style.getPropertyValue('font-weight')) >= 500
					),
					color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
					fill: { color: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])) },
					fontSize: Number(style.getPropertyValue('font-size').replace(/[a-z]/gi, '')),
				}
				const fontFace = ((style.getPropertyValue('font-family') || '').split(',')[0] ?? '')
					.replace(/"/g, '')
					.replace('inherit', '')
					.replace('initial', '')
				const colspan = Number(cell.getAttribute('colspan')) || undefined
				const rowspan = Number(cell.getAttribute('rowspan')) || undefined
				if (fontFace) cellOpts.fontFace = fontFace
				if (colspan) cellOpts.colspan = colspan
				if (rowspan) cellOpts.rowspan = rowspan

				if (['left', 'center', 'right', 'start', 'end'].includes(style.getPropertyValue('text-align'))) {
					const align = style.getPropertyValue('text-align').replace('start', 'left').replace('end', 'right')
					cellOpts.align =
						align === 'center' ? 'center' : align === 'left' ? 'left' : align === 'right' ? 'right' : undefined
				}
				if (['top', 'middle', 'bottom'].includes(style.getPropertyValue('vertical-align'))) {
					const valign = style.getPropertyValue('vertical-align')
					cellOpts.valign =
						valign === 'top' ? 'top' : valign === 'middle' ? 'middle' : valign === 'bottom' ? 'bottom' : undefined
				}

				// C: Add padding [margin] (if any)
				// NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
				if (style.getPropertyValue('padding-left')) {
					const cellMargin: MarginTuple = [0, 0, 0, 0]
					const sidesPad = ['padding-top', 'padding-right', 'padding-bottom', 'padding-left']
					sidesPad.forEach((val, idxs) => {
						// Anything that is not an absolute px length (a `%` padding, a keyword) has no
						// meaning as a point inset, so it insets by nothing rather than by its digits.
						const pad = parseCssPx(style.getPropertyValue(val))
						cellMargin[idxs] = isFinite(pad) ? Math.round(pad) : 0
					})
					cellOpts.margin = cellMargin
				}

				// D: Add border (if any)
				if (
					style.getPropertyValue('border-top-width') ||
					style.getPropertyValue('border-right-width') ||
					style.getPropertyValue('border-bottom-width') ||
					style.getPropertyValue('border-left-width')
				) {
					const cellBorder: BorderTuple = [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }]
					const sidesBor = ['top', 'right', 'bottom', 'left']
					sidesBor.forEach((val, idxb) => {
						cellBorder[idxb] = htmlBorderToProps(
							style.getPropertyValue('border-' + val + '-width'),
							style.getPropertyValue('border-' + val + '-color')
						)
					})
					cellOpts.border = cellBorder
				}

				// LAST: Add cell
				arrObjTabCells.push({
					_type: SlideObjectType.tablecell,
					text: readCellText(cell), // <br> must survive as "\n", so linebreak etc. work later!
					options: cellOpts,
				})
			})
			switch (part) {
				case 'thead':
					arrObjTabHeadRows.push(arrObjTabCells)
					break
				case 'tbody':
					arrObjTabBodyRows.push(arrObjTabCells)
					break
				case 'tfoot':
					arrObjTabFootRows.push(arrObjTabCells)
					break
				default:
					console.log(`table parsing: unexpected table part: ${part}`)
					break
			}
		})
	})

	// STEP 5: Break table into Slides as needed
	// Pass head-rows as there is an option to add to each table and the parse func needs this data to fulfill that option
	opts._arrObjTabHeadRows = arrObjTabHeadRows
	opts.colW = arrColW
	getSlidesForTableRows(
		[...arrObjTabHeadRows, ...arrObjTabBodyRows, ...arrObjTabFootRows],
		opts,
		pptx.presLayout,
		masterSlide
	).forEach((slide, idxTr) => {
		// A: Create new Slide
		const newSlide = pptx.addSlide({ masterTitle: opts.masterTitle || undefined })

		// B: DESIGN: Reset `y` to startY or margin after first Slide
		if (idxTr === 0) opts.y = opts.y || arrInchMargins[0]
		if (idxTr > 0) opts.y = opts.autoPageSlideStartY || arrInchMargins[0]
		if (opts.verbose)
			console.log(
				`| opts.autoPageSlideStartY: ${opts.autoPageSlideStartY} / arrInchMargins[0]: ${arrInchMargins[0]} => opts.y = ${opts.y}`
			)

		// C: Add table to Slide
		newSlide.addTable(slide.rows, {
			x: opts.x || arrInchMargins[3],
			y: opts.y,
			w: Number(emuSlideTabW) / EMU_PER_INCH,
			colW: arrColW,
			autoPage: false,
		})

		// D: Add any additional objects
		if (opts.addImage) {
			opts.addImage.options = opts.addImage.options || {}
			if (!opts.addImage.image || (!opts.addImage.image.path && !opts.addImage.image.data)) {
				warn('tableToSlides.addImage requires either `path` or `data`')
			} else {
				const imageProps = opts.addImage.image.path
					? { path: opts.addImage.image.path }
					: { data: opts.addImage.image.data as string }
				newSlide.addImage({
					...imageProps,
					x: opts.addImage.options.x,
					y: opts.addImage.options.y,
					w: opts.addImage.options.w,
					h: opts.addImage.options.h,
				})
			}
		}
		if (opts.addShape) newSlide.addShape(opts.addShape.shapeName, opts.addShape.options || {})
		if (opts.addTable) newSlide.addTable(opts.addTable.rows, opts.addTable.options || {})
		if (opts.addText) newSlide.addText(opts.addText.text, opts.addText.options || {})
	})
}
