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
	PresSlide,
	TableCell,
	TableToSlidesProps,
	TableCellProps,
} from '../../core-interfaces.js'
import type { SlideLayoutInternal } from '../../types/internal.js'
import { rgbToHex } from '../drawingml/color.js'
import { inch2Emu } from '../../units-internal.js'
import { warn } from '../../log.js'
import { EMU_PER_INCH } from '../../units.js'
import { getSlidesForTableRows } from './autopage.js'

type MarginTuple = [number, number, number, number]
type BorderTuple = [BorderProps, BorderProps, BorderProps, BorderProps]
type TableToSlidesHost = {
	addSlide: (options?: AddSlideProps) => PresSlide
	presLayout: PresLayout
}

// ===== DOM-table helpers (browser path) =====
// DOM-independent helpers factored out of the browser-only tableToSlides() flow
// so they can be unit-tested without a rendered page (see AGENTS.md scope note).

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

// ===== Live-DOM tableToSlides() flow =====
// Browser-only entry (reads a rendered table via getComputedStyle/offsetWidth);
// out of active scope — see AGENTS.md. Internally signposted by `// STEP 1..5`.

/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {TableToSlidesHost} pptx - ts-pptx instance
 * @param {string} tabEleId - HTMLElementID of the table
 * @param {TableToSlidesProps} options - array of options (e.g.: tabsize)
 * @param {SlideLayoutInternal} masterSlide - masterSlide
 */
export function genTableToSlides(
	pptx: TableToSlidesHost,
	tabEleId: string,
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
	let arrInchMargins: [number, number, number, number] = [0.5, 0.5, 0.5, 0.5] // TRBL-style
	let intTabW = 0

	// REALITY-CHECK:
	if (!document.getElementById(tabEleId)) throw new Error('tableToSlides: Table ID "' + tabEleId + '" does not exist!')

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
	let firstRowCells = document.querySelectorAll(`#${tabEleId} tr:first-child th`)
	if (firstRowCells.length === 0) firstRowCells = document.querySelectorAll(`#${tabEleId} tr:first-child td`)
	firstRowCells.forEach((cellEle: Element) => {
		const cell = cellEle as HTMLTableCellElement
		if (cell.getAttribute('colspan')) {
			// Guesstimate (divide evenly) col widths
			// NOTE: both j$query and vanilla selectors return {0} when table is not visible)
			for (let idxc = 0; idxc < Number(cell.getAttribute('colspan')); idxc++) {
				arrTabColW.push(Math.round(cell.offsetWidth / Number(cell.getAttribute('colspan'))))
			}
		} else {
			arrTabColW.push(cell.offsetWidth)
		}
	})
	arrTabColW.forEach((colW) => {
		intTabW += colW
	})

	// STEP 3: Calc/Set column widths by using same column width percent from HTML table
	arrTabColW.forEach((colW, idxW) => {
		const intCalcWidth = Number(((Number(emuSlideTabW) * ((colW / intTabW) * 100)) / 100 / EMU_PER_INCH).toFixed(2))
		const headCell = document.querySelector(`#${tabEleId} thead tr:first-child th:nth-child(${idxW + 1})`)
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
		document.querySelectorAll(`#${tabEleId} ${part} tr`).forEach((row: Element) => {
			const htmlRow = row as HTMLTableRowElement
			const arrObjTabCells: TableCell[] = []
			Array.from(htmlRow.cells).forEach((cell) => {
				// A: Get RGB text/bkgd colors
				const arrRGB1 = window
					.getComputedStyle(cell)
					.getPropertyValue('color')
					.replace(/\s+/gi, '')
					.replace('rgba(', '')
					.replace('rgb(', '')
					.replace(')', '')
					.split(',')
				let arrRGB2 = window
					.getComputedStyle(cell)
					.getPropertyValue('background-color')
					.replace(/\s+/gi, '')
					.replace('rgba(', '')
					.replace('rgb(', '')
					.replace(')', '')
					.split(',')
				if (
					// NOTE: Default for unstyled tables is black bkgd, so use white instead
					window.getComputedStyle(cell).getPropertyValue('background-color') === 'rgba(0, 0, 0, 0)' ||
					window.getComputedStyle(cell).getPropertyValue('transparent')
				) {
					arrRGB2 = ['255', '255', '255']
				}

				// B: Create option object
				const cellOpts: TableCellProps = {
					bold: !!(
						window.getComputedStyle(cell).getPropertyValue('font-weight') === 'bold' ||
						Number(window.getComputedStyle(cell).getPropertyValue('font-weight')) >= 500
					),
					color: rgbToHex(Number(arrRGB1[0]), Number(arrRGB1[1]), Number(arrRGB1[2])),
					fill: { color: rgbToHex(Number(arrRGB2[0]), Number(arrRGB2[1]), Number(arrRGB2[2])) },
					fontSize: Number(window.getComputedStyle(cell).getPropertyValue('font-size').replace(/[a-z]/gi, '')),
				}
				const fontFace = ((window.getComputedStyle(cell).getPropertyValue('font-family') || '').split(',')[0] ?? '')
					.replace(/"/g, '')
					.replace('inherit', '')
					.replace('initial', '')
				const colspan = Number(cell.getAttribute('colspan')) || undefined
				const rowspan = Number(cell.getAttribute('rowspan')) || undefined
				if (fontFace) cellOpts.fontFace = fontFace
				if (colspan) cellOpts.colspan = colspan
				if (rowspan) cellOpts.rowspan = rowspan

				if (
					['left', 'center', 'right', 'start', 'end'].includes(
						window.getComputedStyle(cell).getPropertyValue('text-align')
					)
				) {
					const align = window
						.getComputedStyle(cell)
						.getPropertyValue('text-align')
						.replace('start', 'left')
						.replace('end', 'right')
					cellOpts.align =
						align === 'center' ? 'center' : align === 'left' ? 'left' : align === 'right' ? 'right' : undefined
				}
				if (['top', 'middle', 'bottom'].includes(window.getComputedStyle(cell).getPropertyValue('vertical-align'))) {
					const valign = window.getComputedStyle(cell).getPropertyValue('vertical-align')
					cellOpts.valign =
						valign === 'top' ? 'top' : valign === 'middle' ? 'middle' : valign === 'bottom' ? 'bottom' : undefined
				}

				// C: Add padding [margin] (if any)
				// NOTE: Margins translate: px->pt 1:1 (e.g.: a 20px padded cell looks the same in PPTX as 20pt Text Inset/Padding)
				if (window.getComputedStyle(cell).getPropertyValue('padding-left')) {
					const cellMargin: MarginTuple = [0, 0, 0, 0]
					const sidesPad = ['padding-top', 'padding-right', 'padding-bottom', 'padding-left']
					sidesPad.forEach((val, idxs) => {
						cellMargin[idxs] = Math.round(
							Number(window.getComputedStyle(cell).getPropertyValue(val).replace(/\D/gi, ''))
						)
					})
					cellOpts.margin = cellMargin
				}

				// D: Add border (if any)
				if (
					window.getComputedStyle(cell).getPropertyValue('border-top-width') ||
					window.getComputedStyle(cell).getPropertyValue('border-right-width') ||
					window.getComputedStyle(cell).getPropertyValue('border-bottom-width') ||
					window.getComputedStyle(cell).getPropertyValue('border-left-width')
				) {
					const cellBorder: BorderTuple = [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }]
					const sidesBor = ['top', 'right', 'bottom', 'left']
					sidesBor.forEach((val, idxb) => {
						const style = window.getComputedStyle(cell)
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
					text: cell.innerText, // `innerText` returns <br> as "\n", so linebreak etc. work later!
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
		const newSlide = pptx.addSlide({ masterTitle: opts.masterSlideName || undefined })

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
