/**
 * ts-pptx: HTML-table → slides (browser DOM path)
 *
 * The tableToSlides() flow: reproduce an HTML table as a PowerPoint table across as many
 * slides as needed. The DOM-independent parts are factored into pure helpers
 * (`resolveHtmlColWidth`, `pickColWidthBasis`, `htmlBorderToProps`, `cssColorToHex`) that
 * are unit-tested directly, and the whole flow runs anywhere there is a DOM.
 *
 * Only real *measurement* is out of active scope (see AGENTS.md): without a layout engine
 * `offsetWidth` is `0`, so widths fall back to computed CSS then to an equal split. A fallback
 * rather than a degradation, and the distinction is load-bearing — see {@link pickColWidthBasis},
 * whose two bases measure different boxes.
 *
 * **Coverage.** This file used to be excluded from the report, on the grounds that only the
 * browser entry imported it and the `dist/browser*.js` globs therefore swallowed it. Both
 * halves of that are stale: `ts-pptx/html` imports it too, so tsdown emits it as its own
 * `dist/html-dom-*.js` chunk, and those globs are gone from `vitest.config.ts` entirely. It
 * is covered code now — against happy-dom by the Node suite
 * (test/regression/html-to-slides-node.test.js) and, for the measured width basis that no
 * Node DOM can produce, in a real Chromium (test/browser/table-widths.spec.mjs).
 */

import { SlideObjectType } from '../../enums.js'
import type {
	AddSlideProps,
	BorderProps,
	PresLayout,
	TableCell,
	TableToSlidesDocument,
	TableToSlidesElement,
	TableToSlidesProps,
	TableCellProps,
} from '../../types/index.js'
import type { Slide } from '../../types/slide.js'
import type { SlideLayoutInternal } from '../../types/internal.js'
import { inch2Emu } from '../../units-internal.js'
import { warn } from '../../diagnostics.js'
import { DEFAULT_PX_PER_INCH, EMU_PER_INCH } from '../../units.js'
import { getSlidesForTableRows } from './autopage.js'
import { InvalidOptionError } from '../../errors.js'

type MarginTuple = [number, number, number, number]
type BorderTuple = [BorderProps, BorderProps, BorderProps, BorderProps]
/**
 * What `tableToSlides` needs a presentation to be: somewhere to put slides, and a layout to
 * size them against. Structural on purpose — stating the two members rather than naming the
 * class is what lets the free function on `ts-pptx/html` take any presentation instance
 * (Node, browser, or standalone) without the entry importing one.
 */
export type TableToSlidesHost = {
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
 *
 * This is the one place the public structural element type ({@link TableToSlidesElement},
 * which any DOM satisfies) is narrowed to `lib.dom`'s `Element`, so the rest of the flow can
 * be written against ordinary DOM types. The cast is safe in the only sense that matters here:
 * everything downstream reads standard members and tolerates them being absent — that is what
 * makes the conversion portable in the first place.
 * @param {TableToSlidesElement | string} target - the table element, or the id of one
 * @param {TableToSlidesProps} options - generation options (may carry an explicit `document`)
 * @returns {DomContext} resolved DOM context
 */
function resolveDomContext(target: TableToSlidesElement | string, options: TableToSlidesProps): DomContext {
	const asElement = (element: TableToSlidesElement): Element => element as unknown as Element

	if (typeof target !== 'string') {
		const table = asElement(target)
		return { table, getComputedStyle: resolveStyleReader(table.ownerDocument) }
	}

	const doc = options.document ?? (globalThis as { document?: TableToSlidesDocument }).document
	if (!doc) {
		throw new InvalidOptionError(
			'html/no-document',
			'tableToSlides: no DOM available to resolve the table id "' +
				target +
				'" — pass the <table> element itself, or supply a document via `options.document`.'
		)
	}
	const found = doc.getElementById(target)
	if (!found)
		throw new InvalidOptionError('html/table-not-found', 'tableToSlides: Table ID "' + target + '" does not exist!')
	const table = asElement(found)
	return { table, getComputedStyle: resolveStyleReader(table.ownerDocument) }
}

// ===== DOM-table helpers (browser path) =====
// DOM-independent helpers factored out of the browser-only tableToSlides() flow
// so they can be unit-tested without a rendered page (see AGENTS.md scope note).

/**
 * Convert one 0-255 RGB component to its two-digit hex form.
 *
 * Lives here rather than beside the other color code in `gen/drawingml/color.ts` because
 * `getComputedStyle` colors are the only thing in the package that arrives as RGB
 * components — nothing else in the library calls this.
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
 * `0pt` and silently vanish — the table serializer (`ptsToEmuLenient`) emits fractional points just
 * fine, so there is no reason to integer-round here. A
 * computed width of `0` (or a non-finite value) yields `{ type: 'none' }` so we never emit a
 * zero-width line.
 * @param {string} widthStr - computed `border-<side>-width`, e.g. `"0.5px"`
 * @param {string} colorStr - computed `border-<side>-color`, e.g. `"rgb(102, 102, 102)"`
 * @returns {BorderProps} border props for the cell side
 */
export function htmlBorderToProps(widthStr: string, colorStr: string): BorderProps {
	const pt = Number(String(widthStr).replace('px', ''))
	if (!Number.isFinite(pt) || pt <= 0) return { type: 'none' }
	return { width: pt, color: cssColorToHex(colorStr) ?? '000000' }
}

/**
 * Convert a computed CSS color to a pptx hex color, or report that it has none.
 *
 * A browser normalizes every computed color to `rgb()`/`rgba()`, which is why the flow below
 * originally only parsed that form. Outside a browser there is no normalization step, so a
 * computed color comes back however the stylesheet wrote it — `#ff0000` is what happy-dom
 * returns. Feeding that to the `rgb()` parser produced `Number('#ff0000')` → `NaN` → the
 * literal color string `"NANNANNAN"`, an invalid value emitted without complaint.
 *
 * Handles `rgb()`/`rgba()` and `#rgb`/`#rrggbb`. Everything else — a named color, a `color()`
 * function, an empty string — returns `undefined` rather than a guess, so the caller can
 * apply its own default instead of emitting nonsense.
 *
 * A fully transparent color (`alpha === 0`) also returns `undefined`: it states the absence of
 * a color, and pptx has no way to say "transparent" in a solid fill.
 * @param {string} value - computed CSS color, e.g. `"rgb(255, 0, 0)"` or `"#ff0000"`
 * @returns {string | undefined} six-digit uppercase hex, or `undefined` when unparseable
 */
export function cssColorToHex(value: string): string | undefined {
	const raw = String(value ?? '').trim()

	const rgbMatch = /^rgba?\(([^)]*)\)$/i.exec(raw)
	if (rgbMatch) {
		const parts = (rgbMatch[1] ?? '').split(/[,/\s]+/).filter((part) => part.length > 0)
		if (parts.length < 3) return undefined
		const channels = parts.slice(0, 3).map((part) => Number(part.replace('%', '')))
		if (channels.some((channel) => !Number.isFinite(channel))) return undefined
		if (parts.length > 3 && Number(parts[3]) === 0) return undefined
		// Clamp and round before hex conversion: a browser always computes whole 0-255
		// channels, but nothing guarantees another DOM does, and `(255.5).toString(16)`
		// is `"ff.8"` — a malformed color rather than a wrong one.
		const [r = 0, g = 0, b = 0] = channels.map((channel) => Math.max(0, Math.min(255, Math.round(channel))))
		return rgbToHex(r, g, b)
	}

	const hexMatch = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(raw)
	if (hexMatch) {
		const digits = hexMatch[1] ?? ''
		const full =
			digits.length === 3
				? digits
						.split('')
						.map((digit) => digit + digit)
						.join('')
				: digits
		return full.toUpperCase()
	}

	return undefined
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
	const safeCalc = Number.isFinite(calcWidth) ? calcWidth : 0
	if (Number.isFinite(setWidth) && setWidth > 0) return setWidth
	return Number.isFinite(minWidth) && minWidth > safeCalc ? minWidth : safeCalc
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
		if (!Number.isFinite(magnitude) || magnitude < 0) return []
		const valueUnit = (match[2] ?? 'px').toLowerCase()
		if (!unit) unit = valueUnit
		else if (unit !== valueUnit) return []
		basis.push(magnitude)
	}
	return basis
}

/**
 * Choose which vector the proportional column math runs on, falling back as bases run out.
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
 *
 * **The first two bases do not measure the same box, and the docs used to imply they did.**
 * `offsetWidth` is the border box; a computed `width` is the content box. Padding alone
 * separates them — `test/browser/harness/table.html` is built so it does, giving 1:1 measured
 * against 2:1 from CSS on one table — so choosing arm 2 over arm 1 can change a table's column
 * *proportions*, not just their precision. That is why the fixture discriminates at all, and why
 * "degrades to computed CSS" was the wrong word for it everywhere it appeared.
 *
 * Left as-is rather than reconciled. Normalizing arm 2 to the border box would need the computed
 * padding and border widths, which the DOMs that reach arm 2 need not resolve either (a `%`
 * padding computes to nothing usable without layout), so it would converge the two only
 * sometimes — and it would collapse the one discriminator the browser lane has, turning
 * `table-widths.spec.mjs` back into a test that passes whether or not the measured arm ran.
 * A caller who needs both runtimes to agree on a column states it with `data-pptx-width`.
 * @param {readonly number[]} measured - per-column `offsetWidth`
 * @param {readonly number[]} cssWidths - per-column CSS width basis (`[]` when unusable)
 * @returns {number[]} the basis vector to run the proportional calc on
 */
export function pickColWidthBasis(measured: readonly number[], cssWidths: readonly number[]): number[] {
	const sum = (arr: readonly number[]): number => arr.reduce((acc, n) => acc + (Number.isFinite(n) ? n : 0), 0)
	if (sum(measured) > 0) return [...measured]
	if (cssWidths.length === measured.length && sum(cssWidths) > 0) return [...cssWidths]
	return measured.map(() => 1)
}

/** One cell's span attributes, as far as the grid walk cares. */
type GridCellSpans = { colspan?: number; rowspan?: number }

/**
 * Read a `colspan`/`rowspan` as the number of grid tracks it covers.
 *
 * Lenient on purpose: `0`, a negative, a fraction, an unparseable attribute and an absent one all
 * mean `1`. A bad span attribute should cost its own cell nothing and the rest of the grid nothing
 * — the alternative, propagating a `NaN` or a negative into the column walk, shifts every
 * subsequent column.
 * @param {unknown} value - the raw attribute value
 * @returns {number} tracks covered, at least 1
 */
function gridSpan(value: unknown): number {
	const num = Math.floor(Number(value))
	return Number.isFinite(num) && num > 1 ? num : 1
}

/**
 * Measure the column grid a set of table rows actually occupies.
 *
 * An HTML table is not a rectangle of cells — a row states only the cells it *starts*, so its
 * length says nothing about how many grid columns it fills: a `colspan` fills several, and a
 * `rowspan` started higher up fills one without the row mentioning it. The pptx side has no such
 * model: `<a:tblGrid>` declares N columns and every `<a:tr>` must carry N `<a:tc>`. Something has
 * to translate, and that is a walk of the standard HTML table model — cells pack left to right,
 * skipping columns still held by a rowspan above.
 *
 * Returns the grid width and, per row, how many columns that row reaches. The difference between
 * the two is exactly the padding a ragged row needs; the merge cells for `colspan`/`rowspan` are
 * *not* padding — the table emitter synthesizes those itself, so they are counted as filled here.
 *
 * Spans are read through {@link gridSpan}, so a bad attribute cannot shift the whole grid.
 * @param {readonly (readonly GridCellSpans[])[]} rows - per-row cell spans, in emission order
 * @returns {{ columns: number, filled: number[] }} grid width, and columns reached per row
 */
export function measureGridColumns(rows: readonly (readonly GridCellSpans[])[]): { columns: number; filled: number[] } {
	// carry[c] = how many further rows column c is still held by a rowspan started above.
	const carry: number[] = []
	const filled: number[] = []

	for (const row of rows) {
		let col = 0
		for (const cell of row) {
			while ((carry[col] ?? 0) > 0) col++
			const colspan = gridSpan(cell?.colspan)
			// Written now, decremented at the end of this row, so it leaves `rowspan - 1` behind.
			for (let idx = 0; idx < colspan; idx++) carry[col + idx] = gridSpan(cell?.rowspan)
			col += colspan
		}
		// Trailing columns held from above count too: a row whose last cells are all rowspan
		// continuations reaches past its own final cell.
		while ((carry[col] ?? 0) > 0) col++
		filled.push(col)
		for (let idx = 0; idx < carry.length; idx++) {
			const held = carry[idx] ?? 0
			if (held > 0) carry[idx] = held - 1
		}
	}

	return { columns: filled.reduce((max, reached) => (reached > max ? reached : max), 0), filled }
}

/**
 * Widen a column-width basis to cover a grid wider than the row it was derived from.
 *
 * The basis comes from one row — the first with cells — but the grid is as wide as the *widest*
 * row, and those need not agree (a single spanning header over a wider body is the common shape).
 * The columns the source row never described get the average of the ones it did, so they read as
 * ordinary columns rather than as a zero-width sliver; with no usable basis at all they get `1`,
 * which the proportional calc downstream turns into an equal split.
 * @param {readonly number[]} basis - per-column basis derived from the width-source row
 * @param {number} columns - the full grid width
 * @returns {number[]} basis covering every grid column
 */
export function extendColBasis(basis: readonly number[], columns: number): number[] {
	const extended = [...basis]
	const total = extended.reduce((acc, width) => acc + (Number.isFinite(width) ? width : 0), 0)
	const share = extended.length > 0 && total > 0 ? total / extended.length : 1
	while (extended.length < columns) extended.push(share)
	return extended
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
 * `innerText` is the right answer in a browser — it is the *rendered* text, so it already
 * collapses whitespace, honors `display: none`, and turns `<br>` into `"\n"`. But it is an
 * HTML-spec extra tied to rendering, and outside a browser it is either absent (jsdom does not
 * implement it, so every cell would come out empty) or present but not actually rendered
 * (happy-dom returns `"ab"` for `a<br>b`, so every cell would come out on one line).
 *
 * So both are computed and the better one wins: a `childNodes` walk that maps `<br>` to a line
 * break and collapses whitespace per line, and `innerText`. `innerText` is preferred unless it
 * dropped a line break the walk found — the one thing that provably makes it not the rendered
 * text. Neither can know about `text-transform` without a browser.
 * @param {TextCell} cell - the table cell element
 * @returns {string} cell text, `<br>`-separated lines joined by `"\n"`
 */
export function readCellText(cell: TextCell): string {
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
	const walked = lines.map((line) => line.replace(/\s+/g, ' ').trim()).join('\n')

	// Prefer `innerText` — except when the walk found line breaks it dropped. A DOM that does
	// not render can still expose an `innerText` property that is really just `textContent`
	// (happy-dom returns `"ab"` for `a<br>b`), and that one is not the rendered text this
	// wants: it has silently lost the cell's line structure. Where it kept the breaks, it is
	// the better answer, because it also knows about `display:none` and `text-transform`.
	const rendered = cell.innerText
	if (typeof rendered !== 'string') return walked
	return walked.includes('\n') && !rendered.includes('\n') ? walked : rendered
}

// ===== Live-DOM tableToSlides() flow =====
// Browser-only entry (reads a rendered table via getComputedStyle/offsetWidth);
// out of active scope — see AGENTS.md. Internally signposted by `// STEP 1..5`.

/**
 * Resolve `options.masterTitle` to the slide layout it names, so the auto-pager can take that
 * master's margins.
 *
 * The layout registry is a presentation *internal* — it is not on {@link TableToSlidesHost},
 * and the public `slideLayouts` getter deliberately hides the `_name` this matches on. Rather
 * than widen either, read it defensively: a host that has no registry simply resolves no
 * master, which is what a caller passing a bare `{ addSlide, presLayout }` object should get.
 * The alternative — dropping the lookup on the free-function path — would make `masterTitle`
 * silently do nothing there, which is a trap.
 * @param {TableToSlidesHost} pptx - the presentation the slides are being added to
 * @param {string} masterTitle - the `title` passed to `defineSlideMaster`
 * @returns {SlideLayoutInternal | undefined} the named layout, if this host has one
 */
function resolveMasterSlide(pptx: TableToSlidesHost, masterTitle?: string): SlideLayoutInternal | undefined {
	if (!masterTitle) return undefined
	const layouts = (pptx as { _slideLayouts?: SlideLayoutInternal[] })._slideLayouts
	if (!Array.isArray(layouts)) return undefined
	return layouts.find((layout) => layout._name === masterTitle)
}

/** Which of the three logical sections a row belongs to. A section-less row counts as body. */
type TablePart = 'thead' | 'tbody' | 'tfoot'

/**
 * A row's cells, via `HTMLTableRowElement.cells`.
 *
 * Read off the structural shape rather than the `lib.dom` type so a DOM that does not implement
 * the collection degrades to "this row has no cells" instead of throwing — the same tolerance
 * every other DOM read in this file has.
 * @param {Element} row - a `<tr>`
 * @returns {HTMLTableCellElement[]} the row's cells, in document order
 */
function rowCells(row: Element): HTMLTableCellElement[] {
	const cells = (row as { cells?: ArrayLike<Element> }).cells
	return cells ? (Array.from(cells) as HTMLTableCellElement[]) : []
}

/**
 * Collect every row of the table, bucketed head → body → foot.
 *
 * `HTMLTableElement.rows` is the right source and a `thead tr`/`tbody tr`/`tfoot tr` query is not,
 * for two reasons the query cannot fix:
 *
 * - **A row need not be in a section.** `<table><tr>…` is valid authored markup, and a table built
 *   with `createElement`/`appendChild` has no `<tbody>` at all — the HTML *parser* inserts one, the
 *   DOM API does not. Those rows matched none of the three selectors and were silently dropped,
 *   which is how a perfectly good table reached `addTable` with zero rows and threw
 *   `addTable: Array expected!` (upstream gitbrent/PptxGenJS#1005).
 * - **A descendant query reaches into nested tables.** `tbody tr` matches the rows of a table
 *   inside a cell of this one, folding them into the outer table as extra rows.
 *
 * Each row's own parentage settles both questions: it names the section the row belongs to (the
 * head/body/foot split still matters downstream, for repeating header rows on auto-paged slides),
 * and it says whether the row is this table's at all. The ownership check is not redundant with
 * using `rows` — happy-dom's `rows` is itself a descendant walk, so a nested table's rows arrive
 * there too. Checking is what makes the answer the same on every DOM.
 * @param {Element} table - the `<table>` being converted
 * @returns {Array<{ row: Element, part: TablePart }>} rows in emission order
 */
function collectTableRows(table: Element): Array<{ row: Element; part: TablePart }> {
	const head: Array<{ row: Element; part: TablePart }> = []
	const body: Array<{ row: Element; part: TablePart }> = []
	const foot: Array<{ row: Element; part: TablePart }> = []
	const place = (row: Element): void => {
		const parent = row.parentElement
		const section = (parent?.nodeName ?? '').toUpperCase()
		const sectioned = section === 'THEAD' || section === 'TBODY' || section === 'TFOOT'
		if ((sectioned ? parent?.parentElement : parent) !== table) return
		if (section === 'THEAD') head.push({ row, part: 'thead' })
		else if (section === 'TFOOT') foot.push({ row, part: 'tfoot' })
		else body.push({ row, part: 'tbody' })
	}

	const rows = (table as { rows?: ArrayLike<Element> }).rows
	if (rows) {
		for (let idx = 0; idx < rows.length; idx++) {
			const row = rows[idx]
			if (row) place(row)
		}
	} else {
		// A DOM without `rows` is not one this library has ever run on, but the section query is
		// still a truthful fallback for the sectioned tables it can see.
		for (const part of ['thead', 'tbody', 'tfoot']) {
			table.querySelectorAll(`${part} tr`).forEach(place)
		}
	}

	return [...head, ...body, ...foot]
}

/**
 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
 * @param {TableToSlidesHost} pptx - ts-pptx instance
 * @param {TableToSlidesElement | string} target - the table element, or the HTMLElementID of the table
 * @param {TableToSlidesProps} options - array of options (e.g.: tabsize)
 */
export function genTableToSlides(
	pptx: TableToSlidesHost,
	target: TableToSlidesElement | string,
	options: TableToSlidesProps = {}
): void {
	const opts = options || {}
	const masterSlide = resolveMasterSlide(pptx, opts.masterTitle)
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
		else if (Number.isFinite(masterSlide._margin))
			arrInchMargins = [masterSlide._margin, masterSlide._margin, masterSlide._margin, masterSlide._margin]
		opts.slideMargin = arrInchMargins
	} else if (opts?.slideMargin) {
		if (Array.isArray(opts.slideMargin)) arrInchMargins = opts.slideMargin
		else if (Number.isFinite(opts.slideMargin))
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

	// STEP 2: Resolve the rows and the grid they occupy, then take the column widths from the
	// first row that has any cells (others may have colspans, who cares, we only need col widths from 1).
	// Two bases are collected in the same pass: the rendered `offsetWidth` (what a browser
	// reports) and the computed CSS `width` (the only width statement available where nothing
	// laid the table out). `pickColWidthBasis` decides which one STEP 3 runs on.
	const domRows = collectTableRows(ctx.table)
	const { columns: intColCnt, filled: arrRowFilled } = measureGridColumns(
		domRows.map(({ row }) =>
			rowCells(row).map((cell) => ({
				colspan: Number(cell.getAttribute('colspan')),
				rowspan: Number(cell.getAttribute('rowspan')),
			}))
		)
	)
	// A table with no rows — or only empty ones — has nothing to convert. Say so here: the failure
	// used to surface far downstream as `Reduce of empty array with no initial value` out of the
	// auto-pager, or as `addTable: Array expected!`, neither of which names the table.
	const srcRow = domRows.find(({ row }) => rowCells(row).length > 0)
	if (!srcRow) {
		throw new InvalidOptionError(
			'html/table-has-no-cells',
			'tableToSlides: the table has no cells to convert - expected at least one <tr> with a <td>/<th>.'
		)
	}

	// Per grid column, the width-source cell it came from and that cell's colspan. Built here so
	// STEP 3 can read a column's overrides off the *right* cell: with a colspan in play, column
	// index and cell index part ways, and indexing the row by column silently applied one cell's
	// `data-pptx-width` to the next column along (upstream gitbrent/PptxGenJS#1244).
	const arrColSrc: Array<{ cell: HTMLTableCellElement; span: number }> = []
	const arrCellSpans: number[] = []
	rowCells(srcRow.row).forEach((cell) => {
		const offsetW = Number(cell.offsetWidth)
		const measured = Number.isFinite(offsetW) ? offsetW : 0
		arrTabColCssW.push(ctx.getComputedStyle(cell).getPropertyValue('width'))
		// Guesstimate (divide evenly) col widths across a spanned cell.
		// NOTE: both j$query and vanilla selectors return {0} when table is not visible)
		const span = gridSpan(cell.getAttribute('colspan'))
		arrCellSpans.push(span)
		for (let idxc = 0; idxc < span; idxc++) {
			arrTabColW.push(Math.round(measured / span))
			arrColSrc.push({ cell, span })
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

	// The width-source row can be narrower than the grid (one spanning header over a wider body),
	// so widen the basis to the full grid before the proportional calc — that way the calc still
	// divides the slide across *every* column instead of overflowing it by the missing ones.
	const arrBasisColW = extendColBasis(pickColWidthBasis(arrTabColW, arrCssColW), intColCnt)
	arrBasisColW.forEach((colW) => {
		intTabW += colW
	})

	// STEP 3: Calc/Set column widths by using same column width percent from HTML table
	arrBasisColW.forEach((colW, idxW) => {
		const intCalcWidth = Number(((Number(emuSlideTabW) * ((colW / intTabW) * 100)) / 100 / EMU_PER_INCH).toFixed(2))
		// A `data-pptx-width` on a spanning cell states that *cell's* width, so it divides across
		// the columns the cell covers — exactly as its `offsetWidth` and its CSS width do above.
		// A column past the source row's reach has no cell to be overridden by.
		const src = arrColSrc[idxW]
		const intSetWidth = src ? Number(src.cell.getAttribute('data-pptx-width')) / src.span : 0
		const intMinWidth = src ? Number(src.cell.getAttribute('data-pptx-min-width')) / src.span : 0
		arrColW.push(resolveHtmlColWidth(intCalcWidth, intSetWidth, intMinWidth))
	})
	if (opts.verbose) {
		console.log(`| arrColW ......................................... = [${arrColW.join(', ')}]`)
	}

	// STEP 4: Iterate over each table element and create data arrays (text and opts)
	// NOTE: We create 3 arrays instead of one so we can loop over body then show header/footer rows on first and last page
	domRows.forEach(({ row, part }, idxRow) => {
		const arrObjTabCells: TableCell[] = []
		rowCells(row).forEach((cell) => {
			// The computed style is read a dozen times below; resolve it once per cell.
			const style = ctx.getComputedStyle(cell)

			// A: Get text/bkgd colors
			// NOTE: an unparseable or fully transparent background is the default for an
			// unstyled table; pptx has no "transparent" solid fill, so use white instead.
			const textColor = cssColorToHex(style.getPropertyValue('color')) ?? '000000'
			const fillColor = cssColorToHex(style.getPropertyValue('background-color')) ?? 'FFFFFF'

			// B: Create option object
			const cellOpts: TableCellProps = {
				bold: !!(
					style.getPropertyValue('font-weight') === 'bold' || Number(style.getPropertyValue('font-weight')) >= 500
				),
				color: textColor,
				fill: { color: fillColor },
				fontSize: Number(style.getPropertyValue('font-size').replace(/[a-z]/gi, '')),
			}
			const fontFace = ((style.getPropertyValue('font-family') || '').split(',')[0] ?? '')
				.replace(/"/g, '')
				.replace('inherit', '')
				.replace('initial', '')
			// Read through the same `gridSpan` the grid walk used, so what is emitted and what
			// was measured cannot disagree: a `colspan="-2"` counted as one column above must
			// not reach the table definition as `-2` and corrupt its own column count.
			const colspan = gridSpan(cell.getAttribute('colspan'))
			const rowspan = gridSpan(cell.getAttribute('rowspan'))
			if (fontFace) cellOpts.fontFace = fontFace
			if (colspan > 1) cellOpts.colspan = colspan
			if (rowspan > 1) cellOpts.rowspan = rowspan

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
			// NOTE: `TableCellProps.margin` is INCHES, so computed px must be converted, not copied.
			// CSS defines the reference pixel as 1/96in, and mirroring what the browser rendered is
			// this conversion's whole job — so resolve at `DEFAULT_PX_PER_INCH`, the same density the
			// `"<n>px"` coordinate unit uses. (Until 2026-07-31 the px magnitude was assigned straight
			// through under a stale "px->pt 1:1" note left from when cell margin was points; that made
			// a 4px pad a *4 inch* inset and tripped the legacy-points warning on any pad >= 1px.)
			// No rounding to whole px: the browser's computed value can be fractional, and `inch2Emu`
			// already does the one rounding that matters, at EMU precision.
			if (style.getPropertyValue('padding-left')) {
				const cellMargin: MarginTuple = [0, 0, 0, 0]
				const sidesPad = ['padding-top', 'padding-right', 'padding-bottom', 'padding-left']
				sidesPad.forEach((val, idxs) => {
					// Anything that is not an absolute px length (a `%` padding, a keyword) has no
					// meaning as a fixed inset, so it insets by nothing rather than by its digits.
					const pad = parseCssPx(style.getPropertyValue(val))
					cellMargin[idxs] = Number.isFinite(pad) ? pad / DEFAULT_PX_PER_INCH : 0
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

		// Normalize a ragged row up to the grid width. HTML tolerates a short row and just leaves
		// the tail of the grid empty; pptx does not — `<a:tblGrid>` declares the column count and a
		// row with fewer `<a:tc>` is a malformed table PowerPoint has to repair. The missing cells
		// are blank rather than a copy of a neighbour: the source table never stated them, so there
		// is nothing to carry over. Columns held by a `rowspan` from above are already counted as
		// filled — the table emitter synthesizes those merge cells itself.
		for (let idxPad = arrRowFilled[idxRow] ?? 0; idxPad < intColCnt; idxPad++) {
			arrObjTabCells.push({ _type: SlideObjectType.tablecell, text: '', options: {} })
		}

		if (part === 'thead') arrObjTabHeadRows.push(arrObjTabCells)
		else if (part === 'tfoot') arrObjTabFootRows.push(arrObjTabCells)
		else arrObjTabBodyRows.push(arrObjTabCells)
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
				warn('html/image-missing-source', 'tableToSlides.addImage requires either `path` or `data`')
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
