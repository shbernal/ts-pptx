/**
 * Resolve a table's `a:tableStyleId` against the deck-wide `tableStyles.xml` and
 * compose the style graph into a per-cell fill.
 *
 * A table cell that defines no own `a:tcPr` fill still renders shaded, because the
 * referenced `a:tblStyle` supplies conditional formatting: `wholeTbl` under
 * `band1H`/`band2H` (row banding), `band1V`/`band2V` (column banding),
 * `firstRow`/`lastRow`, `firstCol`/`lastCol`, and the four corner cells. Resolving
 * a cell means picking, in ECMA-376 precedence order, the highest-priority part
 * that actually defines a fill, then resolving that fill through the slide theme.
 *
 * This is read-only; it does not touch the import/copy path in `table-styles.ts`.
 */

import type { OpcPackage } from '../opc/package.js'
import { attr, firstChild, firstChildElement, getElements, type Element } from '../oxml/dom.js'
import { styleRefFill, type ThemeContext } from '../oxml/theme.js'
import { resolveColorElement, type ResolvedColor } from './theme-context.js'

const TABLE_STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'

/**
 * The `a:tblPr` condition flags that gate which style parts apply to a table.
 * `firstRow`/`bandRow` are the ones PowerPoint turns on by default; the rest are
 * off unless the author enabled the corresponding "Header Column" / "Total Row" /
 * "Banded Columns" toggle.
 */
export interface TableConditionFlags {
	firstRow: boolean
	lastRow: boolean
	firstCol: boolean
	lastCol: boolean
	bandRow: boolean
	bandCol: boolean
}

/** A table's resolved entry in `tableStyles.xml`: its GUID, display name, and raw element. */
export interface ResolvedTableStyle {
	/** The matched `@styleId` GUID (e.g. `{5C22544A-…}`). */
	styleId: string
	/** The style's display name (`@styleName`, e.g. `Medium Style 2 - Accent 1`), or `null`. */
	name: string | null
	/** Escape hatch: the underlying `a:tblStyle` element. After mutating it call {@link ResolvedTableStyle.markDirty}, or `save()` writes the original bytes. */
	element_: Element
	/**
	 * Mark `tableStyles.xml` dirty so `save()` reserializes it. Call after mutating
	 * {@link ResolvedTableStyle.element_}. Note the part is deck-wide: an edit here
	 * reaches every table bound to this style.
	 */
	markDirty(): void
}

/**
 * Look up the `a:tblStyle` whose `@styleId` equals `styleId` in the deck's single
 * `tableStyles.xml`, or `null` when there is no such part or no matching style
 * (a built-in `[MS-OE376]` style the deck does not materialise resolves to `null`).
 */
export function resolveTableStyle(opc: OpcPackage, styleId: string | null): ResolvedTableStyle | null {
	if (!styleId) return null
	const part = opc.partsByContentType(TABLE_STYLES_CONTENT_TYPE)[0]
	const list = part?.dom.documentElement
	if (!part || !list) return null
	for (const st of getElements(list, 'a:tblStyle')) {
		if (attr(st, 'styleId') === styleId) {
			return {
				styleId,
				name: attr(st, 'styleName'),
				element_: st,
				markDirty: () => {
					part.markDirty()
				},
			}
		}
	}
	return null
}

/**
 * The style-part qnames that apply to the cell at `(row, col)`, highest precedence
 * first. Corner cells beat the row/column edge conditions, which beat banding,
 * which beats `wholeTbl`. Banding is suppressed on the header/footer rows and the
 * first/last columns, and its band ordinal is counted from the first body cell
 * (i.e. after a `firstRow` header), so the first body row is `band1H`.
 *
 * Only the `firstRow` / `band1H` / `band2H` / `wholeTbl` path is exercised by a
 * fixture today; the column/corner ordering follows the ECMA-376 conditional-format
 * precedence but is not render-verified.
 */
function cellStyleParts(
	flags: TableConditionFlags,
	row: number,
	col: number,
	rowCount: number,
	colCount: number
): string[] {
	const isFirstRow = flags.firstRow && row === 0
	const isLastRow = flags.lastRow && row === rowCount - 1
	const isFirstCol = flags.firstCol && col === 0
	const isLastCol = flags.lastCol && col === colCount - 1
	const parts: string[] = []
	if (isFirstRow && isFirstCol) parts.push('a:nwCell')
	if (isFirstRow && isLastCol) parts.push('a:neCell')
	if (isLastRow && isFirstCol) parts.push('a:swCell')
	if (isLastRow && isLastCol) parts.push('a:seCell')
	if (isFirstRow) parts.push('a:firstRow')
	if (isLastRow) parts.push('a:lastRow')
	if (isFirstCol) parts.push('a:firstCol')
	if (isLastCol) parts.push('a:lastCol')
	if (flags.bandRow && !isFirstRow && !isLastRow) {
		const ordinal = row - (flags.firstRow ? 1 : 0)
		parts.push(ordinal % 2 === 0 ? 'a:band1H' : 'a:band2H')
	}
	if (flags.bandCol && !isFirstCol && !isLastCol) {
		const ordinal = col - (flags.firstCol ? 1 : 0)
		parts.push(ordinal % 2 === 0 ? 'a:band1V' : 'a:band2V')
	}
	parts.push('a:wholeTbl')
	return parts
}

/**
 * The fill a single `a:tcStyle` contributes. `fall` means the part defines no fill
 * (only borders, or nothing) so resolution continues to the next-lower part;
 * `stop` means the part settles the fill — either a resolved solid colour or `null`
 * for an explicit `a:noFill` / a non-solid style fill (a gradient has no single hex).
 */
function tcStyleFill(
	tcStyle: Element,
	ctx: ThemeContext
): { stop: true; color: ResolvedColor | null } | { stop: false } {
	const fillRef = firstChild(tcStyle, 'a:fillRef')
	if (fillRef) {
		const fill = styleRefFill(fillRef, ctx)
		const color = fill && fill.localName === 'solidFill' ? resolveColorElement(firstChildElement(fill), ctx) : null
		return { stop: true, color }
	}
	const fill = firstChild(tcStyle, 'a:fill')
	if (!fill) return { stop: false }
	const solid = firstChild(fill, 'a:solidFill')
	return { stop: true, color: solid ? resolveColorElement(firstChildElement(solid), ctx) : null }
}

/**
 * The fill the table style contributes to the cell at `(row, col)`, resolved to a
 * literal hex through the slide theme, or `null` when the applicable style parts
 * define no solid fill (an explicit `a:noFill`, e.g. a "Light Style" body cell,
 * also reads `null`). The caller supplies the resolved `a:tblStyle`, the table's
 * condition flags, and the grid dimensions.
 */
export function resolveTableCellStyleFill(
	style: Element,
	flags: TableConditionFlags,
	row: number,
	col: number,
	rowCount: number,
	colCount: number,
	ctx: ThemeContext
): ResolvedColor | null {
	for (const partName of cellStyleParts(flags, row, col, rowCount, colCount)) {
		const part = firstChild(style, partName)
		const tcStyle = part && firstChild(part, 'a:tcStyle')
		if (!tcStyle) continue
		const outcome = tcStyleFill(tcStyle, ctx)
		if (outcome.stop) return outcome.color
	}
	return null
}
