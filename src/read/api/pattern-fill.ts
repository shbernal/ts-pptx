/**
 * Shared pattern-fill reader for DrawingML colour-bearing containers: a shape's `p:spPr`, a
 * slide's `p:bg/p:bgPr`, and a table's `a:tblPr` / cell's `a:tcPr`.
 *
 * `a:pattFill` is a preset hatch name plus two wrapped colours, and decoding it is three
 * lines — which is exactly why it had been open-coded in each place that needed it. Sharing
 * it is what keeps the three from drifting into different answers for the same element, the
 * same reason {@link import('./gradient.js').readGradientFill} and
 * {@link import('./picture-fill.js').readPictureFill} are shared.
 */
import { attr, firstChild, firstChildElement, type Element } from '../oxml/dom.js'
import type { ColorContext } from '../oxml/theme.js'
import { readColorRef, type ColorRef } from './theme-context.js'

/**
 * A pattern fill (`a:pattFill`) — a two-colour preset hatch. The write-side
 * `fill: { type: 'pattern', pattern: { preset, fgColor, bgColor } }` emits the same element,
 * so the {@link preset} name and both colours round-trip.
 */
export interface PatternFill {
	/** Preset pattern name (`@prst`, e.g. `pct50`/`diagCross`/`ltUpDiag`), or `null` when unset. */
	preset: string | null
	/** Foreground colour (`a:fgClr`); every field is `null` when the pattern states none. */
	foreground: ColorRef
	/** Background colour (`a:bgClr`); every field is `null` when the pattern states none. */
	background: ColorRef
}

/**
 * Read a container's `a:pattFill`, or `null` when its fill is not a pattern.
 * @param {Element} container - the fill-bearing parent (`p:spPr`, `p:bgPr`, `a:tblPr`, `a:tcPr`, …)
 * @param {ColorContext} ctx - the theme colour context each colour resolves against
 * @returns {PatternFill | null} the decoded hatch, or `null`
 */
export function readPatternFill(container: Element, ctx: ColorContext): PatternFill | null {
	const patt = firstChild(container, 'a:pattFill')
	if (!patt) return null
	// `a:fgClr` and `a:bgClr` wrap their colour element, unlike `a:highlight`'s bare child.
	const colorOf = (qname: string): ColorRef => {
		const wrap = firstChild(patt, qname)
		return readColorRef(wrap ? firstChildElement(wrap) : null, ctx)
	}
	return {
		preset: attr(patt, 'prst') ?? null,
		foreground: colorOf('a:fgClr'),
		background: colorOf('a:bgClr'),
	}
}
