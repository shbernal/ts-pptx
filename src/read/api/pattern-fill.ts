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
import { resolveColorElement, type ResolvedColor } from './theme-context.js'

/**
 * A pattern fill (`a:pattFill`) — a two-colour preset hatch. The write-side
 * `fill: { type: 'pattern', pattern: { preset, fgColor, bgColor } }` emits the same element,
 * so the {@link preset} name and both colours round-trip. Colours resolve against the theme
 * (a scheme token → literal hex) the same way a solid fill does.
 */
export interface PatternFill {
	/** Preset pattern name (`@prst`, e.g. `pct50`/`diagCross`/`ltUpDiag`), or `null` when unset. */
	preset: string | null
	/** Foreground colour (`a:fgClr`) resolved against the theme, or `null`. */
	foreground: ResolvedColor | null
	/** Background colour (`a:bgClr`) resolved against the theme, or `null`. */
	background: ResolvedColor | null
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
	const wrapColor = (qname: string): ResolvedColor | null => {
		const wrap = firstChild(patt, qname)
		const colorEl = wrap && firstChildElement(wrap)
		return colorEl ? resolveColorElement(colorEl, ctx) : null
	}
	return {
		preset: attr(patt, 'prst') ?? null,
		foreground: wrapColor('a:fgClr'),
		background: wrapColor('a:bgClr'),
	}
}
