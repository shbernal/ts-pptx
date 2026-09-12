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
import { colorValueIf } from '../oxml/fill.js'
import type { ColorContext } from '../oxml/theme.js'
import { resolveColorElement, type ResolvedColor } from './theme-context.js'

/**
 * A pattern fill (`a:pattFill`) — a two-colour preset hatch. The write-side
 * `fill: { type: 'pattern', pattern: { preset, fgColor, bgColor } }` emits the same element,
 * so the {@link preset} name and both colours round-trip. Colours resolve against the theme
 * (a scheme token → literal hex) the same way a solid fill does, and a scheme colour's token is
 * reported beside its resolution, the split {@link import('./gradient.js').GradientStop} makes.
 */
export interface PatternFill {
	/** Preset pattern name (`@prst`, e.g. `pct50`/`diagCross`/`ltUpDiag`), or `null` when unset. */
	preset: string | null
	/** Foreground colour (`a:fgClr`) resolved against the theme, or `null`. */
	foreground: ResolvedColor | null
	/** Theme colour token of the foreground (`a:fgClr/a:schemeClr/@val`), or `null` for another colour model. */
	foregroundSchemeColor: string | null
	/** Background colour (`a:bgClr`) resolved against the theme, or `null`. */
	background: ResolvedColor | null
	/** Theme colour token of the background (`a:bgClr/a:schemeClr/@val`), or `null` for another colour model. */
	backgroundSchemeColor: string | null
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
	const colorOf = (qname: string): Element | null => {
		const wrap = firstChild(patt, qname)
		return wrap ? firstChildElement(wrap) : null
	}
	const foreground = colorOf('a:fgClr')
	const background = colorOf('a:bgClr')
	return {
		preset: attr(patt, 'prst') ?? null,
		foreground: foreground ? resolveColorElement(foreground, ctx) : null,
		foregroundSchemeColor: colorValueIf(foreground, 'schemeClr'),
		background: background ? resolveColorElement(background, ctx) : null,
		backgroundSchemeColor: colorValueIf(background, 'schemeClr'),
	}
}
