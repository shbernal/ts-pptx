/**
 * Shared line reader for DrawingML strokes: a shape's `p:spPr/a:ln`, a table cell's edges
 * (`a:tcPr/a:lnL` and its siblings, each a `CT_LineProperties`), and a chart series'
 * `c:spPr/a:ln`.
 *
 * Each of the three decoded the width, the dash preset, the solid colour and an explicit no-line
 * on its own, the drift {@link import('./gradient.js').readGradientFill} and
 * {@link import('./pattern-fill.js').readPatternFill} are shared to prevent. What each reports from
 * the decode stays its own: a chart series' `color` is the raw `a:srgbClr`, because a chart part is
 * read without a theme, and a cell border's is the resolved hex.
 */
import { attr, firstChild, numberValue, type Element } from '../oxml/dom.js'
import { solidFillColor } from '../oxml/fill.js'
import type { ColorContext } from '../oxml/theme.js'
import { ptFromEmu } from './coords.js'
import { resolveSolidFillColor, type ResolvedColor } from './theme-context.js'

/** What every reader of an `a:ln` decodes from it. */
export interface LineBasics {
	/** Stroke width in points (`@w` is EMU; 12700 EMU = 1pt), or `null` when unset. */
	widthPt: number | null
	/** Dash preset (`a:prstDash/@val`, e.g. `sysDash`), or `null` when unset. */
	dash: string | null
	/** Literal stroke colour (`a:solidFill/a:srgbClr/@val`), or `null` for another colour model or no solid fill. */
	color: string | null
	/** Theme colour token (`a:solidFill/a:schemeClr/@val`), or `null` for another colour model or no solid fill. */
	schemeColor: string | null
	/**
	 * The solid stroke colour resolved against the theme, or `null` when the line has no solid fill,
	 * the colour cannot be made literal, or no theme was given to resolve it against.
	 */
	resolvedColor: ResolvedColor | null
	/** `true` when the line is explicitly suppressed (`a:noFill`). */
	noFill: boolean
}

/**
 * Decode a line element's width, dash, solid colour and no-fill flag.
 * @param ln - the `CT_LineProperties` element
 * @param ctx - the theme the colour resolves against, or `null` to leave {@link LineBasics.resolvedColor} unset
 */
export function readLineBasics(ln: Element, ctx: ColorContext | null): LineBasics {
	const dash = firstChild(ln, 'a:prstDash')
	return {
		widthPt: ptFromEmu(numberValue(attr(ln, 'w'))),
		dash: dash ? (attr(dash, 'val') ?? null) : null,
		color: solidFillColor(ln, 'a:srgbClr'),
		schemeColor: solidFillColor(ln, 'a:schemeClr'),
		resolvedColor: ctx ? resolveSolidFillColor(ln, ctx) : null,
		noFill: !!firstChild(ln, 'a:noFill'),
	}
}
