/**
 * Shared line reader for DrawingML strokes: a shape's `p:spPr/a:ln`, a table cell's edges
 * (`a:tcPr/a:lnL` and its siblings, each a `CT_LineProperties`), and a chart series'
 * `c:spPr/a:ln`.
 *
 * Each of the three decoded the width, the dash preset, the solid colour and an explicit no-line
 * on its own, the drift `readGradientFill` and
 * `readPatternFill` are shared to prevent.
 */
import { attr, firstChild, firstChildElement, numberValue, type Element } from '../oxml/dom.js'
import type { ColorContext } from '../oxml/theme.js'
import { ptFromEmu } from './coords.js'
import { readColorRef, type ColorRef } from './theme-context.js'

/** What every reader of an `a:ln` decodes from it. */
export interface LineBasics {
	/** Stroke width in points (`@w` is EMU; 12700 EMU = 1pt), or `null` when unset. */
	widthPt: number | null
	/** Dash preset (`a:prstDash/@val`, e.g. `sysDash`), or `null` when unset. */
	dash: string | null
	/**
	 * The line's solid colour (`a:solidFill`); every field is `null` when it has none. `resolved` is
	 * also `null` when no theme was given to resolve it against.
	 */
	colorRef: ColorRef
	/** `true` when the line is explicitly suppressed (`a:noFill`). */
	noFill: boolean
}

/**
 * Decode a line element's width, dash, solid colour and no-fill flag.
 * @param ln - the `CT_LineProperties` element
 * @param ctx - the theme the colour resolves against, or `null` to leave `colorRef.resolved` unset
 */
export function readLineBasics(ln: Element, ctx: ColorContext | null): LineBasics {
	const dash = firstChild(ln, 'a:prstDash')
	const solidFill = firstChild(ln, 'a:solidFill')
	return {
		widthPt: ptFromEmu(numberValue(attr(ln, 'w'))),
		dash: dash ? (attr(dash, 'val') ?? null) : null,
		colorRef: readColorRef(solidFill ? firstChildElement(solidFill) : null, ctx),
		noFill: !!firstChild(ln, 'a:noFill'),
	}
}
