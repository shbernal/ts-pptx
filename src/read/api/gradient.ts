/**
 * Shared gradient-fill reader for DrawingML colour-bearing containers: a shape's
 * `p:spPr`, a line's `a:ln`, and a slide's `p:bg/p:bgPr`. Decodes an `a:gradFill`
 * into its stops (colour + position) plus the linear/path geometry, resolving each
 * stop's colour against the slide theme.
 */
import { attr, firstChild, firstChildElement, getElements, numberValue, pctAttr, type Element } from '../oxml/dom.js'
import type { ColorContext } from '../oxml/theme.js'
import { resolveColorElement, type ResolvedColor } from './theme-context.js'
import { ANGLE_UNITS_PER_DEGREE } from '../../units.js'

/** One stop of a gradient fill (`a:gsLst/a:gs`), as read from a shape. */
export interface GradientStop {
	/** Stop offset along the gradient, 0–1 (from `@pos`, thousandths of a percent), or `null` if unset. */
	position: number | null
	/** Explicit RGB colour as 6-hex (`a:srgbClr/@val`), or `null` when the stop uses another colour model. */
	color: string | null
	/** Theme colour token (`a:schemeClr/@val`, e.g. `accent1`), or `null` when the stop uses another colour model. */
	schemeColor: string | null
	/**
	 * Preset colour name (`a:prstClr/@val`, e.g. `black`/`cornflowerBlue` — the
	 * ECMA-376 §20.1.10.47 table), or `null` when the stop uses another colour
	 * model. The same raw/resolved split {@link import('./shapes/types.js').RecolorColor}
	 * uses; {@link import('../oxml/preset-color.js').presetColorHex}, exported from
	 * `ts-pptx/read`, makes one literal.
	 *
	 * These three raw fields name three of the five colour models the reader
	 * resolves. A stop written as `a:sysClr` or `a:hslClr` leaves all three `null`
	 * and is reported through {@link resolvedColor} alone, which is the only place
	 * its colour appears; `a:scrgbClr`, which the reader deliberately does not
	 * resolve, leaves {@link resolvedColor} `null` too.
	 */
	presetColor: string | null
	/**
	 * The stop's colour as a full {@link ResolvedColor} — base `hex`, the raw
	 * `transforms` list (`lumMod`/`shade`/…) in document order, and the
	 * `effectiveHex`/`alpha` after applying them — exactly what a solid fill's
	 * `resolvedFill` gives. `null` when the colour cannot be made literal (an
	 * unmapped token, or a colour model we do not resolve).
	 *
	 * This is the field to read when re-authoring a stop against a *different*
	 * theme: {@link effectiveHex} alone is one theme baked in, so carrying it
	 * forward as a literal silently stops the stop tracking the theme it came from.
	 * `transforms` is what says whether there was anything to track, and an empty
	 * list here means the stop stated none — not that the reader could not see them.
	 */
	resolvedColor: ResolvedColor | null
	/**
	 * The stop's colour resolved against the slide theme **with its colour
	 * transforms applied** — the final rendered hex, i.e. exactly
	 * `resolvedColor?.effectiveHex ?? null`, kept as a flat field because painting a
	 * gradient is what most callers want. `null` when the colour cannot be made
	 * literal (an unmapped token, or a colour model we do not resolve).
	 */
	effectiveHex: string | null
	/** The stop's opacity (0–1) when an `alpha*` transform set one, else `undefined`. Mirrors `resolvedColor.alpha`. */
	alpha?: number
}

/**
 * A gradient fill (`a:gradFill`), as read from a container — the stops plus the
 * geometry that a bare stop list omits. `linear` gradients carry an
 * {@link GradientFill.angleDeg} (the `a:lin/@ang` direction); `radial`/`path`
 * gradients carry a {@link GradientFill.path} shape (`a:path/@path`, e.g.
 * `circle`). The angle is in **OOXML degrees** (clockwise from 3 o'clock), the
 * same convention the write-side `GradientFillProps.angle` expects, so it
 * round-trips directly.
 */
export interface GradientFill {
	/** `linear` (`a:lin`) or `path` (`a:path`, i.e. radial/rectangular). `null` when neither child is present. */
	kind: 'linear' | 'path' | null
	/** Linear direction in OOXML degrees (clockwise from 3 o'clock), or `null` for a path gradient / when unset. */
	angleDeg: number | null
	/** Path-gradient shape (`a:path/@path`: `circle`/`rect`/`shape`), or `null` for a linear gradient. */
	path: string | null
	/** The gradient stops in document order. */
	stops: GradientStop[]
}

/**
 * Read `a:gradFill/a:gsLst` stops from a container (`p:spPr` for a fill, `a:ln`
 * for a line stroke, or `p:bgPr` for a slide background). `null` when the
 * container has no gradient; `[]` when the gradient carries no stop list.
 */
export function readGradientStops(container: Element, ctx: ColorContext): GradientStop[] | null {
	const grad = firstChild(container, 'a:gradFill')
	if (!grad) return null
	const gsLst = firstChild(grad, 'a:gsLst')
	if (!gsLst) return []
	return getElements(gsLst, 'a:gs').map((gs) => {
		// `a:gs` holds exactly one `a:EG_ColorChoice` child, so the stop's colour is
		// whichever element that is. Reading it by *position* rather than by hunting
		// for two known tag names is what lets every model `resolveColorElement`
		// handles reach the resolver: a `a:prstClr`/`a:sysClr`/`a:hslClr` stop used to
		// come back blank in every field, though the reader resolves all three
		// everywhere else.
		const colorEl = firstChildElement(gs)
		const model = colorEl?.localName ?? null
		const pos = pctAttr(gs, 'pos')
		const resolved = resolveColorElement(colorEl, ctx)
		const rawOf = (local: string) => (colorEl && model === local ? attr(colorEl, 'val') : null)
		return {
			position: pos,
			color: rawOf('srgbClr'),
			schemeColor: rawOf('schemeClr'),
			presetColor: rawOf('prstClr'),
			resolvedColor: resolved,
			effectiveHex: resolved ? resolved.effectiveHex : null,
			...(resolved?.alpha !== undefined ? { alpha: resolved.alpha } : {}),
		}
	})
}

/**
 * Read the full `a:gradFill` (stops + linear angle / path shape) from a container
 * (`p:spPr` for a fill, `a:ln` for a line stroke, or `p:bgPr` for a slide
 * background). `null` when the container has no gradient.
 */
export function readGradientFill(container: Element, ctx: ColorContext): GradientFill | null {
	const grad = firstChild(container, 'a:gradFill')
	if (!grad) return null
	const lin = firstChild(grad, 'a:lin')
	const path = firstChild(grad, 'a:path')
	const ang = lin ? numberValue(attr(lin, 'ang')) : null
	return {
		kind: lin ? 'linear' : path ? 'path' : null,
		angleDeg: ang === null ? null : ang / ANGLE_UNITS_PER_DEGREE,
		path: path ? (attr(path, 'path') ?? null) : null,
		stops: readGradientStops(container, ctx) ?? [],
	}
}
