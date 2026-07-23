/**
 * Shared gradient-fill reader for DrawingML colour-bearing containers: a shape's
 * `p:spPr`, a line's `a:ln`, and a slide's `p:bg/p:bgPr`. Decodes an `a:gradFill`
 * into its stops (colour + position) plus the linear/path geometry, resolving each
 * stop's colour against the slide theme.
 */
import { attr, firstChild, getElements, intValue, type Element } from '../oxml/dom.js'
import type { ColorContext } from '../oxml/theme.js'
import { resolveColorElement } from './theme-context.js'

/** One stop of a gradient fill (`a:gsLst/a:gs`), as read from a shape. */
export interface GradientStop {
	/** Stop offset along the gradient, 0–1 (from `@pos`, thousandths of a percent), or `null` if unset. */
	position: number | null
	/** Explicit RGB colour as 6-hex (`a:srgbClr/@val`), or `null` when the stop uses a scheme colour. */
	color: string | null
	/** Theme colour token (`a:schemeClr/@val`, e.g. `accent1`), or `null` when the stop uses an explicit colour. */
	schemeColor: string | null
	/**
	 * The stop's colour resolved against the slide theme **with its colour
	 * transforms applied** — the final rendered hex (the gradient counterpart of
	 * {@link import('./theme-context.js').ResolvedColor.effectiveHex}). `null` when
	 * the colour cannot be made literal (an unmapped token, or a colour model we do
	 * not resolve).
	 */
	effectiveHex: string | null
	/** The stop's opacity (0–1) when an `alpha*` transform set one, else `undefined`. */
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
		const pos = intValue(attr(gs, 'pos'))
		const srgb = firstChild(gs, 'a:srgbClr')
		const scheme = firstChild(gs, 'a:schemeClr')
		const resolved = resolveColorElement(srgb ?? scheme ?? null, ctx)
		return {
			position: pos === null ? null : pos / 100000,
			color: srgb ? attr(srgb, 'val') : null,
			schemeColor: scheme ? attr(scheme, 'val') : null,
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
	const ang = lin ? intValue(attr(lin, 'ang')) : null
	return {
		kind: lin ? 'linear' : path ? 'path' : null,
		angleDeg: ang === null ? null : ang / 60000,
		path: path ? (attr(path, 'path') ?? null) : null,
		stops: readGradientStops(container, ctx) ?? [],
	}
}
