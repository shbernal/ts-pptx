/**
 * Decoding a shape's `a:effectLst` — shadow, glow, reflection, soft edge.
 *
 * Free functions over `(effectLst, ctx)` rather than methods, because that is all they ever
 * needed: the block was five getters and two private helpers on `Shape` whose only reference
 * to the class was `themeContext()`. Lifting it out is what makes `shapes/base.ts` a shape
 * rather than a shape plus an effect decoder, and it puts the effects beside the other
 * per-concern modules `shapes/` already has (`oxml.ts`, `geometry.ts`).
 *
 * Every decoder returns `null` for an absent element rather than an empty object, and leaves an
 * unstated attribute off its result rather than defaulting it — "absent" and "stated as the
 * default" are different facts about the source part, and a replica that has to carry the
 * effect needs to know which it is.
 */

import { attr, firstChild, firstChildElement, numberValue, pctAttr, type Element } from '../../oxml/dom.js'
import { resolveColorElement } from '../theme-context.js'
import type { ColorContext } from '../../oxml/theme.js'
import type { Glow, InnerShadow, OuterShadow, Reflection, SoftEdge } from './types.js'
import { ANGLE_UNITS_PER_DEGREE, EMU_PER_POINT } from '../../../units.js'

/**
 * One scaled attribute of an effect element, or `null` when the element does not state it.
 *
 * `a:effectLst`'s children carry their measures in EMU and 60,000ths of a degree, and every
 * decoder below has to divide. Four of them did it inline and one built a local `put` helper
 * four lines from the last inline copy; naming the operation once is what keeps the divisor and
 * the "absent stays absent" rule from being restated per attribute.
 * @param el - the effect element
 * @param name - the attribute name
 * @param divisor - EMU per point, or angle units per degree
 * @returns the scaled value, or `null`
 */
function scaledAttr(el: Element, name: string, divisor: number): number | null {
	const raw = numberValue(attr(el, name))
	return raw === null ? null : raw / divisor
}

/**
 * Resolve `colorEl` against the theme and stamp `color`/`colorToken`/`alpha` onto a result.
 * @param out - the effect result being built
 * @param colorEl - the effect's colour child, if any
 * @param ctx - the host's colour context
 */
function applyEffectColor(
	out: { color: string | null; colorToken?: string; alpha?: number },
	colorEl: Element | null,
	ctx: ColorContext
): void {
	const resolved = resolveColorElement(colorEl, ctx)
	if (resolved) {
		out.color = resolved.effectiveHex
		if (resolved.alpha !== undefined) out.alpha = resolved.alpha
	}
	// A `schemeClr` with no `val` leaves `colorToken` off entirely, which is the read model's one
	// spelling of "not a theme colour" — the same invariant `compact()` keeps downstream.
	const token = colorEl && colorEl.localName === 'schemeClr' ? attr(colorEl, 'val') : null
	if (token !== null) out.colorToken = token
}

/** A named child of an effect list (`a:effectLst/<qname>`), or `null`. */
function effect(effectLst: Element | null, qname: string): Element | null {
	return effectLst ? firstChild(effectLst, qname) : null
}

/**
 * Decode a shadow element (`a:outerShdw`/`a:innerShdw` share the fields), resolving its colour.
 * @param shdw - the shadow element
 * @param ctx - the host's colour context
 * @returns the decoded shadow
 */
function readShadow(shdw: Element, ctx: ColorContext): OuterShadow {
	const out: OuterShadow = { color: null }
	// `a:EG_ColorChoice` is a required, single-member group, so the colour element is the shadow's
	// only child and taking the first one is both correct and total — the same thing `glow` below
	// does. Naming `a:srgbClr` and `a:schemeClr` explicitly dropped the other four models on the
	// floor: `a:sysClr` resolves everywhere else in the read model, and this library emits
	// `a:prstClr` itself (`gen/slide/notes.ts`). `resolveColor` now answers for five of the six
	// (`a:scrgbClr` is the exception, and reports no colour rather than a guessed one).
	applyEffectColor(out, firstChildElement(shdw), ctx)
	const blur = scaledAttr(shdw, 'blurRad', EMU_PER_POINT)
	const dist = scaledAttr(shdw, 'dist', EMU_PER_POINT)
	const dir = scaledAttr(shdw, 'dir', ANGLE_UNITS_PER_DEGREE)
	if (blur !== null) out.blurPt = blur
	if (dist !== null) out.offsetPt = dist
	if (dir !== null) out.angleDeg = dir
	return out
}

/**
 * The outer drop shadow (`a:effectLst/a:outerShdw`), or `null`.
 * @param effectLst - the shape's effect list, if any
 * @param ctx - the host's colour context
 */
export function readOuterShadow(effectLst: Element | null, ctx: ColorContext): OuterShadow | null {
	const shdw = effect(effectLst, 'a:outerShdw')
	return shdw ? readShadow(shdw, ctx) : null
}

/**
 * The inner shadow (`a:effectLst/a:innerShdw`), or `null`.
 * @param effectLst - the shape's effect list, if any
 * @param ctx - the host's colour context
 */
export function readInnerShadow(effectLst: Element | null, ctx: ColorContext): InnerShadow | null {
	const shdw = effect(effectLst, 'a:innerShdw')
	return shdw ? readShadow(shdw, ctx) : null
}

/**
 * The glow halo (`a:effectLst/a:glow`), or `null`.
 * @param effectLst - the shape's effect list, if any
 * @param ctx - the host's colour context
 */
export function readGlow(effectLst: Element | null, ctx: ColorContext): Glow | null {
	const glow = effect(effectLst, 'a:glow')
	if (!glow) return null
	const out: Glow = { color: null }
	applyEffectColor(out, firstChildElement(glow), ctx)
	const rad = scaledAttr(glow, 'rad', EMU_PER_POINT)
	if (rad !== null) out.radiusPt = rad
	return out
}

/**
 * The reflection (`a:effectLst/a:reflection`), or `null`. Carries no colour: a reflection is a
 * transform of what is already painted.
 * @param effectLst - the shape's effect list, if any
 */
export function readReflection(effectLst: Element | null): Reflection | null {
	const refl = effect(effectLst, 'a:reflection')
	if (!refl) return null
	const out: Reflection = {}
	const put = (target: keyof Reflection, name: string, div: number): void => {
		const v = scaledAttr(refl, name, div)
		if (v !== null) out[target] = v
	}
	const putPct = (target: keyof Reflection, name: string): void => {
		const v = pctAttr(refl, name)
		if (v !== null) out[target] = v
	}
	put('blurPt', 'blurRad', EMU_PER_POINT)
	put('offsetPt', 'dist', EMU_PER_POINT)
	put('angleDeg', 'dir', ANGLE_UNITS_PER_DEGREE)
	put('fadeAngleDeg', 'fadeDir', ANGLE_UNITS_PER_DEGREE)
	putPct('startAlpha', 'stA')
	putPct('startPos', 'stPos')
	putPct('endAlpha', 'endA')
	putPct('endPos', 'endPos')
	return out
}

/**
 * The soft (feathered) edge (`a:effectLst/a:softEdge`), or `null`.
 * @param effectLst - the shape's effect list, if any
 */
export function readSoftEdge(effectLst: Element | null): SoftEdge | null {
	const soft = effect(effectLst, 'a:softEdge')
	if (!soft) return null
	const rad = scaledAttr(soft, 'rad', EMU_PER_POINT)
	return { radiusPt: rad ?? 0 }
}
