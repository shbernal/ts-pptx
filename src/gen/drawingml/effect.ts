/**
 * ts-pptx: DrawingML effects (glow and shadow)
 *
 * Emit the `<a:glow>` / `<a:outerShdw>` / `<a:innerShdw>` children of an
 * `<a:effectLst>`, and normalize the user-facing `shadow` options into the
 * internal shape every emit site reads.
 * @see http://officeopenxml.com/drwSp-effects.php
 */

import { DEF_FONT_COLOR } from '../../constants-internal.js'
import { warn } from '../../diagnostics.js'
import type { ShadowProps, TextGlowProps } from '../../types/index.js'
import type { ShadowPropsInternal } from '../../types/internal.js'
import { PERCENT_SCALE } from '../../units.js'
import { clampRangedInput, convertAngleUnits, opacityToAlpha, positiveCoordinateEmu } from '../../units-internal.js'
import { splitRgbaHex, stripHash } from '../../hex-color.js'
import { alphaEl, createColorElement } from './color.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'
import { xsdBool } from '../../ooxml/xsd-boolean.js'

/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * { size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
export function createGlowElement(options: TextGlowProps, defaults: TextGlowProps): string {
	const opts = { ...defaults, ...options }
	// `a:glow/@rad` is `ST_PositiveCoordinate`, exactly like the shadow's `blurRad` and `dist`
	// below, so it takes the same converter. A bare multiply put `size: -5` in the package as
	// `rad="-63500"` and `size: NaN` as `rad="NaN"`, while the structurally identical shadow
	// measure two functions down was clamped and reported -- the asymmetry was inside this one
	// function, with the opacity beside it already guarded.
	const size = positiveCoordinateEmu(opts.size, 'glow/size-out-of-range', 'glow.size')
	const color = opts.color || DEF_FONT_COLOR
	const opacity = opacityToAlpha(opts.opacity ?? 0)

	return el('a:glow', { rad: size }, raw(createColorElement(color, alphaEl(opacity))))
}

/**
 * Creates an `a:outerShdw`/`a:innerShdw` element for a text run, shape, image, or chart.
 * Returns the shadow element only (no wrapping `a:effectLst`) so callers can either combine
 * it with other effects (e.g. glow) inside one `a:effectLst`, or wrap a lone shadow via
 * {@link createShadowEffectLst}.
 *
 * Colors go through {@link createColorElement}, so scheme colors (e.g. `accent1`) are honored
 * — earlier per-site copies hardcoded `a:srgbClr` and silently emitted invalid OOXML for them.
 * @param {ShadowPropsInternal} options shadow properties
 * @param {ShadowPropsInternal} defaults defaults for unspecified properties in `options`
 * @see http://officeopenxml.com/drwSp-effects.php
 * @returns {string} XML string, or '' when type is 'none'
 */
export function createShadowElement(options: ShadowPropsInternal | undefined, defaults: ShadowPropsInternal): string {
	const opts = { ...defaults, ...options }
	if (opts.type === 'none') return ''

	// NOTE: read into locals so we never mutate the caller's options (re-emission
	// would otherwise re-convert pt→EMU and produce absurd values).
	const type = opts.type || 'outer'
	const blur = positiveCoordinateEmu(opts.blur ?? 0, 'shadow/blur-out-of-range', 'shadow.blur')
	const offset = positiveCoordinateEmu(opts.offset ?? 0, 'shadow/offset-out-of-range', 'shadow.offset')
	const angle = convertAngleUnits(opts.angle ?? 0, 'shadow.angle')
	const opacity = Math.round((opts._alpha ?? 0.75) * PERCENT_SCALE)
	const color = opts.color || DEF_FONT_COLOR

	// sx/sy/kx/ky/algn/rotWithShape are valid only on `a:outerShdw` (CT_OuterShadowEffect);
	// `a:innerShdw` (CT_InnerShadowEffect) accepts only blurRad/dist/dir.
	const outerAttrs: XmlAttrs =
		type === 'outer'
			? { sx: 100000, sy: 100000, kx: 0, ky: 0, algn: 'bl', rotWithShape: xsdBool(opts.rotateWithShape) }
			: {}
	const attrs = { ...outerAttrs, blurRad: blur, dist: offset, dir: angle }

	return el(`a:${type}Shdw`, attrs, raw(createColorElement(color, alphaEl(opacity))))
}

/**
 * Wraps a lone shadow in the `a:effectLst` that CT_ShapeProperties (shapes/images) and chart
 * marker/data-point properties require. Returns a self-closing `<a:effectLst/>` when there is
 * no shadow (missing/non-object options, or `type: 'none'`), matching the "no effects" element
 * PowerPoint emits. Use this instead of {@link createShadowElement} when the shadow is the only
 * effect; use `createShadowElement` directly when combining with other effects (e.g. glow).
 * @param {ShadowPropsInternal} options shadow properties
 * @param {ShadowPropsInternal} defaults defaults for unspecified properties in `options`
 * @returns {string} `<a:effectLst>…</a:effectLst>` or `<a:effectLst/>`
 */
export function createShadowEffectLst(options: ShadowPropsInternal | undefined, defaults: ShadowPropsInternal): string {
	if (!options || typeof options !== 'object') return voidEl('a:effectLst')
	const inner = createShadowElement(options, defaults)
	return inner ? el('a:effectLst', null, raw(inner)) : voidEl('a:effectLst')
}

/**
 * Normalize a caller's `shadow` into the internal shape every emit site reads: a legal `type`, a
 * whole-degree `angle` inside `ST_PositiveFixedAngle`, a `_alpha` derived from `transparency` or
 * from an RGBA colour's alpha byte, and a colour with no `#` and no alpha byte.
 *
 * **Pure.** It used to normalize the caller's object *in place* and return it, and its four call
 * sites split on which of the two contracts they used — two assigned the result, two discarded it
 * and relied on the mutation. That made the obvious cleanup a trap: making it pure without
 * converting every caller silently drops `transparency` and RGBA alpha on every text run, because
 * `_alpha` is what `genXmlTextRunProperties` reads and nothing else writes it. The value now
 * reaches the emitter as a value rather than as a side effect, and a shadow literal shared across
 * two shapes comes back untouched.
 * @param shadow - the caller's shadow options
 * @returns a fresh normalized bag, or `undefined` when the caller stated no shadow object
 */
export function normalizeShadowOptions(shadow?: ShadowProps | null): ShadowPropsInternal | undefined {
	if (!shadow || typeof shadow !== 'object') {
		// warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`")
		return undefined
	}
	const corrected: ShadowPropsInternal = { ...shadow }
	// No `opacity` scrub is needed: the derived alpha lives under the private `_alpha` name, so a
	// stray `opacity` from an untyped/legacy caller lands on a field nothing reads (inert) rather
	// than colliding with the internal value.

	// OPT: `type`
	if (corrected.type !== 'outer' && corrected.type !== 'inner' && corrected.type !== 'none') {
		warn('shadow/invalid-type', 'shadow.type options are `outer`, `inner` or `none`.')
		corrected.type = 'outer'
	}

	// OPT: `angle` — clamped into ST_PositiveFixedAngle's 0-359 and rounded to whole degrees.
	// The absence check is `undefined`, not truthiness: `0` is a legal angle (a shadow cast to
	// the right), and under `if (corrected.angle)` a `NaN` fell through the guard entirely and
	// reached the emitter, which wrote `dir="NaN"` into the package.
	if (corrected.angle !== undefined) {
		corrected.angle = Math.round(
			clampRangedInput(corrected.angle, 0, 359, 'shadow/angle-out-of-range', 'shadow.angle', 'shadow/angle-non-finite')
		)
	}

	// OPT: `transparency` (PowerPoint UI term, 0-100) -> internal `_alpha` (0.0-1.0), which
	// every emit site reads.
	if (corrected.transparency !== undefined) {
		const pct = clampRangedInput(
			corrected.transparency,
			0,
			100,
			'shadow/transparency-out-of-range',
			'shadow.transparency'
		)
		corrected._alpha = 1 - pct / 100
	}

	// OPT: `color`
	if (corrected.color) {
		// INCORRECT FORMAT
		if (corrected.color.startsWith('#')) {
			warn('shadow/color-has-hash', 'shadow.color should not include hash (#) character, , e.g. "FF0000"')
			corrected.color = stripHash(corrected.color)
		}

		// 8-char hex (RGBA) — derive `_alpha` from the alpha byte (only when `transparency`
		// didn't already set one), then strip the alpha byte from the color so emit sites
		// produce valid 6-char `<a:srgbClr val="…"/>`.
		const rgba = splitRgbaHex(corrected.color)
		if (rgba.alpha !== undefined) {
			if (corrected._alpha === undefined) corrected._alpha = rgba.alpha
			corrected.color = rgba.rgb
		}
	}

	return corrected
}
