/**
 * ts-pptx: DrawingML effects (glow and shadow)
 *
 * Emit the `<a:glow>` / `<a:outerShdw>` / `<a:innerShdw>` children of an
 * `<a:effectLst>`, and normalize the user-facing `shadow` options into the
 * internal shape every emit site reads.
 * @see http://officeopenxml.com/drwSp-effects.php
 */

import { DEF_FONT_COLOR } from '../../core-enums-internal.js'
import { warn } from '../../diagnostics.js'
import type { ShadowProps, TextGlowProps } from '../../core-interfaces.js'
import type { ShadowPropsInternal } from '../../types/internal.js'
import { ANGLE_UNITS_PER_DEGREE, EMU_PER_POINT, PERCENT_SCALE } from '../../units.js'
import { opacityToAlpha, valToPts } from '../../units-internal.js'
import { createColorElement } from './color.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'

/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * { size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
export function createGlowElement(options: TextGlowProps, defaults: TextGlowProps): string {
	const opts = { ...defaults, ...options }
	const size = Math.round(opts.size * EMU_PER_POINT)
	const color = opts.color || DEF_FONT_COLOR
	const opacity = opacityToAlpha(opts.opacity ?? 0)

	return el('a:glow', { rad: size }, raw(createColorElement(color, alpha(opacity))))
}

/** The `<a:alpha>` child both effects hang off their color element. */
function alpha(value: number): string {
	return voidEl('a:alpha', { val: value })
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
	const blur = valToPts(opts.blur ?? 0)
	const offset = valToPts(opts.offset ?? 0)
	const angle = Math.round((opts.angle ?? 0) * ANGLE_UNITS_PER_DEGREE)
	const opacity = Math.round((opts._alpha ?? 0.75) * PERCENT_SCALE)
	const color = opts.color || DEF_FONT_COLOR

	// sx/sy/kx/ky/algn/rotWithShape are valid only on `a:outerShdw` (CT_OuterShadowEffect);
	// `a:innerShdw` (CT_InnerShadowEffect) accepts only blurRad/dist/dir.
	const outerAttrs: XmlAttrs =
		type === 'outer'
			? { sx: 100000, sy: 100000, kx: 0, ky: 0, algn: 'bl', rotWithShape: opts.rotateWithShape ? 1 : 0 }
			: {}
	const attrs = { ...outerAttrs, blurRad: blur, dist: offset, dir: angle }

	return el(`a:${type}Shdw`, attrs, raw(createColorElement(color, alpha(opacity))))
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
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} ShadowProps - shadow options
 */
export function correctShadowOptions(ShadowProps?: ShadowProps | null): ShadowPropsInternal | undefined {
	if (!ShadowProps || typeof ShadowProps !== 'object') {
		// warn("`shadow` options must be an object. Ex: `{shadow: {type:'none'}}`")
		return undefined
	}
	const corrected: ShadowPropsInternal = ShadowProps
	// No `opacity` scrub is needed: the derived alpha lives under the private `_alpha` name, so a
	// stray `opacity` from an untyped/legacy caller lands on a field nothing reads (inert) rather
	// than colliding with the internal value.

	// OPT: `type`
	if (corrected.type !== 'outer' && corrected.type !== 'inner' && corrected.type !== 'none') {
		warn('shadow/invalid-type', 'shadow.type options are `outer`, `inner` or `none`.')
		corrected.type = 'outer'
	}

	// OPT: `angle`
	if (corrected.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(corrected.angle)) || corrected.angle < 0 || corrected.angle > 359) {
			warn('shadow/angle-out-of-range', 'shadow.angle can only be 0-359')
			corrected.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		corrected.angle = Math.round(Number(corrected.angle))
	}

	// OPT: `transparency` (PowerPoint UI term, 0-100) -> internal `_alpha` (0.0-1.0), which
	// every emit site reads.
	if (corrected.transparency !== undefined) {
		const pct = Number(corrected.transparency)
		if (isNaN(pct) || pct < 0 || pct > 100) {
			warn('shadow/transparency-out-of-range', 'shadow.transparency can only be 0-100')
		} else {
			corrected._alpha = 1 - pct / 100
		}
	}

	// OPT: `color`
	if (corrected.color) {
		// INCORRECT FORMAT
		if (corrected.color.startsWith('#')) {
			warn('shadow/color-has-hash', 'shadow.color should not include hash (#) character, , e.g. "FF0000"')
			corrected.color = corrected.color.replace('#', '')
		}

		// 8-char hex (RGBA) — derive `_alpha` from the alpha byte (only when `transparency`
		// didn't already set one), then strip the alpha byte from the color so emit sites
		// produce valid 6-char `<a:srgbClr val="…"/>`.
		if (/^[0-9a-fA-F]{8}$/.test(corrected.color)) {
			const alphaHex = corrected.color.slice(6, 8)
			if (corrected._alpha === undefined) {
				corrected._alpha = parseInt(alphaHex, 16) / 255
			}
			corrected.color = corrected.color.slice(0, 6)
		}
	}

	return corrected
}
