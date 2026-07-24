/**
 * ts-pptx: DrawingML effects (glow and shadow)
 *
 * Emit the `<a:glow>` / `<a:outerShdw>` / `<a:innerShdw>` children of an
 * `<a:effectLst>`, and normalize the user-facing `shadow` options into the
 * internal shape every emit site reads.
 * @see http://officeopenxml.com/drwSp-effects.php
 */

import { DEF_FONT_COLOR } from '../../core-enums-internal.js'
import { warn } from '../../log.js'
import type { ShadowProps, TextGlowProps } from '../../core-interfaces.js'
import type { ShadowPropsInternal } from '../../types/internal.js'
import { ANGLE_UNITS_PER_DEGREE, EMU_PER_POINT, PERCENT_SCALE } from '../../units.js'
import { opacityToAlpha, valToPts } from '../../units-internal.js'
import { createColorElement } from './color.js'

/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * { size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
export function createGlowElement(options: TextGlowProps, defaults: TextGlowProps): string {
	let strXml = ''
	const opts = { ...defaults, ...options }
	const size = Math.round(opts.size * EMU_PER_POINT)
	const color = opts.color || DEF_FONT_COLOR
	const opacity = opacityToAlpha(opts.opacity ?? 0)

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += '</a:glow>'

	return strXml
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
	const opacity = Math.round((opts.opacity ?? 0.75) * PERCENT_SCALE)
	const color = opts.color || DEF_FONT_COLOR

	// sx/sy/kx/ky/algn/rotWithShape are valid only on `a:outerShdw` (CT_OuterShadowEffect);
	// `a:innerShdw` (CT_InnerShadowEffect) accepts only blurRad/dist/dir.
	const extraAttrs =
		type === 'outer'
			? `sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="${opts.rotateWithShape ? 1 : 0}" `
			: ''
	let strXml = `<a:${type}Shdw ${extraAttrs}blurRad="${blur}" dist="${offset}" dir="${angle}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += `</a:${type}Shdw>`

	return strXml
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
	if (!options || typeof options !== 'object') return '<a:effectLst/>'
	const inner = createShadowElement(options, defaults)
	return inner ? `<a:effectLst>${inner}</a:effectLst>` : '<a:effectLst/>'
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
	// `opacity` is no longer a public input (removed in favor of `transparency`); strip any
	// leftover value from an untyped caller so it doesn't silently keep working through this
	// function's internal reuse of the same field name for the derived alpha.
	delete corrected.opacity

	// OPT: `type`
	if (corrected.type !== 'outer' && corrected.type !== 'inner' && corrected.type !== 'none') {
		warn('shadow.type options are `outer`, `inner` or `none`.')
		corrected.type = 'outer'
	}

	// OPT: `angle`
	if (corrected.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(corrected.angle)) || corrected.angle < 0 || corrected.angle > 359) {
			warn('shadow.angle can only be 0-359')
			corrected.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		corrected.angle = Math.round(Number(corrected.angle))
	}

	// OPT: `transparency` (PowerPoint UI term, 0-100) -> internal `opacity` (0.0-1.0), which
	// every emit site reads.
	if (corrected.transparency !== undefined) {
		const pct = Number(corrected.transparency)
		if (isNaN(pct) || pct < 0 || pct > 100) {
			warn('shadow.transparency can only be 0-100')
		} else {
			corrected.opacity = 1 - pct / 100
		}
	}

	// OPT: `color`
	if (corrected.color) {
		// INCORRECT FORMAT
		if (corrected.color.startsWith('#')) {
			warn('shadow.color should not include hash (#) character, , e.g. "FF0000"')
			corrected.color = corrected.color.replace('#', '')
		}

		// 8-char hex (RGBA) — derive `opacity` from the alpha byte (only when `transparency`
		// didn't already set one), then strip the alpha byte from the color so emit sites
		// produce valid 6-char `<a:srgbClr val="…"/>`.
		if (/^[0-9a-fA-F]{8}$/.test(corrected.color)) {
			const alphaHex = corrected.color.slice(6, 8)
			if (corrected.opacity === undefined) {
				corrected.opacity = parseInt(alphaHex, 16) / 255
			}
			corrected.color = corrected.color.slice(0, 6)
		}
	}

	return corrected
}
