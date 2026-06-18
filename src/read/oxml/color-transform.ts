/**
 * Apply DrawingML colour-transform modifiers (`lumMod`/`lumOff`/`shade`/`tint`/
 * `satMod`/`alpha`/…) to a base sRGB colour, returning the *effective* colour a
 * renderer would paint.
 *
 * The read model resolves a `schemeClr`/`srgbClr` reference to its **base** hex
 * and carries the transform children verbatim (so the `theme: 'preserve'` flatten
 * path can re-emit them byte-for-byte). This module is the missing step that
 * turns that base hex + transform list into the final rendered hex, so a consumer
 * never hand-computes a tint/shade again.
 *
 * Colour-space model (verified against PowerPoint output — see
 * `test/read/color-transform.test.js`): the working colour is kept canonical as
 * sRGB between transforms; each transform converts into the space its family
 * operates in, applies, and converts back to sRGB. Transforms apply in **document
 * order** (ECMA-376 §20.1.2.3).
 *
 *  - `lumMod` / `lumOff` — luminance (L) in HSL: `L' = L·mod + off`. PowerPoint's
 *    "Lighter/Darker N%" family.
 *  - `satMod` / `satOff` — saturation (S) in HSL.
 *  - `hueMod` / `hueOff` — hue (H) in HSL (`hueMod` scales, `hueOff` offsets by
 *    degrees). Not exercised by any current source; implemented for completeness.
 *  - `shade` — toward black in **linear** RGB: `c' = c·k`.
 *  - `tint` — toward white in **linear** RGB: `c' = c·k + (1−k)`.
 *  - `alpha` / `alphaMod` / `alphaOff` — opacity, returned alongside the hex
 *    (never folded into RGB).
 *  - `comp` / `inv` / `gray` / `gamma` / `invGamma` — pass-through for now (no
 *    current source needs them; add when one does).
 *
 * All DrawingML percentage values are stored in thousandths of a percent
 * (`100%` → `100000`), so a modifier's fraction is `val / 100000`.
 */

/** A colour-transform modifier in its read-model form: tag local-name + raw `@val`. */
export interface ColorTransform {
	name: string
	value: string | null
}

/** The effective colour after applying transforms: final sRGB hex, plus opacity when an `alpha*` modifier was present. */
export interface EffectiveColor {
	hex: string
	alpha?: number
}

/** DrawingML percentage (`100%` = `100000`) → fraction, or `null` when unparseable. */
function pct(value: string | null): number | null {
	if (value === null || value === '') return null
	const n = Number(value)
	return Number.isFinite(n) ? n / 100000 : null
}

const clamp01 = (n: number): number => (n < 0 ? 0 : n > 1 ? 1 : n)

/** Parse a 6-hex string (`RRGGBB`, optional `#`) into 0–1 sRGB channels, or `null`. */
function parseHex(hex: string): { r: number; g: number; b: number } | null {
	const h = hex.startsWith('#') ? hex.slice(1) : hex
	if (!/^[0-9a-fA-F]{6}$/.test(h)) return null
	return {
		r: parseInt(h.slice(0, 2), 16) / 255,
		g: parseInt(h.slice(2, 4), 16) / 255,
		b: parseInt(h.slice(4, 6), 16) / 255,
	}
}

/** 0–1 sRGB channels → uppercase `RRGGBB`, rounding and clamping each channel. */
function toHex(c: { r: number; g: number; b: number }): string {
	const ch = (v: number): string =>
		Math.round(clamp01(v) * 255)
			.toString(16)
			.toUpperCase()
			.padStart(2, '0')
	return `${ch(c.r)}${ch(c.g)}${ch(c.b)}`
}

/** sRGB channel (0–1) → linear-light (gamma decode). */
function srgbToLinear(c: number): number {
	return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4)
}

/** Linear-light channel (0–1) → sRGB (gamma encode). */
function linearToSrgb(c: number): number {
	return c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055
}

/** sRGB (0–1) → HSL with H in degrees [0,360), S/L in [0,1]. */
function rgbToHsl(c: { r: number; g: number; b: number }): { h: number; s: number; l: number } {
	const { r, g, b } = c
	const max = Math.max(r, g, b)
	const min = Math.min(r, g, b)
	const l = (max + min) / 2
	const d = max - min
	if (d === 0) return { h: 0, s: 0, l }
	const s = d / (1 - Math.abs(2 * l - 1))
	let h: number
	if (max === r) h = ((g - b) / d) % 6
	else if (max === g) h = (b - r) / d + 2
	else h = (r - g) / d + 4
	h *= 60
	if (h < 0) h += 360
	return { h, s, l }
}

/** HSL (H degrees, S/L 0–1) → sRGB (0–1). */
function hslToRgb({ h, s, l }: { h: number; s: number; l: number }): { r: number; g: number; b: number } {
	const c = (1 - Math.abs(2 * l - 1)) * s
	const hp = (((h % 360) + 360) % 360) / 60
	const x = c * (1 - Math.abs((hp % 2) - 1))
	const [r, g, b] =
		hp < 1 ? [c, x, 0] : hp < 2 ? [x, c, 0] : hp < 3 ? [0, c, x] : hp < 4 ? [0, x, c] : hp < 5 ? [x, 0, c] : [c, 0, x]
	const m = l - c / 2
	return { r: r + m, g: g + m, b: b + m }
}

/**
 * Apply an ordered list of DrawingML colour transforms to a base sRGB hex,
 * returning the effective sRGB hex (and opacity, when an `alpha*` modifier set
 * it). Pure: no DOM, no theme lookup — the caller has already resolved the base
 * colour and extracted each modifier's `{ name, value }`.
 *
 * An unparseable base hex is returned unchanged (with no transforms applied); an
 * unparseable/`null` modifier value is skipped.
 */
export function applyColorTransforms(baseHex: string, transforms: ColorTransform[]): EffectiveColor {
	const parsed = parseHex(baseHex)
	if (!parsed) return { hex: baseHex }
	let rgb = parsed
	let alpha: number | undefined

	for (const { name, value } of transforms) {
		switch (name) {
			case 'lumMod':
			case 'lumOff': {
				const f = pct(value)
				if (f === null) break
				const hsl = rgbToHsl(rgb)
				hsl.l = clamp01(name === 'lumMod' ? hsl.l * f : hsl.l + f)
				rgb = hslToRgb(hsl)
				break
			}
			case 'satMod':
			case 'satOff': {
				const f = pct(value)
				if (f === null) break
				const hsl = rgbToHsl(rgb)
				hsl.s = clamp01(name === 'satMod' ? hsl.s * f : hsl.s + f)
				rgb = hslToRgb(hsl)
				break
			}
			case 'hueMod':
			case 'hueOff': {
				const hsl = rgbToHsl(rgb)
				// hueMod is a percentage scale; hueOff is an angle in 60000ths of a degree.
				if (name === 'hueMod') {
					const f = pct(value)
					if (f === null) break
					hsl.h = (((hsl.h * f) % 360) + 360) % 360
				} else {
					const n = value === null ? NaN : Number(value)
					if (!Number.isFinite(n)) break
					hsl.h = (((hsl.h + n / 60000) % 360) + 360) % 360
				}
				rgb = hslToRgb(hsl)
				break
			}
			case 'shade':
			case 'tint': {
				const k = pct(value)
				if (k === null) break
				const lin = { r: srgbToLinear(rgb.r), g: srgbToLinear(rgb.g), b: srgbToLinear(rgb.b) }
				const apply = (c: number): number => clamp01(name === 'shade' ? c * k : c * k + (1 - k))
				rgb = { r: linearToSrgb(apply(lin.r)), g: linearToSrgb(apply(lin.g)), b: linearToSrgb(apply(lin.b)) }
				break
			}
			case 'alpha': {
				const f = pct(value)
				if (f !== null) alpha = clamp01(f)
				break
			}
			case 'alphaMod': {
				const f = pct(value)
				if (f !== null) alpha = clamp01((alpha ?? 1) * f)
				break
			}
			case 'alphaOff': {
				const f = pct(value)
				if (f !== null) alpha = clamp01((alpha ?? 1) + f)
				break
			}
			// comp/inv/gray/gamma/invGamma: pass-through until a source needs them.
			default:
				break
		}
	}

	return alpha === undefined ? { hex: toHex(rgb) } : { hex: toHex(rgb), alpha }
}
