/**
 * Value-level conversions shared by every mapper: geometry, colour, and the small
 * discipline that keeps an IR diffable.
 *
 * The precision rules here are not style choices, they are the measured boundary of what
 * the write API can carry, and each is pinned by a regression test:
 *
 * - A `Coord`-typed option accepts `` `${n}emu` ``, so EMU passes through untouched and
 *   the source geometry is reproduced exactly. That covers every position and size.
 * - `colW`, `rowH`, and `margin` are number/`Margin`-typed **inches** and reject a raw-EMU
 *   string, so they are where exactness stops. `defineLayout`'s `width`/`height` are a
 *   fourth, reached only by a printer: there the imprecision does not degrade quietly but
 *   *throws*, since `appendSlides` compares the two decks' EMU sizes for equality after
 *   converting the declared inches back. Six decimal places is the minimum
 *   at which `Math.round(printed × 914400)` provably returns the original EMU, bounding the
 *   drift at 0.4572 EMU — about half a millionth of an inch. Rounding shorter to suppress
 *   `0.5000000001`-style noise trades a cosmetic problem for a real geometry loss.
 */
import type { IrValue } from '../ir.js'
import type { GraphicFrame } from '../../read.js'

/** EMU per inch (ECMA-376 §20.1.2.1). */
const EMU_PER_INCH = 914400

/**
 * Decimal places used for the inch-typed options. See the module header: this is the
 * proven minimum for an EMU-exact round-trip, not a formatting preference.
 */
const INCH_DECIMALS = 6

/** Geometry as a `Coord` the write API takes verbatim, preserving the exact EMU. */
export function emu(value: number): string {
	return `${Math.round(value)}emu`
}

/**
 * Geometry as inches, for the three options that reject a raw-EMU string. Rounded to
 * {@link INCH_DECIMALS}, which is EMU-exact on the way back.
 */
export function inches(emuValue: number): number {
	return Number((emuValue / EMU_PER_INCH).toFixed(INCH_DECIMALS))
}

/** EMU → points, the unit the write API uses for line widths and font sizes. */
export function points(emuValue: number): number {
	return emuValue / 12700
}

/**
 * Points → inches, at the same proven precision as {@link inches}. Needed because the read
 * model reports text-body insets in points while the write API's `margin` takes inches —
 * and warns on a value `>= 1`, on the assumption that it is a points value passed by
 * mistake, so handing it points would both mis-scale the text and produce a warning.
 */
export function pointsToInches(pt: number): number {
	return Number((pt / 72).toFixed(INCH_DECIMALS))
}

/**
 * A colour for a write-API `color` field. The read model splits a colour into a raw
 * `schemeClr` token and a resolved literal hex; this prefers the token, because a token
 * re-resolves against the destination theme and so keeps the deck recolourable, which is
 * the whole reason a theme colour was authored.
 *
 * Only the ten tokens the write path's `clrMap` covers survive as tokens — the other
 * seven `ST_SchemeColorVal` values degrade to a hex literal there anyway, so passing them
 * through would produce a silently different colour. Callers hand those to
 * {@link literalColor} instead after recording a note.
 */
const WRITABLE_SCHEME_TOKENS = new Set([
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'bg1',
	'bg2',
	'tx1',
	'tx2',
])

/** `true` when a `schemeClr` token survives as a token rather than degrading to hex. */
export function isWritableSchemeToken(token: string | null): boolean {
	return token !== null && WRITABLE_SCHEME_TOKENS.has(token)
}

/** Normalize a hex to the bare 6-digit uppercase form the write API expects. */
export function literalColor(hex: string): string {
	return hex.replace(/^#/, '').toUpperCase()
}

/**
 * Build an IR object from possibly-absent parts, dropping every `undefined`.
 *
 * Every mapper funnels through this. An IR must spell "absent" exactly one way — a
 * missing key — or two IRs describing the same deck compare unequal because one wrote
 * `{ bold: undefined }` and the other wrote `{}`. That would surface as phantom
 * round-trip failures, so the invariant is enforced in one place rather than trusted to
 * every call site.
 *
 * Returns `undefined` when nothing survived, so a caller can drop the whole option object
 * rather than emit an empty one.
 */
export function compact(source: Record<string, IrValue | undefined>): Record<string, IrValue> | undefined {
	const out: Record<string, IrValue> = {}
	let any = false
	for (const key of Object.keys(source).sort()) {
		const value = source[key]
		if (value === undefined) continue
		out[key] = value
		any = true
	}
	return any ? out : undefined
}

/** {@link compact}, but always an object — for a required options argument. */
export function compactRequired(source: Record<string, IrValue | undefined>): Record<string, IrValue> {
	return compact(source) ?? {}
}

/** Drop a value that is `null`, so the read model's "unset" becomes an absent IR key. */
export function orUndefined<T>(value: T | null): T | undefined {
	return value === null ? undefined : value
}

/**
 * A 0–1 opacity as the write API's 0–100 transparency; `undefined` when unset.
 *
 * Fully opaque is *not* `0` here — a caller that means "no transparency at all" passes
 * `undefined`, because emitting `transparency: 0` for an `a:alphaModFix amt="100000"`
 * would produce a key the output cannot have: the write path emits no `a:alphaModFix` for
 * a zero transparency, so the re-read deck reports no alpha and the round trip sees a
 * difference where there is no visible one.
 */
export function alphaToTransparency(alpha: number | null | undefined): number | undefined {
	if (alpha === undefined || alpha === null) return undefined
	return Math.round((1 - alpha) * 100)
}

/**
 * A graphic frame's position as write-API options. Identical to a shape's, minus rotation and
 * flip: a `p:graphicFrame` has neither, so `a:xfrm` carries only the offset and extent.
 */
export function positionOfFrame(frame: GraphicFrame): Record<string, IrValue> {
	const box = frame.absoluteFrame
	if (!box) return {}
	return { x: emu(box.left), y: emu(box.top), w: emu(box.width), h: emu(box.height) }
}
