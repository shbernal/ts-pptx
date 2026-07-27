/**
 * A {@link DeckIr} reduced to the form a round-trip diff can compare.
 *
 * The round trip is `source → IR₁ → script → deck → IR₂`. Comparing those two IRs raw
 * produces thousands of differences, almost none of which are losses: the write path
 * spells out defaults that OOXML leaves implicit, and PowerPoint spells out defaults the
 * write path leaves implicit. Both directions are noise, and drowning the real losses in
 * it is the same as not checking at all.
 *
 * **Every rule here must be an equivalence, not a convenience.** A rule that merely makes
 * the report shorter hides a defect permanently and silently — the exact failure mode this
 * subsystem exists to catch. So each one below cites the OOXML default that makes the
 * explicit and the absent spelling mean the same thing, and none of them drops a value that
 * a viewer could render differently. When in doubt the value stays in and the difference is
 * reported: a noisy report is recoverable, a quiet one is not.
 *
 * Canonicalising is deliberately *symmetric* — it runs over both IRs, so it cannot favour
 * the output. A rule that only made sense in one direction would be an exclusion dressed up
 * as a normalisation.
 */
import type { AssetIr, DeckIr, IrValue, SlideIr } from '../ir.js'
import { isAssetRef } from '../ir.js'

/** One write-API call, with the source shape name lifted out as an address rather than data. */
export interface CanonicalCall {
	method: string
	/** `p:cNvPr/@name` of the shape this call came from, used to point a difference at it. */
	shapeName: string | null
	/**
	 * Every `objectName` appearing anywhere inside {@link args}, which for `addGroup` is the
	 * group's children.
	 *
	 * A fidelity note about a group's child is scoped to that child's name, but the child is
	 * not a call of its own — it is an argument to the group's. Without this the note and the
	 * difference it predicted would carry different names and never match, and a declared loss
	 * would be reported as a defect.
	 */
	containedNames: string[]
	args: IrValue[]
}

export interface CanonicalSlide {
	number: number
	hidden: boolean
	background: IrValue | null
	notesText: string | null
	calls: CanonicalCall[]
}

export interface CanonicalDeck {
	slideSize: { widthEmu: number; heightEmu: number }
	slides: CanonicalSlide[]
}

/**
 * Option values that mean exactly what their absence means.
 *
 * Keyed by option name because the IR's argument objects are the write API's own option
 * objects, where a name identifies a construct uniquely — there is no `bold` that means
 * something other than `a:rPr/@b`.
 *
 * Each entry is an OOXML default:
 *
 * - `bold`, `italic` — `a:rPr/@b` / `@i` are `ST_OnOff`, default `false`.
 * - `baseline` — `a:rPr/@baseline` is a percentage, default `0` (no super/subscript).
 * - `breakLine` — not OOXML but the IR's own flat encoding: a run that does not end its
 *   paragraph. The final run of a text frame never carries one, so absent and `false` are
 *   the same statement.
 * - `wrap` — `a:bodyPr/@wrap` defaults to `square`, which the write API spells `true`.
 * - `rotate` — `a:xfrm/@rot` defaults to `0`.
 * - `flipH`, `flipV` — `a:xfrm/@flipH` / `@flipV` default to `false`.
 * - `transparency` — an alpha of 0% removed, i.e. fully opaque, which is also the default
 *   when no `a:alpha` transform is present.
 * - `indentLevel` — `a:p/@lvl` defaults to `0`.
 * - `dashType` — `a:ln`'s dash defaults to `solid` when no `a:prstDash` is present.
 *
 * Line width is deliberately **absent**: `a:ln/@w` defaults to 0 (hairline) while the write
 * path defaults it to 1pt, so those two are not the same line and the difference is real.
 */
const IMPLIED_DEFAULTS: Record<string, IrValue> = {
	bold: false,
	italic: false,
	baseline: 0,
	breakLine: false,
	wrap: true,
	rotate: 0,
	flipH: false,
	flipV: false,
	transparency: 0,
	indentLevel: 0,
	dashType: 'solid',
}

/**
 * `a:bodyPr`'s default insets in inches, in the write API's `[top, right, bottom, left]`
 * order: 45720 EMU (0.05in, 3.6pt) top and bottom, 91440 EMU (0.1in, 7.2pt) left and right.
 * A body that spells them out is inset exactly as one that omits them.
 */
const DEFAULT_MARGIN: readonly number[] = [0.05, 0.1, 0.05, 0.1]

/** Reduce a deck IR to the comparable form. */
export function canonicalDeckIr(ir: DeckIr): CanonicalDeck {
	const digests = assetDigests(ir.assets)
	return {
		slideSize: ir.slideSize,
		slides: ir.slides.map((slide) => canonicalSlide(slide, digests)),
	}
}

function canonicalSlide(slide: SlideIr, digests: Map<string, string>): CanonicalSlide {
	return {
		number: slide.number,
		hidden: slide.hidden,
		background: slide.background === undefined ? null : canonicalValue(slide.background as IrValue, digests),
		notesText: slide.notesText ?? null,
		calls: slide.calls.map((call) => {
			const args = call.args.map((arg) => canonicalValue(arg, digests))
			return {
				method: call.method,
				shapeName: call.sourceName ?? null,
				containedNames: [...collectObjectNames(args, new Set())],
				args,
			}
		}),
	}
}

/** Every `objectName` value reachable in `value`, at any depth. */
function collectObjectNames(value: IrValue, out: Set<string>): Set<string> {
	if (Array.isArray(value)) {
		for (const item of value) collectObjectNames(item, out)
		return out
	}
	if (value === null || typeof value !== 'object') return out
	for (const [key, item] of Object.entries(value)) {
		if (key === 'objectName' && typeof item === 'string') out.add(item)
		else collectObjectNames(item as IrValue, out)
	}
	return out
}

/**
 * Recursively drop implied defaults and resolve asset references to a content digest.
 *
 * The digest matters more than it looks: assets are named `image1.png`, `image2.png`, … in
 * the order the walk meets them, and the output deck can meet the same bytes in a different
 * order or under a different extension. Comparing names would report a swap that did not
 * happen and, worse, would *miss* one where two images traded places.
 */
function canonicalValue(value: IrValue, digests: Map<string, string>): IrValue {
	if (isAssetRef(value)) return { $asset: digests.get(value.$asset) ?? `unknown:${value.$asset}` }
	if (Array.isArray(value)) return value.map((item) => canonicalValue(item, digests))
	if (value === null || typeof value !== 'object') return value

	// A run with no characters has no colour anyone can see, and the write path gives every
	// run it emits an explicit one. Narrowed to `color` deliberately: a size or a bullet on an
	// empty paragraph still affects line height and still draws a glyph, so those stay.
	const blankRun = value['text'] === '' && typeof value['options'] === 'object' && value['options'] !== null

	const out: Record<string, IrValue> = {}
	for (const [key, item] of Object.entries(value)) {
		if (key in IMPLIED_DEFAULTS && item === IMPLIED_DEFAULTS[key]) continue
		if (key === 'margin' && isDefaultMargin(item)) continue
		if (blankRun && key === 'options') {
			const visible = { ...(item as Record<string, IrValue>) }
			delete visible['color']
			out[key] = canonicalValue(visible, digests)
			continue
		}
		out[key] = canonicalValue(item, digests)
	}
	return out
}

function isDefaultMargin(value: IrValue): boolean {
	return (
		Array.isArray(value) &&
		value.length === DEFAULT_MARGIN.length &&
		value.every((item, index) => item === DEFAULT_MARGIN[index])
	)
}

/**
 * Content digests keyed by asset name, so the same bytes compare equal under any name.
 *
 * FNV-1a rather than a real hash: this identifies bytes for a comparison both sides of
 * which are in hand, so collision resistance against an adversary is not a property it
 * needs, and the library stays free of `node:crypto` and of the isomorphism problem
 * importing it would create.
 */
function assetDigests(assets: AssetIr[]): Map<string, string> {
	const out = new Map<string, string>()
	for (const asset of assets) {
		let hash = 0x811c9dc5
		for (const byte of asset.bytes) {
			hash ^= byte
			hash = Math.imul(hash, 0x01000193) >>> 0
		}
		out.set(asset.name, `${asset.bytes.length}:${hash.toString(16)}`)
	}
	return out
}
