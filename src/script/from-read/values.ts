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
import type { Shape } from '../../read.js'
import type { NoteScope, RecordableConstruct } from '../fidelity.js'
import { EMU_PER_POINT } from '../../units.js'
// Re-exported so this module stays the one import the mappers reach for; it lives in
// `script/units.ts` because the printer needs it too and may not import from here.
export { inches, INCH_DECIMALS } from '../units.js'
import { INCH_DECIMALS } from '../units.js'
import { stripHash } from '../../hex-color.js'
import { PRESET_LINE_DASHES } from '../../ooxml/st-enums.js'

/** Geometry as a `Coord` the write API takes verbatim, preserving the exact EMU. */
export function emu(value: number): string {
	return `${Math.round(value)}emu`
}

/** EMU → points, the unit the write API uses for line widths and font sizes. */
export function points(emuValue: number): number {
	return emuValue / EMU_PER_POINT
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

/**
 * A source's `schemeClr` token as the write API's `color`, baked to a literal hex (with a
 * note) when the token is one of the seven the write path cannot carry, or `undefined` when
 * the source names no token at all.
 *
 * `undefined` means "this leg has nothing to say" — the caller's own-colour ladder decides
 * from there. Seven sites make this decision and four of them made it correctly: the table
 * fill, the cell fill and the cell border all returned an unwritable token *raw*, so a cell
 * filled `<a:schemeClr val="dk1"/>` emitted `fill: { color: 'dk1' }` and the generated script
 * warned and painted the cell the default text colour, with no note; the gradient stop baked
 * without recording one.
 *
 * @param scheme - the raw `a:schemeClr/@val`, or `null` when the source names none
 * @param resolvedHex - the literal the token resolves to, for the baking leg
 * @param notes - the scope the approximation is recorded on
 * @param construct - the note's construct id, e.g. `table.cell.fill.schemeToken`
 * @param label - how the note's prose names the thing, e.g. `cell fill`
 */
export function schemeColorOption(
	scheme: string | null,
	resolvedHex: string | null,
	notes: NoteScope,
	construct: RecordableConstruct,
	label: string
): string | undefined {
	if (isWritableSchemeToken(scheme)) return scheme as string
	if (scheme === null) return undefined
	notes.note(
		construct,
		'approximated',
		'unwritable',
		`${label} scheme colour "${scheme}" is outside the ten tokens the write path maps, so it is baked to a literal hex and stops tracking the theme`
	)
	return resolvedHex === null ? undefined : literalColor(resolvedHex)
}

/**
 * `a:bodyPr/@anchor` and `a:tcPr/@anchor` -> the write API's `valign`.
 *
 * The inverse of the emitters' own `resolveTextAnchor`, and one table rather than the two
 * byte-identical copies the text and table mappers each carried.
 */
export const ANCHOR_TO_VALIGN: Readonly<Record<string, string>> = { t: 'top', ctr: 'middle', b: 'bottom' }

/**
 * The `a:prstDash` values the write API's `dashType` can spell.
 *
 * A dash outside `ST_PresetLineDashVal` cannot have come from a conformant deck; both mappers
 * that read one drop it to a plain dashed rule and note the loss, and both built this set.
 */
export const WRITABLE_DASHES: ReadonlySet<string> = new Set<string>(PRESET_LINE_DASHES)

/**
 * A shape's or frame's Selection Pane name as the IR's `sourceName`, or nothing when it has none.
 *
 * `sourceName` is how a call in the emitted script says which shape it came from, so it is
 * omitted rather than written empty. Three mappers spelled the ternary; this is it.
 */
export function nameOf(subject: { name: string | null }): { sourceName?: string } {
	return subject.name ? { sourceName: subject.name } : {}
}

/** Normalize a hex to the bare 6-digit uppercase form the write API expects. */
export function literalColor(hex: string): string {
	return stripHash(hex).toUpperCase()
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
 * A 0-1 opacity as the write API's 0-100 transparency; `undefined` when unset **or fully
 * opaque**.
 *
 * Fully opaque is not `0`, and that is this function's rule rather than each caller's:
 * emitting `transparency: 0` for an `a:alphaModFix amt="100000"` produces a key the output
 * cannot have, because the write path emits no `a:alphaModFix` for a zero transparency, so
 * the re-read deck reports no alpha and the round trip sees a difference where there is no
 * visible one. One of five callers implemented that; the other four passed the value straight
 * through, and it stayed invisible only because the canonicaliser drops `transparency: 0` as
 * an implied default.
 */
export function alphaToTransparency(alpha: number | null | undefined): number | undefined {
	if (alpha === undefined || alpha === null) return undefined
	return alpha >= 1 ? undefined : Math.round((1 - alpha) * 100)
}

/**
 * Position as `Coord`-typed EMU strings, so the source geometry survives exactly.
 *
 * `absoluteFrame` rather than the raw `left`/`top`: a shape inside a group is positioned in
 * its group's child coordinate space, which a flattened call list cannot express, so the
 * group transform has to be composed in here. For a top-level shape the two are identical.
 *
 * Takes the `Shape` base rather than `AnyShape` so a `p:graphicFrame` — a table or a chart —
 * gets the same fallback and the same note. It had its own `positionOfFrame`, which returned
 * `{}` in exactly the case this fallback exists for, with no note: `absoluteFrame` is `null`
 * whenever an enclosing group lacks a usable transform, so a table inside such a group emitted
 * `addTable(rows, { objectName, tableStyle })` with no geometry at all.
 */
export function positionOptions(shape: Shape, notes?: NoteScope): Record<string, IrValue> {
	const frame = shape.absoluteFrame
	if (frame) return { x: emu(frame.left), y: emu(frame.top), w: emu(frame.width), h: emu(frame.height) }

	// No usable absolute frame — a placeholder taking its geometry from the layout or master, or
	// a shape inside a group whose own transform is missing or degenerate. `resolvedFrame` walks
	// that chain, and using it is not optional: omitting the geometry does *not* leave it to be
	// inherited, because a regenerated slide's placeholder inherits nothing. It produces
	// `x=0 y=0 w=<slide width> h=0` — a zero-height box in the corner, which is broken output
	// rather than lossy output.
	const resolved = shape.resolvedFrame
	if (!resolved) return {}
	// `source: 'own'` means the shape DOES have its own `a:xfrm` and it is the enclosing group
	// that could not be composed, so the prose says that rather than "takes its geometry from
	// the own".
	notes?.note(
		'shape.frameInherited',
		'flattened',
		'unsupported',
		resolved.source === 'own'
			? 'this shape sits in a group whose transform could not be composed, so its own position is baked in as a slide-absolute one'
			: `this shape has no transform of its own and takes its geometry from the ${resolved.source}; a regenerated slide inherits nothing, so the resolved position is baked in and stops tracking later edits to that ${resolved.source}`
	)
	return { x: emu(resolved.left), y: emu(resolved.top), w: emu(resolved.width), h: emu(resolved.height) }
}
