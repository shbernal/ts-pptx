/**
 * Shared solid-fill read/mutation helpers for DrawingML colour-bearing parents:
 * a run's `a:rPr`, a shape's `p:spPr` / `p:grpSpPr`, and a line's `a:ln`.
 *
 * A parent carries at most one `EG_FillProperties` choice, so setting a solid
 * fill first clears any competing choice, then inserts `a:solidFill` in document
 * order. These helpers never mark a part dirty — callers own that.
 */
import { InvalidOptionError } from '../../errors.js'
import { isHexColor, stripHash } from '../../hex-color.js'
import { FILL_CHOICES } from '../../ooxml/sequence.js'
import { checkEnumOrThrow } from '../../ooxml/check-enum.js'
import { SCHEME_COLOR_VALUES } from '../../ooxml/st-enums.js'
import {
	attr,
	createElement,
	firstChild,
	getOrAddChild,
	ownerDocumentOf,
	removeChildrenByQName,
	setAttr,
	type Element,
} from './dom.js'

/**
 * The mutually-exclusive fill choices (`EG_FillProperties`); a parent has at most one.
 * Declared in `src/ooxml/sequence.ts`, where the schema sequences that embed it live too.
 */
export { FILL_CHOICES }

/** Normalize a 6-hex RGB string (optional leading `#`) to upper-case, or throw. */
export function normalizeHex(value: string): string {
	const hex = stripHash(value)
	if (!isHexColor(hex))
		throw new InvalidOptionError(
			'color/invalid-hex',
			`Expected a 6-digit hex RGB colour, got: ${JSON.stringify(value)}`
		)
	return hex.toUpperCase()
}

/**
 * A theme colour token for `a:schemeClr/@val`, or a thrown error. The scheme-colour counterpart of
 * {@link normalizeHex}: a read-model setter that writes a token checks it here, against
 * `ST_SchemeColorVal`, so a misspelled one is refused rather than written into the part.
 */
export function schemeToken(value: string): string {
	return checkEnumOrThrow(value, SCHEME_COLOR_VALUES, 'scheme colour token', 'color/invalid-scheme-token')
}

/** The `@val` of a colour child (`qname`) under `parent/a:solidFill`, or `null`. */
export function solidFillColor(parent: Element | null, qname: string): string | null {
	const fill = parent && firstChild(parent, 'a:solidFill')
	const clr = fill && firstChild(fill, qname)
	return clr ? attr(clr, 'val') : null
}

/**
 * The `@val` of an **already-resolved** colour element, when it is exactly the named kind.
 *
 * The sibling of {@link solidFillColor}, for the callers that hold the colour element itself
 * rather than its parent: a run's `a:highlight` child, a bullet's `a:buClr` child, a picture
 * recolour's `a:duotone` entries. Each of those wrote the `localName` test inline, and each
 * paired it with the same `attr(el, 'val')`.
 *
 * `null` for a different colour kind is the point, not a failure: a `schemeClr` has no literal
 * hex to give, and a caller with no theme context has nothing to resolve it against.
 *
 * @param el - a DrawingML colour element (`a:srgbClr`, `a:schemeClr`, …), or nothing
 * @param kind - the colour element this caller can use
 */
export function colorValueIf(el: Element | null | undefined, kind: 'srgbClr' | 'schemeClr' | 'prstClr'): string | null {
	return el && el.localName === kind ? (attr(el, 'val') ?? null) : null
}

/** A solid colour a read-model setter writes: a 6-hex RGB value, or a theme colour token. */
export type SolidFillEdit = { readonly hex: string } | { readonly scheme: string }

/**
 * Whether `container` states a fill of its own: any `EG_FillProperties` choice, solid or not.
 *
 * A container that does decides its own colour. When that choice is not a solid colour
 * (`a:noFill`, a gradient), a resolver reports no colour rather than falling through to one the
 * element would otherwise inherit and does not paint in.
 */
export function hasFillChoice(container: Element | null | undefined): container is Element {
	return !!container && FILL_CHOICES.some((qname) => firstChild(container, qname))
}

/**
 * Clear `container`'s solid fill, or replace its fill with a solid colour, and report whether the
 * container changed.
 *
 * Every read-model paint setter goes through this, so they agree on when a part is dirty. A clear
 * changes something only when there was an `a:solidFill` to remove, and leaves any other fill
 * choice alone. A colour is checked (`normalizeHex`, `schemeToken`) before anything is touched,
 * then replaces whatever fill choice was there.
 * @param container - the element holding the fill, or `null` when it does not exist yet
 * @param edit - the colour to write, or `null` to clear
 * @param getOrAdd - the container, created in document order when absent; called only to write
 * @param after - the successors `a:solidFill` is inserted before
 */
export function applySolidFill(container: Element | null, edit: null): boolean
export function applySolidFill(
	container: Element | null,
	edit: SolidFillEdit,
	getOrAdd: () => Element,
	after: readonly string[]
): boolean
export function applySolidFill(
	container: Element | null,
	edit: SolidFillEdit | null,
	getOrAdd?: () => Element,
	after: readonly string[] = []
): boolean {
	if (edit === null) {
		if (!container || !firstChild(container, 'a:solidFill')) return false
		removeChildrenByQName(container, ['a:solidFill'])
		return true
	}
	const color =
		'hex' in edit
			? { qname: 'a:srgbClr', val: normalizeHex(edit.hex) }
			: { qname: 'a:schemeClr', val: schemeToken(edit.scheme) }
	const target = getOrAdd?.() ?? container
	if (!target) return false
	setSolidFill(target, after, color)
	return true
}

/**
 * Replace `container`'s fill with an explicit `a:noFill`, and report whether the container
 * changed. It does not when the container already holds `a:noFill` and no other fill choice.
 * @param container - the element holding the fill, or `null` when it does not exist yet
 * @param getOrAdd - the container, created in document order when absent
 * @param after - the successors `a:noFill` is inserted before
 */
export function applyNoFill(container: Element | null, getOrAdd: () => Element, after: readonly string[]): boolean {
	const alreadyNoFill =
		!!container &&
		!!firstChild(container, 'a:noFill') &&
		FILL_CHOICES.every((qname) => qname === 'a:noFill' || !firstChild(container, qname))
	if (alreadyNoFill) return false
	const target = getOrAdd()
	removeChildrenByQName(target, FILL_CHOICES)
	getOrAddChild(target, 'a:noFill', after)
	return true
}

/**
 * Replace `parent`'s solid fill with a single colour element. Any competing fill
 * choice is dropped first, then `a:solidFill` is inserted before `afterOrder`
 * (its schema successors). To *clear* a fill instead, remove `a:solidFill`
 * directly via `removeChildrenByQName`, which leaves any other choice untouched.
 */
export function setSolidFill(
	parent: Element,
	afterOrder: readonly string[],
	color: { qname: string; val: string }
): void {
	removeChildrenByQName(parent, FILL_CHOICES)
	const fill = getOrAddChild(parent, 'a:solidFill', afterOrder)
	const clr = createElement(ownerDocumentOf(parent), color.qname)
	setAttr(clr, 'val', color.val)
	fill.appendChild(clr)
}
