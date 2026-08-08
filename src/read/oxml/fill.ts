/**
 * Shared solid-fill read/mutation helpers for DrawingML colour-bearing parents:
 * a run's `a:rPr`, a shape's `p:spPr` / `p:grpSpPr`, and a line's `a:ln`.
 *
 * A parent carries at most one `EG_FillProperties` choice, so setting a solid
 * fill first clears any competing choice, then inserts `a:solidFill` in document
 * order. These helpers never mark a part dirty — callers own that.
 */
import { InvalidOptionError } from '../../errors.js'
import { FILL_CHOICES } from '../../ooxml/sequence.js'
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
	const hex = value.startsWith('#') ? value.slice(1) : value
	if (!/^[0-9a-fA-F]{6}$/.test(hex))
		throw new InvalidOptionError(
			'color/invalid-hex',
			`Expected a 6-digit hex RGB colour, got: ${JSON.stringify(value)}`
		)
	return hex.toUpperCase()
}

/** The `@val` of a colour child (`qname`) under `parent/a:solidFill`, or `null`. */
export function solidFillColor(parent: Element | null, qname: string): string | null {
	const fill = parent && firstChild(parent, 'a:solidFill')
	const clr = fill && firstChild(fill, qname)
	return clr ? attr(clr, 'val') : null
}

/**
 * Replace `parent`'s solid fill with a single colour element. Any competing fill
 * choice is dropped first, then `a:solidFill` is inserted before `afterOrder`
 * (its schema successors). To *clear* a fill instead, remove `a:solidFill`
 * directly via `removeChildrenByQName`, which leaves any other choice untouched.
 */
export function setSolidFill(parent: Element, afterOrder: string[], color: { qname: string; val: string }): void {
	removeChildrenByQName(parent, FILL_CHOICES)
	const fill = getOrAddChild(parent, 'a:solidFill', afterOrder)
	const clr = createElement(ownerDocumentOf(parent), color.qname)
	setAttr(clr, 'val', color.val)
	fill.appendChild(clr)
}
