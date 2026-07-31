/**
 * Ordered mutation of a table cell's `a:tcPr`, and the schema facts every such edit needs.
 *
 * `CT_TableCellProperties` declares its children as a **sequence**, so an append-only
 * implementation produces an out-of-order `a:tcPr` — which PowerPoint reports as a corrupt
 * file rather than as a bad edit. Every insertion therefore goes through
 * {@link getOrAddChild} / {@link insertInOrder} with the new element's *successors*, which
 * is what {@link TCPR_AFTER} tabulates. Writing that table once, here, is the point of this
 * module: a setter that hand-rolls its own successor list is one rename away from silently
 * producing an invalid part.
 *
 * The attribute values are the other half. A value outside its `ST_` union is equally fatal
 * and equally silent, so each is checked against the enum before it is written — the same
 * validate-at-the-boundary rule the write path follows, with the difference that here there
 * is a caller to throw at rather than a deck to degrade.
 */
import { InvalidOptionError } from '../../errors.js'
import type { InvalidOptionErrorCode } from '../../codes.js'
import { getOrAddChild, insertInOrder, type Element } from '../oxml/dom.js'

/**
 * `CT_TableCellProperties`' child sequence, in declaration order.
 * @see ECMA-376 Part 1 §21.1.3.17 (`a:tcPr`)
 */
export const TCPR_SEQUENCE: readonly string[] = [
	'a:lnL',
	'a:lnR',
	'a:lnT',
	'a:lnB',
	'a:lnTlToBr',
	'a:lnBlToTr',
	'a:cell3D',
	// EG_FillProperties — one of these at most.
	'a:noFill',
	'a:solidFill',
	'a:gradFill',
	'a:blipFill',
	'a:pattFill',
	'a:grpFill',
	'a:headers',
	'a:extLst',
]

/**
 * For each `a:tcPr` child, the children that must follow it — i.e. what a new one is
 * inserted *before*. Derived from {@link TCPR_SEQUENCE} rather than written out, so the two
 * cannot disagree and adding a child to the sequence updates every successor list at once.
 */
export const TCPR_AFTER: Record<string, string[]> = Object.fromEntries(
	TCPR_SEQUENCE.map((qname, index) => [qname, TCPR_SEQUENCE.slice(index + 1)])
)

/** The four edges plus the two diagonals, keyed as {@link TableCellEdge} names them. */
export const EDGE_QNAMES = {
	left: 'a:lnL',
	right: 'a:lnR',
	top: 'a:lnT',
	bottom: 'a:lnB',
	tlToBr: 'a:lnTlToBr',
	blToTr: 'a:lnBlToTr',
} as const

/** Which border of a cell an edit addresses. */
export type TableCellEdge = keyof typeof EDGE_QNAMES

/** `ST_TextAnchoringType` — `a:tcPr/@anchor`. */
export const ANCHOR_VALUES: readonly string[] = ['t', 'ctr', 'b', 'just', 'dist']

/** `ST_TextVerticalType` — `a:tcPr/@vert`. */
export const VERT_VALUES: readonly string[] = [
	'horz',
	'vert',
	'vert270',
	'wordArtVert',
	'eaVert',
	'mongolianVert',
	'wordArtVertRtl',
]

/** `ST_TextHorzOverflowType` — `a:tcPr/@horzOverflow`. */
export const HORZ_OVERFLOW_VALUES: readonly string[] = ['clip', 'overflow']

/**
 * Check a value against its `ST_` union, throwing when it is outside.
 *
 * Throwing rather than warn-and-drop, which is the opposite of what the *write* path does
 * with the same enums — and deliberately so. On the write path an option comes from a deck
 * being built, and dropping one value is better than failing a whole build; here it comes
 * from a caller editing one attribute, so silently doing nothing would leave them looking at
 * an unchanged deck with no explanation.
 *
 * @param {string} value - the value the caller asked for
 * @param {readonly string[]} valid - the enum's members
 * @param {string} attribute - the attribute's name, for the message
 * @param {InvalidOptionErrorCode} code - the condition to report
 * @returns {string} `value`, when it is legal
 */
export function checkEnum(
	value: string,
	valid: readonly string[],
	attribute: string,
	code: InvalidOptionErrorCode
): string {
	if (valid.includes(value)) return value
	throw new InvalidOptionError(
		code,
		`Invalid ${attribute}: ${JSON.stringify(value)}. Expected one of: ${valid.join(', ')}.`
	)
}

/**
 * Get or create a child of `a:tcPr` at its schema position.
 * @param {Element} tcPr - the cell's properties element
 * @param {string} qname - the child to get or create
 * @returns {Element} the existing or newly-inserted child
 */
export function tcPrChild(tcPr: Element, qname: string): Element {
	return getOrAddChild(tcPr, qname, TCPR_AFTER[qname] ?? [])
}

/**
 * Insert an already-built subtree into `a:tcPr` at its schema position.
 * @param {Element} tcPr - the cell's properties element
 * @param {string} qname - the qname of the node being inserted (decides where it goes)
 * @param {Element} node - the subtree to insert
 */
export function insertTcPrChild(tcPr: Element, qname: string, node: Element): void {
	insertInOrder(tcPr, node, TCPR_AFTER[qname] ?? [])
}

/**
 * Round a measurement to whole EMU, rejecting a value that cannot be written.
 * `NaN`/`Infinity` would reach the attribute verbatim and make the part invalid, which is
 * exactly the "silent coercion of invalid input" the project's policy rules out.
 */
export function checkFiniteEmu(value: number, field: string, code: InvalidOptionErrorCode): number {
	if (!Number.isFinite(value))
		throw new InvalidOptionError(code, `${field} must be a finite number of EMU, got: ${String(value)}`)
	return Math.round(value)
}
