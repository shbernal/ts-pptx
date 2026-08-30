/**
 * Ordered mutation of a table cell's `a:tcPr`, and the cell-scoped schema facts every such
 * edit needs.
 *
 * `CT_TableCellProperties` declares its children as a **sequence**, so an append-only
 * implementation produces an out-of-order `a:tcPr` — which PowerPoint reports as a corrupt
 * file rather than as a bad edit. Every insertion therefore goes through
 * {@link getOrAddChild} / {@link insertInOrder} with the new element's *successors*, which is
 * what `TCPR_AFTER` tabulates; it and the sequence it derives from live in
 * `src/ooxml/sequence.ts` alongside the other complexType orderings.
 *
 * The attribute values are the other half. A value outside its `ST_` union is equally fatal
 * and equally silent, so each is checked against the enum before it is written — the same
 * validate-at-the-boundary rule the write path follows, with the difference that here there
 * is a caller to throw at rather than a deck to degrade (see `ooxml/check-enum.ts`).
 */
import { getOrAddChild, insertInOrder, type Element } from '../oxml/dom.js'
import { TCPR_AFTER } from '../../ooxml/sequence.js'
import { TEXT_ANCHORS, TEXT_HORZ_OVERFLOW, TEXT_VERTICAL } from '../../ooxml/st-enums.js'

export { TCPR_AFTER, TCPR_SEQUENCE } from '../../ooxml/sequence.js'
export { checkEnumOrThrow as checkEnum } from '../../ooxml/check-enum.js'

/** `ST_TextAnchoringType` — `a:tcPr/@anchor`. */
export const ANCHOR_VALUES = TEXT_ANCHORS
/** `ST_TextVerticalType` — `a:tcPr/@vert`. */
export const VERT_VALUES = TEXT_VERTICAL
/** `ST_TextHorzOverflowType` — `a:tcPr/@horzOverflow`. */
export const HORZ_OVERFLOW_VALUES = TEXT_HORZ_OVERFLOW

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
