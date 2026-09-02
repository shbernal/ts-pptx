/**
 * Shape-element plumbing: the small element lookups every shape kind performs, plus the
 * shape-scoped re-exports of the shared schema-successor lists.
 *
 * The successor arrays are the reason a setter can create a missing `a:xfrm` / `a:solidFill` /
 * `a:ln` without corrupting the part: `getOrAddChild` inserts before the first listed sibling it
 * finds, so each list must name exactly the children that legally *follow* the one being added.
 * They are declared once in `src/ooxml/sequence.ts` (derived from each complexType's sequence)
 * and re-exported here so the shape modules keep importing them from one place.
 */

import {
	ELEMENT_NODE,
	OOXML_NS,
	attr,
	childElements,
	firstChild,
	getOrAddChild,
	numberValue,
	type Element,
} from '../../oxml/dom.js'
import { SHAPE_AFTER_SPPR, SPPR_AFTER_XFRM } from '../../../ooxml/sequence.js'

export {
	GRPSPPR_AFTER_XFRM,
	GRPSPPR_FILL_AFTER,
	LN_FILL_AFTER,
	SHAPE_AFTER_SPPR,
	SPPR_AFTER_XFRM,
	SPPR_FILL_AFTER,
	SPPR_LN_AFTER,
} from '../../../ooxml/sequence.js'

/**
 * A shape's properties element (`p:spPr` / `p:grpSpPr`) paired with the schema
 * successors for ordered insertion of its `a:solidFill` and `a:ln` children.
 * `lnAfter` is `null` for kinds with no `a:ln` (group shapes).
 */
export interface ShapeProperties {
	props: Element
	fillAfter: string[]
	lnAfter: string[] | null
}

/** First `<p:cNvPr>` reached through the shape's non-visual properties wrapper (`p:nv*Pr`). */
export function nonVisualCNvPr(element: Element): Element | null {
	for (let node = element.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const child = node as Element
		if (child.namespaceURI === OOXML_NS.p && child.localName?.startsWith('nv')) {
			return firstChild(child, 'p:cNvPr')
		}
	}
	return null
}

export function emuFrom(parent: Element | null, qname: string, attribute: string): number | null {
	const element = parent && firstChild(parent, qname)
	return element ? numberValue(attr(element, attribute)) : null
}

export { childElements }

/** Get-or-add `p:spPr/a:xfrm` for shapes whose transform lives in `p:spPr` (`p:sp`, `p:pic`, `p:cxnSp`). */
export function getOrAddSpPrXfrm(shapeElement: Element): Element {
	const spPr = getOrAddChild(shapeElement, 'p:spPr', SHAPE_AFTER_SPPR)
	return getOrAddChild(spPr, 'a:xfrm', SPPR_AFTER_XFRM)
}
