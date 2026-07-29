/**
 * Shape-element plumbing: the schema-successor lists that keep a get-or-added child in document
 * order, and the small element lookups every shape kind performs.
 *
 * The successor arrays are the reason a setter can create a missing `a:xfrm` / `a:solidFill` /
 * `a:ln` without corrupting the part: `getOrAddChild` inserts before the first listed sibling it
 * finds, so each list must name exactly the children that legally *follow* the one being added.
 */

import { ELEMENT_NODE, OOXML_NS, attr, firstChild, getOrAddChild, intValue, type Element } from '../../oxml/dom.js'

// Schema successors used to keep elements in document order when a geometry
// setter has to create one.
export const SPPR_AFTER_XFRM = [
	'a:custGeom',
	'a:prstGeom',
	'a:noFill',
	'a:solidFill',
	'a:gradFill',
	'a:blipFill',
	'a:pattFill',
	'a:grpFill',
	'a:ln',
	'a:effectLst',
	'a:effectDag',
	'a:scene3d',
	'a:sp3d',
	'a:extLst',
]
export const GRPSPPR_AFTER_XFRM = [
	'a:noFill',
	'a:solidFill',
	'a:gradFill',
	'a:blipFill',
	'a:pattFill',
	'a:grpFill',
	'a:effectLst',
	'a:effectDag',
	'a:scene3d',
	'a:extLst',
]
// spPr itself sits before p:style / p:txBody within p:sp (and before p:style
// within p:pic / p:cxnSp); blipFill / nv*Pr precede it and are excluded.
export const SHAPE_AFTER_SPPR = ['p:style', 'p:txBody']

// Successor arrays for inserting a fill / line *into* a properties element.
// Distinct from the *_AFTER_XFRM arrays above, which sequence a:xfrm (the first
// child): a:solidFill and a:ln sit mid-sequence, so their `before` lists must
// contain only the children that legally follow them (CT_ShapeProperties /
// CT_GroupShapeProperties / CT_LineProperties).
export const SPPR_FILL_AFTER = ['a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
export const SPPR_LN_AFTER = ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
export const GRPSPPR_FILL_AFTER = ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:extLst']
export const LN_FILL_AFTER = [
	'a:prstDash',
	'a:custDash',
	'a:round',
	'a:bevel',
	'a:miter',
	'a:headEnd',
	'a:tailEnd',
	'a:extLst',
]

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
	return element ? intValue(attr(element, attribute)) : null
}

/** Direct child *elements* of `parent`, in document order. */
export function childElements(parent: Element): Element[] {
	const out: Element[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) out.push(node as Element)
	}
	return out
}

/** Get-or-add `p:spPr/a:xfrm` for shapes whose transform lives in `p:spPr` (`p:sp`, `p:pic`, `p:cxnSp`). */
export function getOrAddSpPrXfrm(shapeElement: Element): Element {
	const spPr = getOrAddChild(shapeElement, 'p:spPr', SHAPE_AFTER_SPPR)
	return getOrAddChild(spPr, 'a:xfrm', SPPR_AFTER_XFRM)
}
