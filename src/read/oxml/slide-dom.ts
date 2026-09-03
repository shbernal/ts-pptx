/**
 * Small DOM navigation helpers shared by the read-model slide/import paths.
 *
 * Each is a pure query over a live OOXML node -- no package state, no mutation -- which
 * is why they sit outside `Presentation` rather than as private methods on it.
 */

import type { Part } from '../opc/part.js'
import { ELEMENT_NODE, OOXML_NS, attr, firstChild, firstChildElement, type Element } from './dom.js'

/**
 * The `p:cSld` of a slide/layout/master/notes part root.
 *
 * Takes a nullable root and returns a nullable element so a caller can chain from
 * `part?.dom.documentElement` without a guard of its own — which is what a dozen callers were
 * writing by hand, each in a slightly different spelling (`root && …`, `root ? … : null`).
 */
export function cSldOf(root: Element | null | undefined): Element | null {
	return root ? firstChild(root, 'p:cSld') : null
}

/** The `p:spTree` of a slide/layout/master/notes part root, through its `p:cSld`. */
export function spTreeOf(root: Element | null | undefined): Element | null {
	const cSld = cSldOf(root)
	return cSld ? firstChild(cSld, 'p:spTree') : null
}

/**
 * The `p:nvPr` of a shape, whichever `*nvPr` wrapper its kind uses.
 *
 * The wrapper's name varies by shape kind — `p:nvSpPr` on a `p:sp`, `p:nvPicPr` on a `p:pic`,
 * `p:nvGraphicFramePr`, `p:nvCxnSpPr`, `p:nvGrpSpPr` — but it is always the shape's first child
 * element, and `p:nvPr` is always inside it. Taking the first child element rather than naming
 * the wrapper is what lets one helper serve every kind; the callers that named `p:nvSpPr`
 * silently returned null on a picture.
 */
export function nvPrOf(shape: Element): Element | null {
	const nv = firstChildElement(shape)
	return nv ? firstChild(nv, 'p:nvPr') : null
}

/** The `p:cSld@name` of a slide/layout/master part (`''` when absent). */
export function cSldName(part: Part | undefined): string {
	const cSld = cSldOf(part?.dom.documentElement)
	return (cSld && attr(cSld, 'name')) ?? ''
}

/**
 * Whether a `p:spTree` child is one of the tree's own children rather than a shape:
 * `p:nvGrpSpPr` and `p:grpSpPr` before the shapes, `p:extLst` after them.
 *
 * `CT_GroupShape` sequences `nvGrpSpPr, grpSpPr, (shape)*, extLst?`, so all three are
 * positional and none of them is something a shape walk should return. Three helpers below
 * asked this question and only one of them excluded `p:extLst`, so `carriedDecorations`
 * reported it as a decoration to carry -- and the slide importer inserts each of those BEFORE
 * the destination's shapes, which puts an `extLst` where the content model has none and makes
 * the part invalid. The read corpus holds no direct `p:spTree/p:extLst` at all, across 776
 * shape trees, so nothing existing could have caught it.
 */
function isSpTreeOwnChild(el: Element): boolean {
	return (
		el.namespaceURI === OOXML_NS.p &&
		(el.localName === 'nvGrpSpPr' || el.localName === 'grpSpPr' || el.localName === 'extLst')
	)
}

/** Whether a `p:spTree` child is a placeholder shape (its `*nvPr` carries a `p:ph`). */
function isPlaceholderShape(shape: Element): boolean {
	const nvPr = nvPrOf(shape)
	return !!(nvPr && firstChild(nvPr, 'p:ph'))
}

/** The decorative shapes on a layout/master `p:spTree`: every shape child except placeholders. */
export function carriedDecorations(root: Element | null): Element[] {
	const spTree = spTreeOf(root)
	if (!spTree) return []
	const out: Element[] = []
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (isSpTreeOwnChild(el) || isPlaceholderShape(el)) continue
		out.push(el)
	}
	return out
}

/**
 * The `n`-th shape child of a `p:spTree` in document (z-)order, skipping the tree's own
 * children. Returns `null` when `n` is past the last shape (the caller then appends).
 */
export function nthShapeChild(spTree: Element, n: number): Element | null {
	let i = 0
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (isSpTreeOwnChild(el)) continue
		if (i === n) return el
		i++
	}
	return null
}

/** The first shape child of a `p:spTree` (skipping the tree's own children), or `null`. */
export function firstShapeChild(spTree: Element): Element | null {
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (!isSpTreeOwnChild(el)) return el
	}
	return null
}

/** Collect `node` and all its descendant elements (document order) into `out`. */
export function collectElements(node: Element, out: Element[]): void {
	out.push(node)
	for (let child = node.firstChild; child; child = child.nextSibling) {
		if (child.nodeType === ELEMENT_NODE) collectElements(child as Element, out)
	}
}
