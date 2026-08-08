/**
 * Small DOM navigation helpers shared by the read-model slide/import paths.
 *
 * Each is a pure query over a live OOXML node -- no package state, no mutation -- which
 * is why they sit outside `Presentation` rather than as private methods on it.
 */

import type { Part } from '../opc/part.js'
import { ELEMENT_NODE, OOXML_NS, attr, firstChild, firstChildElement, type Element } from './dom.js'

/** The `p:cSld@name` of a slide/layout/master part (`''` when absent). */
export function cSldName(part: Part | undefined): string {
	const root = part?.dom.documentElement
	const cSld = root && firstChild(root, 'p:cSld')
	return (cSld && attr(cSld, 'name')) ?? ''
}

/** Whether `el` is a `p:spTree`'s own group properties (`p:nvGrpSpPr`/`p:grpSpPr`), not a shape. */
function isSpTreeProperty(el: Element): boolean {
	return el.namespaceURI === OOXML_NS.p && (el.localName === 'nvGrpSpPr' || el.localName === 'grpSpPr')
}

/** Whether a `p:spTree` child is a placeholder shape (its `*nvPr` carries a `p:ph`). */
function isPlaceholderShape(shape: Element): boolean {
	const nv = firstChildElement(shape)
	const nvPr = nv && firstChild(nv, 'p:nvPr')
	return !!(nvPr && firstChild(nvPr, 'p:ph'))
}

/** The decorative shapes on a layout/master `p:spTree`: every shape child except placeholders. */
export function carriedDecorations(root: Element | null): Element[] {
	if (!root) return []
	const cSld = firstChild(root, 'p:cSld')
	const spTree = cSld && firstChild(cSld, 'p:spTree')
	if (!spTree) return []
	const out: Element[] = []
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (isSpTreeProperty(el) || isPlaceholderShape(el)) continue
		out.push(el)
	}
	return out
}

/**
 * The `n`-th shape child of a `p:spTree` in document (z-)order, skipping the
 * tree's own `nvGrpSpPr`/`grpSpPr` and any trailing `p:extLst`. Returns `null`
 * when `n` is past the last shape (the caller then appends).
 */
export function nthShapeChild(spTree: Element, n: number): Element | null {
	let i = 0
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (isSpTreeProperty(el)) continue
		if (el.namespaceURI === OOXML_NS.p && el.localName === 'extLst') continue
		if (i === n) return el
		i++
	}
	return null
}

/** The first shape child of a `p:spTree` (skipping `nvGrpSpPr`/`grpSpPr`), or `null`. */
export function firstShapeChild(spTree: Element): Element | null {
	for (let node = spTree.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const el = node as Element
		if (!isSpTreeProperty(el)) return el
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
