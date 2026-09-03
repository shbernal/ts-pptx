/**
 * Thin wrapper over `@xmldom/xmldom` — the only module in `src/read/` that
 * imports it directly.
 *
 * `@xmldom/xmldom` ships its own DOM types that are NOT assignable to the
 * lib.dom `Document`/`Element`; all of `src/read/` must import DOM types from
 * here, never from lib.dom and never from `@xmldom/xmldom` directly.
 */
import {
	DOMParser,
	MIME_TYPE,
	XMLSerializer,
	onErrorStopParsing,
	type Document,
	type Element,
	type Node,
} from '@xmldom/xmldom'

export type { Document, Element, Node } from '@xmldom/xmldom'
import { InternalError, InvalidOptionError } from '../../errors.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'
import { boolValue } from '../../ooxml/xsd-boolean.js'
import { PERCENT_SCALE } from '../../units.js'

/** DOM `Node.ELEMENT_NODE` constant (xmldom does not expose it statically). */
export const ELEMENT_NODE = 1

/** Parse an XML string strictly: any well-formedness error throws. */
export function parseXml(xml: string): Document {
	return new DOMParser({ onError: onErrorStopParsing }).parseFromString(xml, MIME_TYPE.XML_TEXT)
}

/**
 * Serialize a Document to an XML string. Does not pretty-print or normalize
 * whitespace. The XML declaration is not part of the DOM; callers that need
 * it (see `Part.serialize`) must prepend it themselves.
 */
export function serializeXml(doc: Document): string {
	return new XMLSerializer().serializeToString(doc)
}

// The registry itself lives in `ooxml/namespaces.ts`, which has no runtime imports, so the
// write side can reach it without pulling `@xmldom/xmldom` into the write-only bundle. Re-exported
// here so every existing `from '.../oxml/dom.js'` import keeps working. `qn` stays in this module:
// it raises a diagnostic, and giving the import-free registry a dependency on `errors.ts` would
// cost exactly what moving the table there bought.
export { OOXML_NS } from '../../ooxml/namespaces.js'

/** Build a prefixed qname string, e.g. `qn('p', 'sld')` → `"p:sld"`. */
export function qn(prefix: string, local: string): string {
	if (!(prefix in OOXML_NS))
		throw new InvalidOptionError('oxml/unknown-namespace-prefix', `Unknown OOXML namespace prefix: ${prefix}`)
	return `${prefix}:${local}`
}

function splitQName(qname: string): { uri: string; local: string } {
	const colon = qname.indexOf(':')
	if (colon < 0)
		throw new InvalidOptionError('oxml/invalid-qname', `Expected a prefixed qname like "p:sld", got: ${qname}`)
	const prefix = qname.slice(0, colon)
	const uri = (OOXML_NS as Record<string, string>)[prefix]
	if (!uri) throw new InvalidOptionError('oxml/unknown-namespace-prefix', `Unknown OOXML namespace prefix: ${prefix}`)
	return { uri, local: qname.slice(colon + 1) }
}

/**
 * Direct child elements of `parent` matching a qname. Matching is by
 * namespace URI + local name, so it is independent of the prefixes the
 * document happens to declare.
 */
export function getElements(parent: Node, qname: string): Element[] {
	const { uri, local } = splitQName(qname)
	const out: Element[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) {
			const element = node as Element
			if (element.localName === local && element.namespaceURI === uri) out.push(element)
		}
	}
	return out
}

/** All direct child elements in document order, regardless of qname. */
export function childElements(parent: Node): Element[] {
	const out: Element[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) out.push(node as Element)
	}
	return out
}

/**
 * All *descendant* elements of `root` matching a namespace + local name, snapshotted into an
 * array. Unlike {@link getElements} (direct children only) this searches the whole subtree, and
 * unlike the live `HTMLCollection` that `getElementsByTagNameNS` returns, the result does not
 * change as the caller mutates the tree — which is what makes it safe to iterate while replacing
 * the very elements being iterated.
 *
 * Note the DOM's own exclusion: `root` itself is never a candidate, only its descendants.
 */
export function descendantsByTag(root: Element, ns: string, local: string): Element[] {
	const out: Element[] = []
	for (const element of root.getElementsByTagNameNS(ns, local)) out.push(element)
	return out
}

/** First direct child element matching a qname, or `null`. */
export function firstChild(parent: Node, qname: string): Element | null {
	const { uri, local } = splitQName(qname)
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) {
			const element = node as Element
			if (element.localName === local && element.namespaceURI === uri) return element
		}
	}
	return null
}

/** First direct child *element* of `parent` (skipping text/comment nodes), or `null`. */
export function firstChildElement(parent: Node): Element | null {
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) return node as Element
	}
	return null
}

/** First direct child element matching any of the given qnames, or `null`. */
function firstChildMatchingAny(parent: Node, qnames: readonly string[]): Element | null {
	const wanted = qnames.map(splitQName)
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const element = node as Element
		if (wanted.some(({ uri, local }) => element.localName === local && element.namespaceURI === uri)) return element
	}
	return null
}

// --- Mutation helpers (editing) --------------------------------------------
//
// `src/read/` mutates the live DOM in place; these are the only sanctioned way
// to create elements, set attributes, and keep a parent's children in the
// sequence OOXML requires. They never mark a part dirty — callers own that.

/** Create a namespaced element from a prefixed qname (e.g. `createElement(doc, 'a:off')`). */
export function createElement(doc: Document, qname: string): Element {
	const { uri } = splitQName(qname)
	return doc.createElementNS(uri, qname)
}

/**
 * The owning document of a node. `Node.ownerDocument` is typed `Document | null`
 * (it is null only for a `Document` itself), but every element parsed from or
 * created within a package document has one; this encodes that invariant and
 * throws loudly if it is ever violated rather than deferring to a downstream NPE.
 */
export function ownerDocumentOf(node: Node): Document {
	const doc = node.ownerDocument
	if (!doc) throw new InternalError('oxml/node-has-no-document', 'Node has no ownerDocument')
	return doc
}

/**
 * Replace `oldNode` with `newNode` in `oldNode`'s parent. `Node.parentNode` is
 * typed `Node | null`; a node reached by walking a tree always has a parent, so
 * this captures the common replace-in-place pattern and fails clearly when the
 * node is unexpectedly detached.
 */
export function replaceInParent(oldNode: Node, newNode: Node): void {
	const parent = oldNode.parentNode
	if (!parent) throw new InternalError('oxml/node-has-no-parent', 'Node has no parent to replace within')
	parent.replaceChild(newNode, oldNode)
}

/**
 * Set an attribute by qname. An unprefixed name (`sz`, `x`) sets a plain
 * attribute; a prefixed name (`r:id`) resolves the prefix to its namespace and
 * sets it via `setAttributeNS`. The reserved `xml:` prefix is handled by the
 * DOM itself.
 */
export function setAttr(element: Element, qname: string, value: string): void {
	const colon = qname.indexOf(':')
	if (colon < 0 || qname.startsWith('xml:')) {
		element.setAttribute(qname, value)
		return
	}
	const { uri } = splitQName(qname)
	element.setAttributeNS(uri, qname, value)
}

/** Remove an attribute by qname; a no-op when the attribute is absent. */
export function removeAttr(element: Element, qname: string): void {
	const colon = qname.indexOf(':')
	if (colon < 0 || qname.startsWith('xml:')) {
		element.removeAttribute(qname)
		return
	}
	const { uri, local } = splitQName(qname)
	element.removeAttributeNS(uri, local)
}

/**
 * Get the first child element matching `qname`, creating and inserting it when
 * absent. A newly created element is inserted before the first existing child
 * whose qname appears in `before` (the schema successors of the new element),
 * or appended when none are present — keeping the parent's children in the
 * sequence order OOXML mandates.
 */
export function getOrAddChild(parent: Element, qname: string, before: readonly string[] = []): Element {
	const existing = firstChild(parent, qname)
	if (existing) return existing
	const doc = parent.ownerDocument
	if (!doc)
		throw new InternalError(
			'oxml/node-has-no-document',
			`Cannot create <${qname}>: parent element has no owner document`
		)
	const child = createElement(doc, qname)
	const successor = before.length ? firstChildMatchingAny(parent, before) : null
	parent.insertBefore(child, successor) // insertBefore(node, null) appends
	return child
}

/**
 * Insert an already-built `node` into `parent` in document order: before the
 * first existing child whose qname is in `before` (the new node's schema
 * successors), or appended when none are present. Unlike {@link getOrAddChild}
 * this always inserts the given node (no get-or-create), so callers can place a
 * freshly-constructed subtree at the right position.
 */
export function insertInOrder(parent: Element, node: Node, before: readonly string[] = []): void {
	const successor = before.length ? firstChildMatchingAny(parent, before) : null
	parent.insertBefore(node, successor)
}

/** Remove every direct child element matching any of the given qnames. */
export function removeChildrenByQName(parent: Element, qnames: readonly string[]): void {
	const toRemove: Element[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		toRemove.push(node as Element)
	}
	const wanted = qnames.map(splitQName)
	for (const element of toRemove) {
		if (wanted.some(({ uri, local }) => element.localName === local && element.namespaceURI === uri)) {
			parent.removeChild(element)
		}
	}
}

/**
 * Read an attribute by qname. An unprefixed name (`sz`, `b`) reads the plain
 * attribute; a prefixed name (`r:id`, `r:embed`) resolves the prefix to its
 * namespace URI and reads via `getAttributeNS`, so it is robust to whatever
 * prefix the document declared. Returns `null` when the attribute is absent.
 */
export function attr(element: Element, qname: string): string | null {
	const colon = qname.indexOf(':')
	if (colon < 0) return element.getAttribute(qname)
	const { uri, local } = splitQName(qname)
	return element.getAttributeNS(uri, local)
}

/**
 * Parse a numeric OOXML attribute; `null`/empty/non-finite → `null`.
 *
 * Not an integer parser: the schema types it reads span `xsd:int`, `xsd:double`
 * and the fractional scales chartEx writes, and every one of them goes through
 * the same `Number()`.
 */
export function numberValue(value: string | null): number | null {
	if (value === null || value === '') return null
	const number = Number(value)
	return Number.isFinite(number) ? number : null
}

/** {@link numberValue} over an attribute read by qname. */
export function numberAttr(element: Element, qname: string): number | null {
	return numberValue(attr(element, qname))
}

// The parser itself lives in `ooxml/xsd-boolean.ts` for the same reason the namespace registry
// does: it is a fact about the schema, it has no runtime imports, and code outside `read/` should
// not have to take a `@xmldom/xmldom` dependency to reach it. Re-exported here so every existing
// `from '.../oxml/dom.js'` import keeps working.
export { boolValue } from '../../ooxml/xsd-boolean.js'

/** {@link boolValue} over an attribute read by qname. */
export function boolAttr(element: Element, qname: string): boolean | null {
	return boolValue(attr(element, qname))
}

/**
 * DrawingML percentage → fraction, or `null` when unparseable.
 *
 * `a:ST_Percentage` is a *union* in the Transitional profile: the fixed-point integer form
 * Office writes (`100%` → `100000`) and a decimal string with a literal `%` (`-?[0-9]+(\.[0-9]+)?%`),
 * which is the only form the Strict profile has. Both are read here — a reader that took only the
 * first dropped a schema-legal value silently, and the two are one `endsWith` apart.
 * @param value - the raw attribute value, or `null` when the attribute is absent
 */
export function parsePercent(value: string | null): number | null {
	if (value === null || value === '') return null
	if (value.endsWith('%')) {
		const n = Number(value.slice(0, -1))
		return Number.isFinite(n) ? n / 100 : null
	}
	const n = Number(value)
	return Number.isFinite(n) ? n / PERCENT_SCALE : null
}

/** {@link parsePercent} over an attribute read by qname. */
export function pctAttr(element: Element, qname: string): number | null {
	return parsePercent(attr(element, qname))
}

/**
 * The text of a DrawingML rich-text body: every `a:t` under `rich`, concatenated, or `null` when
 * there is none.
 *
 * `null` rather than `''` for an empty body, because a title element with no runs and no title
 * element at all are the same fact to a caller. The two chart readers (`c:rich` and `cx:rich`)
 * were the only occurrences of this loop in `src/` and each had its own copy.
 * @param rich - the `a:*`-bearing rich-text element, or `null`
 */
export function concatDrawingMLText(rich: Element | null): string | null {
	if (!rich) return null
	let out = ''
	for (const t of rich.getElementsByTagNameNS(OOXML_NS.a, 't')) out += t.textContent ?? ''
	return out === '' ? null : out
}
