/**
 * Thin wrapper over `@xmldom/xmldom` — the only module in `src/read/` that
 * imports it directly.
 *
 * `@xmldom/xmldom` ships its own DOM types that are NOT assignable to the
 * lib.dom `Document`/`Element`; all of `src/read/` must import DOM types from
 * here, never from lib.dom and never from `@xmldom/xmldom` directly.
 */
import { DOMParser, MIME_TYPE, XMLSerializer, onErrorStopParsing } from '@xmldom/xmldom'

export type { Document, Element, Node } from '@xmldom/xmldom'
import type { Document, Element, Node } from '@xmldom/xmldom'

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

/** Canonical OOXML prefix → namespace URI registry. */
export const OOXML_NS: Readonly<Record<string, string>> = Object.freeze({
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
	cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
	ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
	dc: 'http://purl.org/dc/elements/1.1/',
	dcterms: 'http://purl.org/dc/terms/',
	dgm: 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
	ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
	p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
	pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
	r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
})

/** Build a prefixed qname string, e.g. `qn('p', 'sld')` → `"p:sld"`. */
export function qn(prefix: string, local: string): string {
	if (!(prefix in OOXML_NS)) throw new Error(`Unknown OOXML namespace prefix: ${prefix}`)
	return `${prefix}:${local}`
}

function splitQName(qname: string): { uri: string; local: string } {
	const colon = qname.indexOf(':')
	if (colon < 0) throw new Error(`Expected a prefixed qname like "p:sld", got: ${qname}`)
	const prefix = qname.slice(0, colon)
	const uri = OOXML_NS[prefix]
	if (!uri) throw new Error(`Unknown OOXML namespace prefix: ${prefix}`)
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

/** First direct child element matching any of the given qnames, or `null`. */
function firstChildMatchingAny(parent: Node, qnames: string[]): Element | null {
	const wanted = qnames.map(splitQName)
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType !== ELEMENT_NODE) continue
		const element = node as Element
		if (wanted.some(({ uri, local }) => element.localName === local && element.namespaceURI === uri)) return element
	}
	return null
}

// --- Mutation helpers (Phase 3 editing) -----------------------------------
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
export function getOrAddChild(parent: Element, qname: string, before: string[] = []): Element {
	const existing = firstChild(parent, qname)
	if (existing) return existing
	const doc = parent.ownerDocument
	if (!doc) throw new Error(`Cannot create <${qname}>: parent element has no owner document`)
	const child = createElement(doc, qname)
	const successor = before.length ? firstChildMatchingAny(parent, before) : null
	parent.insertBefore(child, successor) // insertBefore(node, null) appends
	return child
}

/** Remove every direct child element matching any of the given qnames. */
export function removeChildrenByQName(parent: Element, qnames: string[]): void {
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

/** Parse an integer-valued OOXML attribute; `null`/empty/non-finite → `null`. */
export function intValue(value: string | null): number | null {
	if (value === null || value === '') return null
	const number = Number(value)
	return Number.isFinite(number) ? number : null
}

/** Parse an `xsd:boolean` OOXML attribute (`1`/`0`/`true`/`false`); else `null`. */
export function boolValue(value: string | null): boolean | null {
	if (value === '1' || value === 'true') return true
	if (value === '0' || value === 'false') return false
	return null
}

/** Escape a string for use inside a double-quoted XML attribute value. */
export function escapeXmlAttribute(value: string): string {
	return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;')
}
