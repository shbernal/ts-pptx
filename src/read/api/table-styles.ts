/**
 * Presentation-level `tableStyles.xml` merging for imported slides and masters.
 *
 * `tableStyles.xml` is a single deck-wide part that records no ownership -- it does not
 * say which style belongs to which master -- so copying styles across decks is its own
 * traversal rather than part of the master/layout copy chain. Both entry points union
 * `<a:tblStyle>` by `styleId` and are idempotent.
 *
 * The functions take a structural {@link TableStyleTarget} rather than a `Presentation`,
 * so this module stays independent of the class that calls it.
 */

import type { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { relativePartName } from '../opc/partnames.js'
import { OOXML_NS, attr, getElements, ownerDocumentOf, type Element } from '../oxml/dom.js'

const TABLE_STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles'
const TABLE_STYLES_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml'
// The well-known "no style / table grid" GUID PowerPoint uses as a tblStyleLst @def.
const TABLE_STYLES_DEFAULT_GUID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

const textEncoder = new TextEncoder()

/** The destination deck: just the package and its main part. `Presentation` satisfies it. */
export interface TableStyleTarget {
	readonly opc: OpcPackage
	readonly presentationPart: Part
}

/**
 * Copy any table style the restyled slide references from the source
 * `tableStyles.xml` into this deck's (the `remapLiterals` table leg). A restyled
 * table resolves its `@tableStyleId` against the *destination* `tableStyles` and
 * silently falls back when that id is absent; the source `<a:tblStyle>` is itself
 * symbolic (scheme colours), so copying it verbatim under the **same** id lets the
 * table keep its structure while re-branding to this deck's theme — no table-XML
 * rewrite needed. Idempotent: an id this deck already defines is left as-is, and a
 * referenced id with no source definition is skipped. Creates and wires a
 * `tableStyles.xml` part if this deck has none.
 */
export function copySourceTableStyles(dest: TableStyleTarget, sourceOpc: OpcPackage, slideRoot: Element): void {
	const ids = new Set<string>()
	for (const idEl of slideRoot.getElementsByTagNameNS(OOXML_NS.a, 'tableStyleId')) {
		const id = idEl.textContent?.trim()
		if (id) ids.add(id)
	}
	if (ids.size === 0) return

	const sourceList = tableStyleList(sourceOpc)
	if (!sourceList) return
	const sourceStyles = new Map(getElements(sourceList, 'a:tblStyle').map((st) => [attr(st, 'styleId'), st] as const))

	const destPart = ensureTableStylesPart(dest)
	const destList = destPart.dom.documentElement
	if (!destList) return

	const referenced: Element[] = []
	for (const id of ids) {
		const src = sourceStyles.get(id)
		if (src) referenced.push(src)
	}
	if (addTableStyles(destList, referenced)) destPart.markDirty()
}

/**
 * Append each source `<a:tblStyle>` this deck does not already define, keyed by
 * `styleId` — the shared core of {@link copySourceTableStyles} (which offers only
 * the styles a restyled slide references) and {@link carryTableStyles} (which
 * offers the source's whole list). Destination-wins per id: a style this deck
 * already defines is left untouched, so a re-call adds nothing. Returns whether
 * anything was added.
 */
function addTableStyles(destList: Element, styles: Iterable<Element>): boolean {
	const present = new Set(getElements(destList, 'a:tblStyle').map((st) => attr(st, 'styleId')))
	let added = false
	for (const style of styles) {
		const id = attr(style, 'styleId')
		if (!id || present.has(id)) continue
		destList.appendChild(ownerDocumentOf(destList).importNode(style, true))
		present.add(id)
		added = true
	}
	return added
}

/**
 * Copy the source deck's **whole** `tableStyles.xml` into this deck: every
 * `<a:tblStyle>` it defines, plus its `a:tblStyleLst@def` (the default table
 * style). Presentation-level, so a separate traversal from the master/layout copy
 * chain — and whole-deck, since `tableStyles.xml` does not record which style
 * belongs to which master. See `ImportSlideMastersOptions.tableStyles`.
 *
 * Styles union by `styleId` (destination-wins per id, via {@link addTableStyles}),
 * but `def` is **source-wins**: a caller asking for the source's table styling
 * wants its default too, and the destination's `def` is typically a generator stub
 * rather than a deliberate choice. Carrying the styles without the `def` would be
 * the worse half-measure — the standard default GUID is one most templates also
 * define, so a new table would silently resolve to the wrong style rather than
 * visibly to none. Idempotent: a re-call re-adds nothing and re-sets `def` to the
 * same value.
 */
export function carryTableStyles(dest: TableStyleTarget, sourceOpc: OpcPackage): void {
	const sourceList = tableStyleList(sourceOpc)
	if (!sourceList) return

	const destPart = ensureTableStylesPart(dest)
	const destList = destPart.dom.documentElement
	if (!destList) return

	let changed = addTableStyles(destList, getElements(sourceList, 'a:tblStyle'))

	const sourceDef = attr(sourceList, 'def')
	if (sourceDef && attr(destList, 'def') !== sourceDef) {
		destList.setAttribute('def', sourceDef)
		changed = true
	}
	if (changed) destPart.markDirty()
}

/** The `a:tblStyleLst` root of a package's `tableStyles.xml`, or `null` when it has none. */
function tableStyleList(opc: OpcPackage): Element | null {
	const part = opc.partsByContentType(TABLE_STYLES_CONTENT_TYPE)[0]
	return part ? part.dom.documentElement : null
}

/**
 * The deck's `tableStyles.xml` part, creating an empty one (and wiring its
 * `presentation.xml` relationship + content type) when the deck has none.
 */
function ensureTableStylesPart(dest: TableStyleTarget): Part {
	const existing = dest.opc.partsByContentType(TABLE_STYLES_CONTENT_TYPE)[0]
	if (existing) return existing
	const partName = dest.opc.reservePartNameLike('/ppt/tableStyles.xml')
	const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<a:tblStyleLst xmlns:a="${OOXML_NS.a}" def="${TABLE_STYLES_DEFAULT_GUID}"/>`
	const part = dest.opc.addPart(partName, TABLE_STYLES_CONTENT_TYPE, textEncoder.encode(xml))
	const presRels = dest.opc.relationshipsFor(dest.presentationPart.partName)
	presRels.add(TABLE_STYLES_REL, relativePartName(dest.presentationPart.partName, partName))
	return part
}
