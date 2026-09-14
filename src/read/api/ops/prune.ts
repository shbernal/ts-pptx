/**
 * Orphan pruning after a part stops being referenced.
 *
 * Removing a slide can strand the parts only it used — its notes slide, media, chart parts. This
 * walks that fringe: a part is dropped when nothing left in the package (or the package root)
 * still resolves an internal relationship to it, and the parts *it* referenced are then examined
 * the same way. Shared chrome (masters, layouts, themes, the notes master) is exempt, so a deck
 * that momentarily has no slides does not lose its design.
 */

import type { Element } from '@xmldom/xmldom'
import type { Relationships } from '../../opc/relationships.js'
import type { Presentation } from '../presentation.js'
import { OOXML_NS } from '../../../ooxml/namespaces.js'
import { descendantsByTag } from '../../oxml/dom.js'
import { SHARED_PARTS } from '../../../ooxml/rel-types.js'

/**
 * Content types that are shared deck chrome, not owned by any one slide: the `deck` rows of
 * {@link SHARED_PARTS}, reached through the presentation → master → layout → theme graph or named
 * by every comment part. {@link Presentation.removeSlide} never prunes these as a removed slide's
 * orphan, even while momentarily unreferenced.
 */
const SHARED_CHROME_CONTENT_TYPES: ReadonlySet<string> = new Set(
	SHARED_PARTS.flatMap((kind) => (kind.scope === 'deck' && kind.contentType !== null ? [kind.contentType] : []))
)

/**
 * Remove `partName` if it is neither shared chrome nor still referenced by any
 * remaining part, then recurse into the parts it referenced. The pruning a
 * removed slide triggers (notes/media/charts the slide alone used).
 * @returns every partname removed, `partName` first when it was
 */
export function pruneIfOrphan(pres: Presentation, partName: string): string[] {
	const part = pres.opc.part(partName)
	if (!part || SHARED_CHROME_CONTENT_TYPES.has(part.contentType)) return []
	if (isReferenced(pres, partName)) return []
	const rels = pres.opc.relationshipsFor(partName)
	const childTargets = [...rels].filter((rel) => rel.targetMode !== 'External').map((rel) => rels.resolveTarget(rel.id))
	pres.opc.removePart(partName)
	const removed = [partName]
	for (const child of childTargets) removed.push(...pruneIfOrphan(pres, child))
	return removed
}

/** A part that still held a relationship to a removed part, and whether it was unlinked. */
export interface InboundLink {
	/** The part holding the relationship, or `/` for the package root. */
	readonly referrer: string
	/** `true` when the relationship and the link elements naming it were removed; `false` when it was kept. */
	readonly unlinked: boolean
}

/**
 * Unlink what still points at `partName`, a part that has left the package, where that can be done
 * without leaving markup behind that names a relationship which is gone.
 *
 * A relationship left pointing at the removed name dangles, and the next part given that name
 * silently becomes its target. Removing only the relationship is worse: PowerPoint refuses a
 * package whose `a:hlinkClick` names an `r:id` the part does not hold. So a relationship is
 * removed when nothing but click and hover links name it, and those links go with it; the run or
 * shape carrying them stays. A relationship any other element names is kept, because that element
 * cannot be dropped the same way. A custom show's `p:sld` is not one of them: removing a slide
 * drops it first ({@link dropFromShowsAndSections}).
 *
 * Dropping the links is a deliberate departure from PowerPoint, which keeps a jump link to a
 * deleted slide and, once it renumbers the slides, lets it jump to whichever slide takes the
 * deleted one's part name (`test/read/fixtures/slide-jump-link-target-deleted.pptx`).
 * @returns each part that held such a relationship, in package order
 */
export function unlinkInbound(pres: Presentation, partName: string): InboundLink[] {
	const found: InboundLink[] = []
	for (const owner of relationshipOwners(pres)) {
		const rels = pres.opc.relationshipsFor(owner)
		const inbound = [...rels].filter((rel) => rel.targetMode !== 'External' && rels.resolveTarget(rel.id) === partName)
		if (inbound.length === 0) continue
		const part = owner === '/' ? undefined : pres.opc.part(owner)
		const root = part?.isXmlPart ? part.dom.documentElement : null
		let unlinked = root !== null
		for (const rel of inbound) {
			const references = root ? elementsNamingRel(root, rel.id) : []
			if (!root || !references.every(isLinkElement)) {
				unlinked = false
				continue
			}
			for (const element of references) element.parentNode?.removeChild(element)
			if (references.length > 0) part?.markDirty()
			rels.remove(rel.id)
		}
		found.push({ referrer: owner, unlinked })
	}
	return found
}

/**
 * Take a slide that is leaving the package out of the presentation's custom shows and sections.
 *
 * PowerPoint does this when a slide is deleted: the slide's `p:sld` leaves every custom show and
 * its `p14:sldId` leaves every section, and a show or section left empty stays
 * (`test/read/fixtures/slide-jump-link-target-deleted.pptx`). Call it while the presentation's
 * relationships still resolve to the slide. A relationship nothing names once the entries are gone
 * is removed here, so {@link unlinkInbound} does not report it as a reference it had to keep.
 * @param root - the presentation part's root element
 * @param rels - the presentation part's relationships
 * @param partName - the slide leaving the package
 * @param slideId - its `p:sldId/@id`
 * @returns whether anything was removed
 */
export function dropFromShowsAndSections(
	root: Element,
	rels: Relationships,
	partName: string,
	slideId: number | null
): boolean {
	const toSlide = new Set(
		[...rels]
			.filter((rel) => rel.targetMode !== 'External' && rels.resolveTarget(rel.id) === partName)
			.map((rel) => rel.id)
	)
	let changed = false
	for (const sld of descendantsByTag(root, OOXML_NS.p, 'sld')) {
		if (!toSlide.has(sld.getAttributeNS(OOXML_NS.r, 'id') ?? '')) continue
		sld.parentNode?.removeChild(sld)
		changed = true
	}
	if (slideId !== null) {
		for (const sldId of descendantsByTag(root, OOXML_NS.p14, 'sldId')) {
			if (sldId.getAttribute('id') !== String(slideId)) continue
			sldId.parentNode?.removeChild(sldId)
			changed = true
		}
	}
	if (changed) {
		for (const relId of toSlide) if (elementsNamingRel(root, relId).length === 0) rels.remove(relId)
	}
	return changed
}

/** Link elements that can be dropped with the relationship they name, leaving their run or shape intact. */
const LINK_ELEMENTS: ReadonlySet<string> = new Set(['hlinkClick', 'hlinkMouseOver'])

function isLinkElement(element: Element): boolean {
	return element.namespaceURI === OOXML_NS.a && LINK_ELEMENTS.has(element.localName ?? '')
}

/** Every element under `root`, `root` included, with an attribute in the relationships namespace naming `relId`. */
function elementsNamingRel(root: Element, relId: string): Element[] {
	const out: Element[] = []
	for (const element of [root, ...descendantsByTag(root, '*', '*')]) {
		for (let i = 0; i < element.attributes.length; i++) {
			const attribute = element.attributes[i]
			if (attribute?.namespaceURI === OOXML_NS.r && attribute.value === relId) {
				out.push(element)
				break
			}
		}
	}
	return out
}

/** Whether any remaining part (or the package root) resolves an internal relationship to `partName`. */
function isReferenced(pres: Presentation, partName: string): boolean {
	for (const owner of relationshipOwners(pres)) {
		const rels = pres.opc.relationshipsFor(owner)
		for (const rel of rels) {
			if (rel.targetMode === 'External') continue
			if (rels.resolveTarget(rel.id) === partName) return true
		}
	}
	return false
}

/** Every partname that can own relationships: each part that is not itself a `.rels` part, and the package root. */
function relationshipOwners(pres: Presentation): string[] {
	return [...[...pres.opc.parts.keys()].filter((partName) => !partName.endsWith('.rels')), '/']
}
