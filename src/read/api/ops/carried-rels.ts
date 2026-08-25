/**
 * Relationship rewriting for content carried between packages.
 *
 * When a shape or decoration is deep-copied out of a source deck, every relationship
 * reference it carries (`r:embed` on a picture fill, `r:id` on a chart frame, `r:link` on a
 * linked image, …) still names an id in the *source* part's relationship list. Left alone
 * those ids either dangle or — worse — silently resolve to whatever the destination part
 * happens to have under the same id. This rewrites each one to a fresh relationship on the
 * destination part, copying the referenced part into the destination package on first sight.
 *
 * Both functions are pure with respect to the deck: everything they touch arrives as an
 * argument. They were private methods on `Presentation` and used no instance state, which
 * made the import subsystem look more entangled with the read model than it is.
 */

import { OOXML_NS, setAttr, type Element } from '../../oxml/dom.js'
import type { Relationships } from '../../opc/relationships.js'
import { InvalidOptionError } from '../../../errors.js'
import { relativePartName } from '../../opc/partnames.js'
import { collectElements } from '../../oxml/slide-dom.js'
import { copyPart, type ImportContext, type OwnedScope } from './part-copy.js'
import { isSharedByPageCopies } from './page-owned.js'

/**
 * Rewrite every relationship reference inside a carried subtree to a fresh destination-local id.
 * @param {Element} node - root of the carried subtree (rewritten in place)
 * @param {ImportContext} ctx - the open import (destination deck + source package + copy registry)
 * @param {Relationships} sourceRels - relationships of the source part the subtree came from
 * @param {string} destPartName - partname of the destination part the subtree is being placed in
 * @param {Relationships} destRels - relationships of that destination part
 * @param {Map<string, string>} relIdMap - per-call cache (source part + source id → new id) deduping references shared within one import
 * @param {OwnedScope} [owned] - ownership scope of the thing being carried, so the chart or diagram under a carried shape is copied for it rather than shared with an earlier copy
 */
export function rewriteCarriedRels(
	node: Element,
	ctx: ImportContext,
	sourceRels: Relationships,
	destPartName: string,
	destRels: Relationships,
	relIdMap: Map<string, string>,
	owned?: OwnedScope
): void {
	const elements: Element[] = []
	collectElements(node, elements)
	for (const el of elements) {
		const refs: { local: string; id: string }[] = []
		const attrs = el.attributes
		for (let i = 0; i < attrs.length; i++) {
			const a = attrs.item(i)
			if (!a || a.namespaceURI !== OOXML_NS.r || !a.value) continue
			if (!sourceRels.get(a.value)) continue // an r-namespaced attribute that isn't a relationship id
			refs.push({ local: a.localName ?? a.name, id: a.value })
		}
		for (const { local, id } of refs) {
			setAttr(el, `r:${local}`, carryRel(ctx, sourceRels, id, destPartName, destRels, relIdMap, owned))
		}
	}
}

/**
 * Resolve one carried relationship to a fresh destination-local id, copying its internal target.
 * @param {ImportContext} ctx - the open import
 * @param {Relationships} sourceRels - relationships of the source part
 * @param {string} id - the source relationship id to carry
 * @param {string} destPartName - partname of the destination part
 * @param {Relationships} destRels - relationships of that destination part
 * @param {Map<string, string>} relIdMap - per-call dedupe cache
 * @param {OwnedScope} [owned] - ownership scope of the carried content, for the parts it owns
 * @return {string} the new relationship id on the destination part
 */
function carryRel(
	ctx: ImportContext,
	sourceRels: Relationships,
	id: string,
	destPartName: string,
	destRels: Relationships,
	relIdMap: Map<string, string>,
	owned?: OwnedScope
): string {
	const rel = sourceRels.get(id)
	if (!rel)
		throw new InvalidOptionError(
			'relationship/not-found',
			`Relationships of ${sourceRels.sourcePartName}: no relationship with id ${id}`
		)
	// The dedupe cache is for what may be shared. A part the carried content owns —
	// the chart under a copied frame, the diagram under a copied SmartArt — is copied
	// into its scope instead, so carrying the same shape twice does not hand two
	// frames one chart part, which is a deck PowerPoint refuses to open.
	const ownsIt = owned !== undefined && rel.targetMode !== 'External' && !isSharedByPageCopies(rel.type)
	const key = `${sourceRels.sourcePartName}|${id}`
	if (!ownsIt) {
		const cached = relIdMap.get(key)
		if (cached) return cached
	}
	const newId =
		rel.targetMode === 'External'
			? destRels.add(rel.type, rel.target, 'External').id
			: destRels.add(
					rel.type,
					relativePartName(destPartName, copyPart(ctx, sourceRels.resolveTarget(id), ownsIt ? owned : undefined))
				).id
	if (!ownsIt) relIdMap.set(key, newId)
	return newId
}
