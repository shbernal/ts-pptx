/**
 * Slide-master bookkeeping in `presentation.xml` and in a master's own
 * `p:sldLayoutIdLst`: registering a master so renderers treat it as active,
 * promoting masters to the front of the id list, allocating ids out of the shared
 * master/layout id space, and emptying a freshly-copied master's layout list.
 *
 * These are the destination-side half of the import machinery -- they read and
 * write only this deck, never a source package and never the copy registry -- so
 * they take a structural {@link DeckTarget} rather than a `Presentation`. The
 * part-copy traversal that calls most of them still lives on the class.
 */

import { relativePartName } from '../opc/partnames.js'
import {
	attr,
	createElement,
	firstChild,
	getElements,
	getOrAddChild,
	intValue,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import type { DeckTarget } from './deck-target.js'

const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'

// ST_SlideMasterId and ST_SlideLayoutId share this floor.
const MIN_SLIDE_MASTER_ID = 2147483648

/**
 * Wire a freshly-copied slide master into `presentation.xml`: add a
 * presentation→master relationship and a `p:sldMasterId` entry in
 * `p:sldMasterIdLst`. A master that is reachable only through the
 * slide→layout→master rel chain but absent from `p:sldMasterIdLst` is treated
 * as inactive by renderers, so its background/graphics never paint. No-op when
 * the master is already registered (shared across imports from one source).
 */
export function registerMaster(dest: DeckTarget, masterPartName: string): void {
	const presPart = dest.presentationPart
	const presRels = dest.opc.relationshipsFor(presPart.partName)
	for (const rel of presRels.byType(SLIDE_MASTER_REL)) {
		if (presRels.resolveTarget(rel.id) === masterPartName) return
	}
	const relId = presRels.add(SLIDE_MASTER_REL, relativePartName(presPart.partName, masterPartName)).id

	const root = presPart.dom.documentElement
	if (!root) throw new Error('presentation.xml has no document element to register a master in')
	// `p:sldMasterIdLst` is the first child of CT_Presentation; create it before
	// any later sibling if a (degenerate) deck lacks one.
	const lst = getOrAddChild(root, 'p:sldMasterIdLst', [
		'p:notesMasterIdLst',
		'p:handoutMasterIdLst',
		'p:sldIdLst',
		'p:sldSz',
		'p:notesSz',
		'p:embeddedFontLst',
		'p:custShowLst',
		'p:photoAlbum',
		'p:custDataLst',
		'p:kinsoku',
		'p:defaultTextStyle',
		'p:modifyVerifier',
		'p:extLst',
	])
	const entry = createElement(presPart.dom, 'p:sldMasterId')
	setAttr(entry, 'id', String(nextMasterLayoutId(dest)))
	setAttr(entry, 'r:id', relId)
	lst.appendChild(entry)
	presPart.markDirty()
}

/**
 * Move the `p:sldMasterId` entries for `masterPartNames` to the front of
 * `p:sldMasterIdLst`, preserving their existing relative order, so the first of
 * them becomes the deck's `Designs(1)` — the theme PowerPoint's Design tab shows.
 * Reorders the id list only; relationships, ids, and all other parts are left as
 * they are. A no-op when the named masters already lead (idempotent re-call).
 */
export function promoteMasters(dest: DeckTarget, masterPartNames: string[]): void {
	if (masterPartNames.length === 0) return
	const presPart = dest.presentationPart
	const root = presPart.dom.documentElement
	const lst = root && firstChild(root, 'p:sldMasterIdLst')
	if (!lst) return
	const rels = dest.opc.relationshipsFor(presPart.partName)
	const promote = new Set(masterPartNames)
	const entries = getElements(lst, 'p:sldMasterId')
	// Entries whose relationship resolves to a promoted master, kept in their
	// current order, then everyone else in theirs — a stable partition.
	const isPromoted = (entry: Element): boolean => {
		const relId = attr(entry, 'r:id')
		return relId ? promote.has(rels.resolveTarget(relId)) : false
	}
	const desired = [...entries.filter(isPromoted), ...entries.filter((e) => !isPromoted(e))]
	if (desired.every((entry, i) => entry === entries[i])) return
	for (const entry of desired) lst.appendChild(entry)
	presPart.markDirty()
}

/**
 * The next free id in the presentation-wide slide-master / slide-layout id space.
 *
 * `p:sldMasterId/@id` and every master's `p:sldLayoutId/@id` draw from ONE shared
 * id space (both ST_SlideMasterId/ST_SlideLayoutId, floor 2147483648) and must be
 * unique across the WHOLE presentation, not just within their own list. A duplicate
 * anywhere in this space makes PowerPoint report the file as corrupt on open
 * ("found a problem with content" / 0x80070570), even though LibreOffice silently
 * tolerates it — so an imported master AND each of its layouts must take an id past
 * the max of both the master-id list and every layout-id list. Recomputed per
 * allocation so ids appended earlier in the same import are counted.
 */
export function nextMasterLayoutId(dest: DeckTarget): number {
	let max = MIN_SLIDE_MASTER_ID - 1
	const presPart = dest.presentationPart
	const root = presPart.dom.documentElement
	const masterLst = root && firstChild(root, 'p:sldMasterIdLst')
	if (masterLst) {
		for (const entry of getElements(masterLst, 'p:sldMasterId')) {
			const id = intValue(attr(entry, 'id'))
			if (id !== null && id > max) max = id
		}
	}
	// Every master's layout-id list shares the same id space; scan them all.
	const presRels = dest.opc.relationshipsFor(presPart.partName)
	for (const rel of presRels.byType(SLIDE_MASTER_REL)) {
		const masterPart = dest.opc.part(presRels.resolveTarget(rel.id))
		const masterRoot = masterPart?.dom.documentElement
		const layoutLst = masterRoot && firstChild(masterRoot, 'p:sldLayoutIdLst')
		if (!layoutLst) continue
		for (const entry of getElements(layoutLst, 'p:sldLayoutId')) {
			const id = intValue(attr(entry, 'id'))
			if (id !== null && id > max) max = id
		}
	}
	return max + 1
}

/** Empty a freshly-copied master's `p:sldLayoutIdLst`; copied layouts repopulate it. */
export function clearLayoutIdList(dest: DeckTarget, masterPartName: string): void {
	const masterPart = dest.opc.part(masterPartName)
	const root = masterPart?.dom.documentElement
	const lst = root && firstChild(root, 'p:sldLayoutIdLst')
	if (!masterPart || !lst) return
	removeChildrenByQName(lst, ['p:sldLayoutId'])
	masterPart.markDirty()
}

/**
 * Append a `p:sldLayoutId` entry for `layoutPartName` to `masterPartName`'s
 * `p:sldLayoutIdLst`, adding the master→layout relationship it points at. The
 * destination-side half of linking a copied layout into its copied master; the
 * caller resolves which master that is from the source-side rels.
 */
export function addLayoutToMaster(dest: DeckTarget, masterPartName: string, layoutPartName: string): void {
	const masterPart = dest.opc.part(masterPartName)
	const root = masterPart?.dom.documentElement
	if (!masterPart || !root) return

	const masterRels = dest.opc.relationshipsFor(masterPartName)
	const relId = masterRels.add(SLIDE_LAYOUT_REL, relativePartName(masterPartName, layoutPartName)).id
	const lst = getOrAddChild(root, 'p:sldLayoutIdLst', ['p:transition', 'p:timing', 'p:hf', 'p:txStyles', 'p:extLst'])
	const entry = createElement(masterPart.dom, 'p:sldLayoutId')
	setAttr(entry, 'id', String(nextMasterLayoutId(dest)))
	setAttr(entry, 'r:id', relId)
	lst.appendChild(entry)
	masterPart.markDirty()
}
