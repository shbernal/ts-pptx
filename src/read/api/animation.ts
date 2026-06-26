/**
 * Spid-aware, otherwise-opaque handling of a slide's build-animation tree
 * (`p:timing` / `p:bldLst`).
 *
 * The `p:timing` subtree is `CT_TimeNodeList` — deeply recursive and modeled
 * here as opaque DOM (an unmodified slide round-trips it byte-identically for
 * free). The only structured data we extract or rewrite is the set of `spid`
 * references — `<p:spTgt spid>` inside the timing tree and `<p:bldP spid>` in the
 * build list — so animations stay coherent when shape ids change. Three purely
 * structural operations: **enumerate**, **remap**, and **prune**. See
 * `docs/animations-and-transitions.md` ("spid-awareness").
 */
import {
	OOXML_NS,
	attr,
	childElements,
	createElement,
	firstChild,
	insertInOrder,
	intValue,
	parseXml,
	setAttr,
	type Document,
	type Element,
} from '../oxml/dom.js'

const P_NS = OOXML_NS.p

/** Every element carrying an animation `spid` reference, across the timing tree and build list. */
function spidElements(root: Element): Element[] {
	const out: Element[] = []
	const spTgts = root.getElementsByTagNameNS(OOXML_NS.p, 'spTgt')
	for (let i = 0; i < spTgts.length; i++) out.push(spTgts[i])
	const bldPs = root.getElementsByTagNameNS(OOXML_NS.p, 'bldP')
	for (let i = 0; i < bldPs.length; i++) out.push(bldPs[i])
	return out
}

/**
 * Whether the slide carries build animations — a `<p:bldP>` entry or any timing
 * node with a `presetID` (distinguishing animation timing from the media-loop
 * timing that has neither).
 */
export function hasAnimations(root: Element): boolean {
	if (root.getElementsByTagNameNS(OOXML_NS.p, 'bldP').length > 0) return true
	const cTns = root.getElementsByTagNameNS(OOXML_NS.p, 'cTn')
	for (let i = 0; i < cTns.length; i++) {
		if (attr(cTns[i], 'presetID') !== null) return true
	}
	return false
}

/** The sorted, de-duplicated set of `spid`s referenced by the slide's animations. */
export function enumerateSpids(root: Element): number[] {
	const seen = new Set<number>()
	for (const element of spidElements(root)) {
		const spid = intValue(attr(element, 'spid'))
		if (spid !== null) seen.add(spid)
	}
	return [...seen].sort((a, b) => a - b)
}

/**
 * Rewrite every animation `spid` according to `mapping` (old → new). Returns
 * `true` when at least one reference changed, so callers can `markDirty()` only
 * when needed.
 */
export function remapSpids(root: Element, mapping: Map<number, number>): boolean {
	let changed = false
	for (const element of spidElements(root)) {
		const spid = intValue(attr(element, 'spid'))
		if (spid === null) continue
		const next = mapping.get(spid)
		if (next !== undefined && next !== spid) {
			setAttr(element, 'spid', String(next))
			changed = true
		}
	}
	return changed
}

/** The nearest ancestor `<p:par>` whose `<p:cTn>` carries a `presetID` (the effect node), or `null`. */
function effectParFor(spTgt: Element): Element | null {
	for (let node: Element | null = spTgt.parentNode as Element | null; node; node = node.parentNode as Element | null) {
		if (node.localName === 'par' && node.namespaceURI === OOXML_NS.p) {
			const cTn = firstChild(node, 'p:cTn')
			if (cTn && attr(cTn, 'presetID') !== null) return node
		}
	}
	return null
}

/** Whether a `<p:par>` wrapper node has become empty (its `p:cTn/p:childTnLst` has no element children). */
function isEmptyWrapperPar(par: Element): boolean {
	if (par.localName !== 'par' || par.namespaceURI !== OOXML_NS.p) return false
	const cTn = firstChild(par, 'p:cTn')
	if (!cTn || attr(cTn, 'presetID') !== null) return false // an effect par, not a wrapper
	const childTnLst = firstChild(cTn, 'p:childTnLst')
	return !childTnLst || childElements(childTnLst).length === 0
}

/**
 * Remove all animation traces of the given `spid`s: each shape's `<p:bldP>` and
 * the effect-level `<p:par>` nodes whose `<p:spTgt>` targets it, then any wrapper
 * `<p:par>` left empty by the removal. Keeps the opaque timing tree free of
 * dangling references when shapes are deleted. Returns `true` when anything was
 * removed.
 */
export function pruneSpids(root: Element, spids: Iterable<number>): boolean {
	const drop = new Set<number>()
	for (const s of spids) drop.add(s)
	if (drop.size === 0) return false

	let changed = false
	const wrappersToCheck = new Set<Element>()

	// Remove effect pars whose spTgt targets a pruned shape.
	const spTgts = Array.from(root.getElementsByTagNameNS(OOXML_NS.p, 'spTgt'))
	const effectPars = new Set<Element>()
	for (const spTgt of spTgts) {
		const spid = intValue(attr(spTgt, 'spid'))
		if (spid === null || !drop.has(spid)) continue
		const effectPar = effectParFor(spTgt)
		if (effectPar) effectPars.add(effectPar)
	}
	for (const effectPar of effectPars) {
		const parent = effectPar.parentNode as Element | null
		effectPar.parentNode?.removeChild(effectPar)
		changed = true
		// The wrapper par is two levels up: childTnLst → par.
		const childTnLst = parent
		const wrapperPar = childTnLst?.parentNode?.parentNode as Element | null
		if (wrapperPar) wrappersToCheck.add(wrapperPar)
	}

	// Remove now-empty wrapper pars, climbing while they keep collapsing.
	let progress = true
	while (progress) {
		progress = false
		for (const wrapper of [...wrappersToCheck]) {
			if (isEmptyWrapperPar(wrapper)) {
				const grandWrapper = wrapper.parentNode?.parentNode?.parentNode as Element | null
				wrapper.parentNode?.removeChild(wrapper)
				wrappersToCheck.delete(wrapper)
				if (grandWrapper) wrappersToCheck.add(grandWrapper)
				progress = true
				changed = true
			}
		}
	}

	// Remove the build-list entries.
	const bldPs = Array.from(root.getElementsByTagNameNS(OOXML_NS.p, 'bldP'))
	for (const bldP of bldPs) {
		const spid = intValue(attr(bldP, 'spid'))
		if (spid !== null && drop.has(spid)) {
			bldP.parentNode?.removeChild(bldP)
			changed = true
		}
	}

	// Drop an emptied <p:bldLst>.
	const bldLsts = Array.from(root.getElementsByTagNameNS(OOXML_NS.p, 'bldLst'))
	for (const bldLst of bldLsts) {
		if (childElements(bldLst).length === 0) bldLst.parentNode?.removeChild(bldLst)
	}

	return changed
}

/**
 * Flatten the whole slide to its final static state by removing the build
 * animation timeline: drop the `<p:timing>` block (which carries the `<p:bldLst>`
 * and the effect tree) so every shape renders at once with no click-through
 * staging. Gated on {@link hasAnimations}: a `<p:timing>` that is purely a
 * media loop (no `<p:bldP>` and no `presetID`-bearing node) is left untouched so
 * media playback survives. Returns `true` when a timing block was removed.
 *
 * This is the whole-slide counterpart to {@link pruneSpids} (which drops the
 * builds of specific shapes). It only removes staging; it never deletes shapes,
 * so a slide that animated alternating states over the same region will show
 * them all at once after flattening — that is "drop animations", distinct from
 * "remove staged/duplicate shapes".
 */
export function flattenAnimations(root: Element): boolean {
	if (!hasAnimations(root)) return false
	const timing = firstChild(root, 'p:timing')
	if (!timing) return false
	timing.parentNode?.removeChild(timing)
	return true
}

// --- carry: bring a copied shape's build animation into a destination slide ---

const P_XMLNS = `xmlns:p="${P_NS}"`
/** A fresh `p:timing` carrying only an empty `tmRoot` (no effects yet). */
const TIMING_SCAFFOLD = `<p:timing ${P_XMLNS}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst/></p:cTn></p:par></p:tnLst></p:timing>`
/** A fresh `mainSeq` `p:seq` (its `cTn@id` is filled in by the caller). */
const MAIN_SEQ_SCAFFOLD =
	`<p:seq ${P_XMLNS} concurrent="1" nextAc="seek"><p:cTn id="0" dur="indefinite" nodeType="mainSeq"><p:childTnLst/></p:cTn>` +
	'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
	'<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'

/** Largest `<p:cTn @id>` in a subtree (0 when none). */
function maxCTnId(scope: Element): number {
	let max = 0
	const cTns = scope.getElementsByTagNameNS(P_NS, 'cTn')
	for (let i = 0; i < cTns.length; i++) {
		const id = intValue(attr(cTns[i], 'id'))
		if (id !== null && id > max) max = id
	}
	return max
}

/** Renumber every `<p:cTn @id>` in `node` (document order) starting at `start`; returns the next free id. */
function renumberCTnIds(node: Element, start: number): number {
	let id = start
	const cTns = node.getElementsByTagNameNS(P_NS, 'cTn')
	for (let i = 0; i < cTns.length; i++) setAttr(cTns[i], 'id', String(id++))
	return id
}

/** Whether `par`'s subtree targets any of `spids` via a `<p:spTgt>`. */
function targetsAnySpid(par: Element, spids: Set<number>): boolean {
	const spTgts = par.getElementsByTagNameNS(P_NS, 'spTgt')
	for (let i = 0; i < spTgts.length; i++) {
		const spid = intValue(attr(spTgts[i], 'spid'))
		if (spid !== null && spids.has(spid)) return true
	}
	return false
}

/** Parse a namespaced scaffold string and import its root into `doc`. */
function importScaffold(doc: Document, xml: string): Element {
	return doc.importNode(parseXml(xml).documentElement as Element, true) as Element
}

/** The destination `mainSeq` `p:childTnLst` to append click groups into, creating the timing/seq scaffold as needed. */
function getOrCreateMainSeqChildTnLst(root: Element, doc: Document): Element {
	let timing = firstChild(root, 'p:timing')
	if (!timing) {
		timing = importScaffold(doc, TIMING_SCAFFOLD)
		insertInOrder(root, timing, ['p:extLst'])
	}
	// Reuse an existing mainSeq if present.
	const existingSeq = timing.getElementsByTagNameNS(P_NS, 'seq')[0] as Element | undefined
	if (existingSeq) {
		const seqCTn = firstChild(existingSeq, 'p:cTn')
		if (seqCTn) {
			let childLst = firstChild(seqCTn, 'p:childTnLst')
			if (!childLst) {
				childLst = createElement(doc, 'p:childTnLst')
				seqCTn.appendChild(childLst)
			}
			return childLst
		}
	}
	// Otherwise insert a fresh mainSeq into the tmRoot child list (before any media nodes).
	let tmRootChildLst = timing.getElementsByTagNameNS(P_NS, 'childTnLst')[0] as Element | undefined
	if (!tmRootChildLst) {
		// Degenerate timing with no tmRoot child list — graft the full scaffold's tnLst.
		const fresh = importScaffold(doc, TIMING_SCAFFOLD)
		const freshTnLst = firstChild(fresh, 'p:tnLst')
		if (freshTnLst) insertInOrder(timing, freshTnLst, ['p:bldLst', 'p:extLst'])
		tmRootChildLst = timing.getElementsByTagNameNS(P_NS, 'childTnLst')[0] as Element
	}
	const seq = importScaffold(doc, MAIN_SEQ_SCAFFOLD)
	setAttr(firstChild(seq, 'p:cTn') as Element, 'id', String(maxCTnId(timing) + 1))
	tmRootChildLst.insertBefore(seq, tmRootChildLst.firstChild)
	return firstChild(firstChild(seq, 'p:cTn') as Element, 'p:childTnLst') as Element
}

/** The destination `p:bldLst` (under `p:timing`), creating it before any `p:extLst`. */
function getOrCreateBldLst(timing: Element, doc: Document): Element {
	let bldLst = firstChild(timing, 'p:bldLst')
	if (!bldLst) {
		bldLst = createElement(doc, 'p:bldLst')
		insertInOrder(timing, bldLst, ['p:extLst'])
	}
	return bldLst
}

/**
 * Carry the build animation of one or more copied shapes from `sourceRoot`'s
 * timing into `targetRoot`'s, remapping shape ids via `spidMap` (source id → new
 * id on the destination). For each mapped shape: its mainSeq click-group `<p:par>`
 * (the whole click step) and `<p:bldP>` are cloned, their `spid`s remapped, their
 * `<p:cTn>` ids renumbered to stay collision-free in the destination, and appended
 * after any existing build — matching how PowerPoint merges a pasted shape's
 * animation. Returns `true` when anything was carried (so the caller can
 * `markDirty()`), `false` when the source shapes have no animation.
 */
export function carryShapeAnimations(sourceRoot: Element, targetRoot: Element, spidMap: Map<number, number>): boolean {
	const sourceTiming = firstChild(sourceRoot, 'p:timing')
	if (!sourceTiming) return false
	const carriedSpids = new Set(spidMap.keys())

	// Source click groups (direct children of the mainSeq child list) that target a carried shape.
	const sourceSeq = sourceTiming.getElementsByTagNameNS(P_NS, 'seq')[0] as Element | undefined
	const groups: Element[] = []
	if (sourceSeq) {
		const seqCTn = firstChild(sourceSeq, 'p:cTn')
		const seqChildLst = seqCTn ? firstChild(seqCTn, 'p:childTnLst') : null
		if (seqChildLst) {
			for (const par of childElements(seqChildLst)) {
				if (targetsAnySpid(par, carriedSpids)) groups.push(par)
			}
		}
	}
	// Source build-list entries for the carried shapes.
	const sourceBldLst = firstChild(sourceTiming, 'p:bldLst')
	const bldPs: Element[] = []
	if (sourceBldLst) {
		for (const bldP of childElements(sourceBldLst)) {
			const spid = intValue(attr(bldP, 'spid'))
			if (spid !== null && carriedSpids.has(spid)) bldPs.push(bldP)
		}
	}
	if (groups.length === 0 && bldPs.length === 0) return false

	const doc = targetRoot.ownerDocument as Document
	const destChildLst = getOrCreateMainSeqChildTnLst(targetRoot, doc)
	const destTiming = firstChild(targetRoot, 'p:timing') as Element
	let nextId = maxCTnId(destTiming) + 1

	for (const par of groups) {
		const clone = doc.importNode(par, true) as Element
		remapSpids(clone, spidMap)
		nextId = renumberCTnIds(clone, nextId)
		destChildLst.appendChild(clone)
	}
	if (bldPs.length > 0) {
		const destBldLst = getOrCreateBldLst(destTiming, doc)
		for (const bldP of bldPs) {
			const clone = doc.importNode(bldP, true) as Element
			const spid = intValue(attr(clone, 'spid'))
			if (spid !== null && spidMap.has(spid)) setAttr(clone, 'spid', String(spidMap.get(spid)))
			destBldLst.appendChild(clone)
		}
	}
	return true
}
