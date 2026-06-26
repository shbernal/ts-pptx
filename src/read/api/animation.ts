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
import { OOXML_NS, attr, childElements, firstChild, intValue, setAttr, type Element } from '../oxml/dom.js'

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
