// Read-model coverage for the spid-aware animation surgery in
// src/read/api/animation.ts — hasAnimations / enumerateSpids / remapSpids /
// pruneSpids / flattenAnimations. These are structural DOM operations over the
// opaque `p:timing` tree; the real fixture decks only carry one well-formed
// timing tree apiece, so the tier-selection, empty-wrapper collapse, null-spid,
// and no-op branches never all fire. Here a synthetic Slide (a hand-authored
// `p:sld` fed through `new Part`, wrapped in `new Slide(null, part, id, idx)` —
// the animation methods touch only `this.part.dom`, never the package) drives
// each branch directly, including malformed / degenerate timing trees.

import { describe, test } from 'vitest'
import { Part, Slide } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

/** A synthetic read-model Slide over a hand-authored `p:sld` body (shape tree is empty). */
function slide(bodyXml) {
	const xml = `<p:sld xmlns:p="${P_NS}"><p:cSld><p:spTree/></p:cSld>${bodyXml}</p:sld>`
	const part = new Part('/ppt/slides/slide1.xml', SLIDE_CT, new TextEncoder().encode(xml))
	return new Slide(/** @type {any} */ (null), part, 1, 0)
}

/** An effect `<p:par>` (its `<p:cTn>` carries a presetID) targeting `spid`. */
const effectPar = (spid, id) =>
	`<p:par><p:cTn id="${id}" presetID="1"><p:childTnLst>` +
	`<p:set><p:cBhvr><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:set>` +
	`</p:childTnLst></p:cTn></p:par>`

/** A wrapper `<p:par>` (no presetID) around `inner`. */
const wrapPar = (id, inner) => `<p:par><p:cTn id="${id}"><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`

/** A `p:timing` whose mainSeq child list holds `clickGroups`, with an optional `p:bldLst`. */
function timing(clickGroups, bldLst = '') {
	return (
		`<p:timing><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst>` +
		`<p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst>${clickGroups}</p:childTnLst></p:cTn></p:seq>` +
		`</p:childTnLst></p:cTn></p:par></p:tnLst>${bldLst}</p:timing>`
	)
}

describe('hasAnimations — presetID vs bldP vs media-only', () => {
	test('a slide with a <p:bldP> reports animations', () => {
		const s = slide(timing('', '<p:bldLst><p:bldP spid="5"/></p:bldLst>'))
		assert(s.hasAnimations, 'a bldP entry means animations')
	})

	test('a presetID-bearing time node with no bldP still reports animations', () => {
		const s = slide(timing(wrapPar(3, effectPar(5, 4))))
		assert(s.hasAnimations, 'a presetID cTn is an animation even without a build list')
	})

	test('a media-only timing (no bldP, no presetID) reports no animations', () => {
		// A bare timing scaffold — the media-loop shape has neither a build nor a preset.
		const s = slide(timing(''))
		assert(!s.hasAnimations, 'media-loop timing is not a build animation')
	})
})

describe('enumerateSpids — dedup, sort, null-spid skip', () => {
	test('spids from spTgt and bldP are merged, sorted, de-duplicated; a non-numeric spid is skipped', () => {
		const s = slide(
			timing(
				wrapPar(3, effectPar(9, 4)) + wrapPar(5, effectPar(2, 6)),
				// bldP repeats 9 (dedup) and carries a non-numeric spid (skipped).
				'<p:bldLst><p:bldP spid="9"/><p:bldP spid="2"/><p:bldP spid="notanid"/></p:bldLst>'
			)
		)
		assertEqual(JSON.stringify(s.animationSpids()), '[2,9]', 'sorted + de-duplicated, junk spid dropped')
	})

	test('a slide with no timing enumerates to an empty list', () => {
		assertEqual(JSON.stringify(slide('').animationSpids()), '[]', 'no animations → no spids')
	})
})

describe('remapSpids — change / identity / absent / null spid', () => {
	test('only genuinely-changed references mark the slide dirty', () => {
		const s = slide(
			timing(
				wrapPar(3, effectPar(9, 4)) + wrapPar(5, effectPar(2, 6)),
				'<p:bldLst><p:bldP spid="9"/><p:bldP spid="notanid"/></p:bldLst>'
			)
		)
		let dirtied = 0
		s.part.markDirty = () => dirtied++
		// 9 → 20 (change), 2 → 2 (identity, no change), 7 → 30 (absent from tree). null spid skipped.
		s.remapAnimationSpids(
			new Map([
				[9, 20],
				[2, 2],
				[7, 30],
			])
		)
		assertEqual(JSON.stringify(s.animationSpids()), '[2,20]', '9 remapped, 2 unchanged')
		assertEqual(dirtied, 1, 'a real change dirties exactly once')
	})

	test('a mapping that changes nothing leaves the slide clean', () => {
		const s = slide(timing(wrapPar(3, effectPar(9, 4))))
		let dirtied = 0
		s.part.markDirty = () => dirtied++
		s.remapAnimationSpids(
			new Map([
				[9, 9],
				[100, 200],
			])
		) // identity + absent only
		assertEqual(dirtied, 0, 'no reference changed → no markDirty')
	})
})

describe('pruneSpids — effect removal, wrapper collapse, bldLst drop', () => {
	test('an empty spid set is a no-op', () => {
		const s = slide(timing(wrapPar(3, effectPar(9, 4)), '<p:bldLst><p:bldP spid="9"/></p:bldLst>'))
		let dirtied = 0
		s.part.markDirty = () => dirtied++
		s.pruneAnimationSpids([])
		assertEqual(dirtied, 0, 'nothing to drop → no markDirty')
		assertEqual(JSON.stringify(s.animationSpids()), '[9]', 'the animation is untouched')
	})

	test('pruning collapses nested empty wrappers up to the mainSeq and drops the emptied bldLst', () => {
		// wrapperB > wrapperA > effect(spid 7); a sibling group keeps spid 9.
		const s = slide(
			timing(
				wrapPar(3, wrapPar(4, effectPar(7, 5))) + wrapPar(6, effectPar(9, 7)),
				'<p:bldLst><p:bldP spid="7"/></p:bldLst>'
			)
		)
		s.pruneAnimationSpids([7])
		assertEqual(JSON.stringify(s.animationSpids()), '[9]', 'spid 7 fully removed, spid 9 retained')
		// The bldLst held only spid 7 → it is now empty and dropped, so hasAnimations rides on the
		// surviving presetID effect for spid 9.
		assert(s.hasAnimations, 'the surviving effect keeps animations reported')
	})

	test('a spTgt outside any effect par (no presetID ancestor) is left in place', () => {
		// The spid-5 spTgt sits directly under a wrapper — effectParFor returns null, so prune
		// finds no effect par to remove and only the bldP (if any) would go. Here there is none.
		const s = slide(timing(wrapPar(3, '<p:set><p:cBhvr><p:tgtEl><p:spTgt spid="5"/></p:tgtEl></p:cBhvr></p:set>')))
		let dirtied = 0
		s.part.markDirty = () => dirtied++
		s.pruneAnimationSpids([5])
		assertEqual(dirtied, 0, 'no effect par and no bldP → nothing removed')
		assertEqual(JSON.stringify(s.animationSpids()), '[5]', 'the bare spTgt reference survives')
	})

	test('an effect par nested directly inside another effect par is not treated as an empty wrapper', () => {
		// The outer par also carries a presetID: after the inner effect is removed, isEmptyWrapperPar
		// sees the presetID and bails (it is an effect, not a wrapper) rather than collapsing it.
		const inner = effectPar(7, 5)
		const outerEffect = `<p:par><p:cTn id="4" presetID="1"><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`
		const s = slide(timing(wrapPar(3, outerEffect), '<p:bldLst><p:bldP spid="7"/></p:bldLst>'))
		s.pruneAnimationSpids([7])
		assertEqual(JSON.stringify(s.animationSpids()), '[]', 'the targeted effect is gone')
		// The outer presetID par is preserved (not collapsed as a wrapper).
		assert(!s.hasAnimations || s.animationSpids().length === 0, 'no dangling spid remains')
	})
})

describe('flattenAnimations — timing removal and its gates', () => {
	test('a build animation is flattened by removing the timing block', () => {
		const s = slide(timing(wrapPar(3, effectPar(9, 4)), '<p:bldLst><p:bldP spid="9"/></p:bldLst>'))
		let dirtied = 0
		s.part.markDirty = () => dirtied++
		assert(s.flattenAnimations(), 'flatten returns true when a timing block is removed')
		assert(!s.hasAnimations, 'no animations remain after flattening')
		assertEqual(dirtied, 1, 'flattening dirties the slide')
	})

	test('a media-only timing is preserved (flatten is a no-op)', () => {
		const s = slide(timing(''))
		assert(!s.flattenAnimations(), 'flatten does not touch media-loop timing')
	})

	test('animations reported outside a <p:timing> block cannot be flattened', () => {
		// Malformed: a bldLst directly under the slide, no p:timing wrapper. hasAnimations is true
		// (the bldP is found anywhere), but there is no timing block to remove.
		const s = slide('<p:bldLst><p:bldP spid="5"/></p:bldLst>')
		assert(s.hasAnimations, 'the loose bldP still reads as an animation')
		assert(!s.flattenAnimations(), 'no timing block → nothing to flatten')
	})
})
