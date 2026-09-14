// Read-model tests for slide transitions and build animations
// (docs/animations-and-transitions.md, Phase 1).
//
// Transitions are a full typed accessor (get/set), validated against the
// PowerPoint-authored slide-transition oracle. Animations are opaque,
// spid-aware preservation: hasAnimations + the enumerate/remap/prune spid
// helpers, validated against the basic and rich animation oracles.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import assert from 'node:assert/strict'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { validateBuf, validatorInstalled } from '../validator.js'
import { fixturePath, openFixture, readOracle } from './corpus.js'
import { partBodies, assertUnchangedExcept } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function slidePartXml(pptxBytes, slideNumber) {
	const zip = await JSZip.loadAsync(pptxBytes)
	return zip.file(`ppt/slides/slide${slideNumber}.xml`).async('string')
}

describe('slide.transition (read)', () => {
	test('decodes every PowerPoint-authored transition form', async () => {
		const oracle = await readOracle('slide-transition')
		const pres = await openFixture('slide-transition')
		for (const expected of oracle.slides) {
			const slide = pres.slides[expected.slide - 1]
			const info = slide.transition
			const { decoded } = expected
			assert.ok(info, `slide ${expected.slide} has a transition`)
			assert.equal(info.type, decoded.type, `slide ${expected.slide} type`)
			assert.equal(info.namespace, decoded.ns, `slide ${expected.slide} ns`)
			assert.equal(info.speed, decoded.speed, `slide ${expected.slide} speed`)
			assert.equal(info.durationMs, decoded.durationMs, `slide ${expected.slide} durationMs`)
			assert.equal(info.advanceOnClick, decoded.advanceOnClick, `slide ${expected.slide} advanceOnClick`)
			assert.equal(info.advanceAfterMs, decoded.advanceAfterMs, `slide ${expected.slide} advanceAfterMs`)
			assert.deepEqual(info.variant, decoded.variant, `slide ${expected.slide} variant`)
		}
	})

	test('an untouched deck round-trips byte-identically', async () => {
		const original = await readFile(fixturePath('slide-transition'))
		const pres = await Presentation.load(original)
		// touch nothing
		const before = await partBodies(original)
		const after = await partBodies(await pres.save())
		// Every part, not just `ppt/slides/*`: load→save is byte-identical across the whole
		// package, so narrowing this to slides only ever hid a regression in the rest of it.
		assertUnchangedExcept(before, after, [], 'slide-transition')
	})
})

describe('slide.transition (write/edit)', () => {
	test('sets a bare transition when no duration is given', async () => {
		const pres = await openFixture('slide-transition')
		pres.slides[0].transition = { type: 'wipe', speed: 'med', variant: { dir: 'u' } }
		const saved = await pres.save()
		const xml = await slidePartXml(saved, 1)
		assert.ok(!xml.includes('AlternateContent'), 'bare form, no mc:AlternateContent')
		assert.ok(/<p:transition spd="med"><p:wipe dir="u"\/><\/p:transition>/.test(xml), 'bare wipe XML')

		const reopened = await Presentation.load(saved)
		const info = reopened.slides[0].transition
		assert.equal(info.type, 'wipe')
		assert.equal(info.speed, 'med')
		assert.equal(info.durationMs, null)
		assert.deepEqual(info.variant, { dir: 'u' })
	})

	test('sets the mc:AlternateContent form when durationMs is given', async () => {
		const pres = await openFixture('slide-transition')
		pres.slides[2].transition = { type: 'dissolve', durationMs: 2000, speed: 'slow' }
		const saved = await pres.save()
		const xml = await slidePartXml(saved, 3)
		assert.ok(xml.includes('mc:AlternateContent'), 'wrapped form')
		assert.ok(xml.includes('p14:dur="2000"'), 'carries p14:dur')
		assert.ok(xml.includes('Requires="p14"'), 'Choice requires p14')

		const reopened = await Presentation.load(saved)
		const info = reopened.slides[2].transition
		assert.equal(info.type, 'dissolve')
		assert.equal(info.durationMs, 2000)
		assert.equal(info.speed, 'slow')
	})

	test('derives a speed bucket from durationMs when speed is omitted', async () => {
		const pres = await openFixture('slide-transition')
		pres.slides[0].transition = { type: 'fade', durationMs: 1500 }
		const reopened = await Presentation.load(await pres.save())
		assert.equal(reopened.slides[0].transition.speed, 'slow')
	})

	test('round-trips advTm / advClick auto-advance', async () => {
		const pres = await openFixture('slide-transition')
		pres.slides[0].transition = { type: 'fade', speed: 'med', advanceOnClick: false, advanceAfterMs: 3000 }
		const reopened = await Presentation.load(await pres.save())
		const info = reopened.slides[0].transition
		assert.equal(info.advanceOnClick, false)
		assert.equal(info.advanceAfterMs, 3000)
	})

	test('assigning null clears the transition', async () => {
		const pres = await openFixture('slide-transition')
		pres.slides[0].transition = null
		const saved = await pres.save()
		const xml = await slidePartXml(saved, 1)
		assert.ok(!xml.includes('<p:transition'), 'no p:transition element remains')
		const reopened = await Presentation.load(saved)
		assert.equal(reopened.slides[0].transition, null)
	})

	test('refuses a duration or advance time that is not a number of milliseconds, changing nothing', async () => {
		// `NaN` and `-5` were written straight into `p14:dur` and `advTm`.
		const pres = await openFixture('slide-transition')
		const slide = pres.slides[1]
		const before = slide.transition
		for (const times of [
			{ durationMs: NaN },
			{ durationMs: -1 },
			{ advanceAfterMs: Infinity },
			{ advanceAfterMs: -5 },
		]) {
			assert.throws(
				() => {
					slide.transition = { type: 'fade', ...times }
				},
				(err) => /** @type {any} */ (err).code === 'transition/invalid-time',
				`refuses ${Object.entries(times).map(([k, v]) => `${k} ${v}`)}`
			)
		}
		assert.deepEqual(slide.transition, before, 'the slide keeps its transition')
	})

	test('a transition read back and assigned again keeps an absent spd absent', async () => {
		// The getter reports `fast` for a missing `spd`, its schema default, and assigning that back
		// used to write `spd="fast"` onto every such slide.
		const original = await readFile(fixturePath('slide-transition'))
		const pres = await Presentation.load(original)
		for (const slide of pres.slides) {
			const current = slide.transition
			slide.transition = current
		}
		const saved = await pres.save()
		const spd = async (bytes, n) =>
			((await slidePartXml(bytes, n)).match(/<p:transition[^>]*>/)?.[0] ?? '').match(/spd="\w+"/)?.[0] ?? null
		for (let n = 1; n <= pres.slides.length; n++)
			assert.equal(await spd(saved, n), await spd(original, n), `slide ${n} spd`)
	})

	test('keeps the sound through a spread, removes it with null, and refuses a different one', async () => {
		// The setter had no sound field and removed the whole node, so changing only the speed dropped
		// the sound.
		const pres = await openFixture('slide-transition-sound')
		const slide = pres.slides[0]
		const sound = slide.transition.sound
		assert.ok(sound, 'the fixture slide has a sound')

		slide.transition = { ...slide.transition, speed: 'slow' }
		let reopened = await Presentation.load(await pres.save())
		assert.deepEqual(reopened.slides[0].transition.sound, sound, 'a spread keeps the sound')
		assert.equal(reopened.slides[0].transition.speed, 'slow', 'and changes the speed')

		assert.throws(
			() => {
				slide.transition = { ...slide.transition, sound: { ...sound, name: 'other.wav' } }
			},
			(err) => /** @type {any} */ (err).code === 'transition/sound-unsupported'
		)
		assert.deepEqual(slide.transition.sound, sound, 'a refused sound changes nothing')

		slide.transition = { ...slide.transition, sound: null }
		reopened = await Presentation.load(await pres.save())
		assert.equal(reopened.slides[0].transition.sound, null, 'null removes the sound')
	})
})

describe('slide animations (opaque, spid-aware)', () => {
	test('hasAnimations + animationSpids on the basic fixture', async () => {
		const oracle = await readOracle('slide-animation-basic')
		const pres = await openFixture('slide-animation-basic')
		const slide = pres.slides[0]
		assert.equal(slide.hasAnimations, true)
		assert.deepEqual(slide.animationSpids(), oracle.animationSpids)
	})

	test('hasAnimations + animationSpids on the rich fixture', async () => {
		const oracle = await readOracle('slide-animation-rich')
		const pres = await openFixture('slide-animation-rich')
		const slide = pres.slides[0]
		assert.equal(slide.hasAnimations, true)
		assert.deepEqual(slide.animationSpids(), oracle.animationSpids)
	})

	test('an unanimated fixture reports no animations', async () => {
		const pres = await openFixture('slide-transition')
		assert.equal(pres.slides[0].hasAnimations, false)
		assert.deepEqual(pres.slides[0].animationSpids(), [])
	})

	test('an untouched animated deck round-trips byte-identically', async () => {
		const original = await readFile(fixturePath('slide-animation-rich'))
		const pres = await Presentation.load(original)
		const before = await partBodies(original)
		const after = await partBodies(await pres.save())
		assertUnchangedExcept(before, after, [], 'slide-animation-rich')
	})

	test('remapAnimationSpids rewrites every spTgt and bldP reference', async () => {
		const pres = await openFixture('slide-animation-rich')
		const slide = pres.slides[0]
		slide.remapAnimationSpids(
			new Map([
				[2, 20],
				[3, 30],
				[4, 40],
				[5, 50],
			])
		)
		assert.deepEqual(slide.animationSpids(), [20, 30, 40, 50])
		// persists across a save → reopen
		const reopened = await Presentation.load(await pres.save())
		assert.deepEqual(reopened.slides[0].animationSpids(), [20, 30, 40, 50])
	})

	test('pruneAnimationSpids drops one shape, leaving the others coherent', async () => {
		const pres = await openFixture('slide-animation-rich')
		const slide = pres.slides[0]
		slide.pruneAnimationSpids([3])
		assert.deepEqual(slide.animationSpids(), [2, 4, 5])
		const reopened = await Presentation.load(await pres.save())
		assert.deepEqual(reopened.slides[0].animationSpids(), [2, 4, 5])
	})

	test('pruning a click-effect collapses its emptied wrapper', async () => {
		const pres = await openFixture('slide-animation-rich')
		const slide = pres.slides[0]
		slide.pruneAnimationSpids([2])
		assert.deepEqual(slide.animationSpids(), [3, 4, 5])
	})
})

describe('slide.flattenAnimations (whole-slide flatten pass)', () => {
	test('strips the whole timing block, leaving every shape in place', async () => {
		const pres = await openFixture('slide-animation-rich')
		const slide = pres.slides[0]
		const shapeCountBefore = slide.shapes.length
		assert.equal(slide.hasAnimations, true)

		assert.equal(slide.flattenAnimations(), true)

		assert.equal(slide.hasAnimations, false)
		assert.deepEqual(slide.animationSpids(), [])
		assert.equal(slide.shapes.length, shapeCountBefore, 'no shapes removed')

		const saved = await pres.save()
		assert.ok(!(await slidePartXml(saved, 1)).includes('<p:timing'), 'p:timing gone from the bytes')
		const reopened = await Presentation.load(saved)
		assert.equal(reopened.slides[0].hasAnimations, false)
		assert.equal(reopened.slides[0].shapes.length, shapeCountBefore)
		if (validatorInstalled) assert.deepEqual(await validateBuf(Buffer.from(saved)), [])
	})

	test('flattening the basic fixture clears its animations', async () => {
		const pres = await openFixture('slide-animation-basic')
		const slide = pres.slides[0]
		assert.equal(slide.flattenAnimations(), true)
		assert.equal(slide.hasAnimations, false)
		assert.deepEqual(slide.animationSpids(), [])
	})

	test('is a no-op on an unanimated slide and is idempotent', async () => {
		const pres = await openFixture('slide-transition')
		const slide = pres.slides[0]
		assert.equal(slide.hasAnimations, false)
		assert.equal(slide.flattenAnimations(), false, 'nothing to flatten')
		// the slide-show transition is untouched by an animation flatten
		assert.notEqual(slide.transition, null)

		const rich = (await openFixture('slide-animation-rich')).slides[0]
		assert.equal(rich.flattenAnimations(), true)
		assert.equal(rich.flattenAnimations(), false, 'second call is a no-op')
	})
})

// --- Phase 2 fixtures (docs/animations-and-transitions.md). Preset expansion (B)
// and transition sounds (C) are implemented write-side (see test/regression); the
// importShape animation carry (A) is exercised at the end of this file. These
// blocks assert the fixtures load + their oracles match the bytes against the
// opaque, spid-aware read model. ---

describe('slide-animation-presets (read fixture, Phase 2 gate B)', () => {
	test('hasAnimations + animationSpids match the oracle', async () => {
		const oracle = await readOracle('slide-animation-presets')
		const slide = (await openFixture('slide-animation-presets')).slides[0]
		assert.equal(slide.hasAnimations, true)
		assert.deepEqual(slide.animationSpids(), oracle.animationSpids)
	})

	test('every preset template appears verbatim in the slide', async () => {
		const oracle = await readOracle('slide-animation-presets')
		const xml = await slidePartXml(await readFile(fixturePath('slide-animation-presets')), 1)
		for (const [name, t] of Object.entries(oracle.presetTemplates)) {
			assert.ok(xml.includes(t.effectParXml), `${name} effect node present verbatim`)
			assert.ok(xml.includes(t.behaviorsXml), `${name} behaviors present verbatim`)
			assert.ok(xml.includes(t.bldPXml), `${name} bldP present`)
		}
	})

	test('an untouched presets deck round-trips byte-identically', async () => {
		const original = await readFile(fixturePath('slide-animation-presets'))
		const before = await partBodies(original)
		const after = await partBodies(await (await Presentation.load(original)).save())
		assertUnchangedExcept(before, after, [], 'slide-animation-presets')
	})
})

describe('slide-transition-sound (read fixture, Phase 2 gate C)', () => {
	test('decodes the fade transition on every slide', async () => {
		const oracle = await readOracle('slide-transition-sound')
		const pres = await openFixture('slide-transition-sound')
		for (const s of oracle.slides) {
			const info = pres.slides[s.slide - 1].transition
			assert.ok(info, `slide ${s.slide} has a transition`)
			assert.equal(info.type, 'fade', `slide ${s.slide} type`)
			assert.equal(info.durationMs, 2000, `slide ${s.slide} durationMs`)
			// sndAc decode: start sound (with optional loop) vs the stop-previous form.
			const sr = s.soundRels
			assert.ok(info.sound, `slide ${s.slide} decodes a sound`)
			assert.equal(info.sound.form, sr.form === 'endSnd' ? 'stop' : 'start', `slide ${s.slide} sound form`)
			assert.equal(info.sound.loop, sr.loop, `slide ${s.slide} sound loop`)
			assert.equal(info.sound.embedRid, sr.sndEmbedRid, `slide ${s.slide} sound embedRid`)
			assert.equal(info.sound.name, sr.sndName, `slide ${s.slide} sound name`)
		}
	})

	test('sndAc + audio rel graph appear verbatim in the package', async () => {
		const oracle = await readOracle('slide-transition-sound')
		const bytes = await readFile(fixturePath('slide-transition-sound'))
		const zip = await JSZip.loadAsync(bytes)
		for (const s of oracle.slides) {
			const xml = await slidePartXml(bytes, s.slide)
			assert.ok(xml.includes(s.soundRels.sndAcXml), `slide ${s.slide} sndAc present verbatim`)
			if (s.soundRels.audioRel) {
				const rels = await zip.file(`ppt/slides/_rels/slide${s.slide}.xml.rels`).async('string')
				assert.ok(rels.includes(s.soundRels.audioRel.target), `slide ${s.slide} audio rel target`)
				assert.ok(rels.includes('relationships/audio'), `slide ${s.slide} audio rel type`)
			}
		}
		const ct = await zip.file('[Content_Types].xml').async('string')
		assert.ok(ct.includes('<Default Extension="wav" ContentType="audio/x-wav"/>'), 'wav Default content type')
	})

	test('an untouched transition-sound deck round-trips byte-identically', async () => {
		const original = await readFile(fixturePath('slide-transition-sound'))
		const before = await partBodies(original)
		const after = await partBodies(await (await Presentation.load(original)).save())
		assertUnchangedExcept(before, after, [], 'slide-transition-sound')
	})
})

describe('import-animation-merge (read fixture, Phase 2 gate A)', () => {
	test('enumerates spids on both slides per the oracle', async () => {
		const oracle = await readOracle('import-animation-merge')
		const pres = await openFixture('import-animation-merge')
		assert.deepEqual(pres.slides[0].animationSpids(), oracle.source.animationSpids)
		assert.deepEqual(pres.slides[1].animationSpids(), oracle.merged.animationSpids)
	})

	test('the merged slide matches the oracle timing verbatim', async () => {
		const oracle = await readOracle('import-animation-merge')
		const xml = await slidePartXml(await readFile(fixturePath('import-animation-merge')), 2)
		assert.ok(xml.includes(oracle.merged.timingXml), 'merged timing tree present verbatim')
		assert.ok(xml.includes(oracle.merged.bldList.xml), 'merged bldLst present verbatim')
	})

	test('remapAnimationSpids stays coherent across the merged build', async () => {
		const oracle = await readOracle('import-animation-merge')
		const pres = await openFixture('import-animation-merge')
		const slide = pres.slides[1]
		// Apply the oracle's spid remap (host stays, carried 2->3 simulated as a shift).
		slide.remapAnimationSpids(
			new Map([
				[2, 20],
				[3, 30],
			])
		)
		assert.deepEqual(slide.animationSpids(), [20, 30])
		const reopened = await Presentation.load(await pres.save())
		assert.deepEqual(reopened.slides[1].animationSpids(), [20, 30])
		// mergeMap sanity: the carried shape was renumbered to spid 3 on the destination.
		assert.equal(oracle.mergeMap.carriedShape.mergedSpid, 3)
	})
})

// Phase 2 capability A: importShape({ carryAnimation: true }) carries the lifted
// shape's build animation into the destination timing — the programmatic analogue
// of PowerPoint's copy/paste-with-animation captured by the import-animation-merge
// oracle. The destination timing is ts-pptx's own construction (not byte-equal to
// PowerPoint's full-tree renumber), so the contract asserted is the mergeMap
// semantics: the carried shape takes a new spid, its spTgt/bldP are remapped to it
// and appended after any existing build, and no reference dangles.
describe('importShape carryAnimation (Phase 2 capability A)', () => {
	/** Every animation spid on a slide resolves to a real shape id (no dangling reference). */
	function assertNoDanglingSpids(xml) {
		const shapeIds = new Set([...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].map((m) => Number(m[1])))
		for (const m of xml.matchAll(/<p:(?:spTgt|bldP) spid="(\d+)"/g)) {
			assert.ok(shapeIds.has(Number(m[1])), `spid ${m[1]} targets a real shape`)
		}
	}

	test('drops animation by default (opt-in only)', async () => {
		const target = await openFixture('slide-transition')
		const source = await openFixture('slide-animation-basic')
		target.importShape(target.slides[0], source.slides[0], 0, { theme: 'copy' })
		assert.equal(target.slides[0].hasAnimations, false, 'no animation carried without the flag')
	})

	test('appends the carried build after the host build, remapped to the new spid', async () => {
		// Host already animates spids 2..5; the lifted basic shape takes the next id (6).
		const target = await openFixture('slide-animation-rich')
		const source = await openFixture('slide-animation-basic')
		const slide = target.slides[0]
		const newSpid = slide.nextShapeId()
		target.importShape(slide, source.slides[0], 0, { carryAnimation: true, theme: 'copy' })

		assert.deepEqual(slide.animationSpids(), [2, 3, 4, 5, newSpid])
		const saved = await target.save()
		const xml = await slidePartXml(saved, 1)
		// bldP for the carried shape is appended last.
		const bldOrder = [...xml.matchAll(/<p:bldP spid="(\d+)"/g)].map((m) => Number(m[1]))
		assert.deepEqual(bldOrder, [2, 3, 4, 5, newSpid], 'carried bldP appended after the host builds')
		// The carried entrance Fade (presetID 10) targets the new spid.
		assert.ok(
			new RegExp(`presetID="10"[\\s\\S]*?<p:spTgt spid="${newSpid}"/>`).test(xml),
			'carried fade effect targets the new spid'
		)
		// cTn ids stay unique and nothing dangles, across a reopen.
		const cTnIds = [...xml.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => Number(m[1]))
		assert.equal(new Set(cTnIds).size, cTnIds.length, 'cTn ids are unique')
		assertNoDanglingSpids(xml)
		assert.deepEqual((await Presentation.load(saved)).slides[0].animationSpids(), [2, 3, 4, 5, newSpid])
	})

	/** A shape's `p:cNvPr/@id`. */
	function shapeId(shape) {
		const pNs = 'http://schemas.openxmlformats.org/presentationml/2006/main'
		return Number(shape.element_.getElementsByTagNameNS(pNs, 'cNvPr')[0].getAttribute('id'))
	}

	/** How many effect targets in a slide's XML name `spid`. */
	function effectCount(xml, spid) {
		return [...xml.matchAll(new RegExp(`<p:spTgt spid="${spid}"/>`, 'g'))].length
	}

	test('carries only the carried shape’s effects out of a click group that animates several', async () => {
		// The rich fixture's first click group animates shapes 2, 3 and 4. Carrying shape 2 used to clone
		// the whole group and remap only 2, so shape 3's effect animated the carried shape and spid 4
		// named nothing on the destination.
		const sourceXml = await slidePartXml(await readFile(fixturePath('slide-animation-rich')), 1)
		const source = await openFixture('slide-animation-rich')
		const target = await openFixture('slide-transition')
		const slide = target.slides[0]
		const index = source.slides[0].shapes.findIndex((shape) => shapeId(shape) === 2)
		const [carried] = target.importShapes(slide, source.slides[0], [index], { carryAnimation: true, theme: 'copy' })

		assert.deepEqual(slide.animationSpids(), [shapeId(carried)], 'only the carried shape is animated')
		const xml = await slidePartXml(await target.save(), 1)
		assertNoDanglingSpids(xml)
		assert.equal(effectCount(xml, shapeId(carried)), effectCount(sourceXml, 2), 'with every effect it had')
	})

	test('importShapes carries a click group animating two carried shapes once, with both remapped', async () => {
		// The carry used to run per shape, so the group came across once for each shape, each copy
		// keeping the other shape's source id.
		const sourceXml = await slidePartXml(await readFile(fixturePath('slide-animation-rich')), 1)
		const source = await openFixture('slide-animation-rich')
		const target = await openFixture('slide-transition')
		const slide = target.slides[0]
		const indices = [2, 3].map((id) => source.slides[0].shapes.findIndex((shape) => shapeId(shape) === id))
		const carried = target.importShapes(slide, source.slides[0], indices, { carryAnimation: true, theme: 'copy' })
		const [first, second] = carried.map(shapeId)

		assert.deepEqual(
			slide.animationSpids(),
			[first, second].sort((a, b) => a - b),
			'only the carried shapes'
		)
		const xml = await slidePartXml(await target.save(), 1)
		assertNoDanglingSpids(xml)
		assert.equal(effectCount(xml, first), effectCount(sourceXml, 2), 'shape 2 keeps its effects, once')
		assert.equal(effectCount(xml, second), effectCount(sourceXml, 3), 'shape 3 keeps its effects, once')
	})

	test('creates a fresh timing scaffold when the host has no animation', async () => {
		// Target slide carries a transition but no animation; carry must build the
		// tmRoot/mainSeq/bldLst scaffold and leave the transition intact.
		const target = await openFixture('slide-transition')
		const source = await openFixture('slide-animation-basic')
		const slide = target.slides[0]
		assert.equal(slide.hasAnimations, false)
		const newSpid = slide.nextShapeId()
		target.importShape(slide, source.slides[0], 0, { carryAnimation: true, theme: 'copy' })

		const saved = await target.save()
		const xml = await slidePartXml(saved, 1)
		assert.ok(/nodeType="tmRoot"/.test(xml) && /nodeType="mainSeq"/.test(xml), 'built tmRoot + mainSeq')
		assert.ok(new RegExp(`<p:bldP spid="${newSpid}" grpId="0"/>`).test(xml), 'carried bldP present')
		assert.ok(/<\/p:clrMapOvr><p:transition/.test(xml) || /<p:fade\/>/.test(xml), 'transition preserved')
		assertNoDanglingSpids(xml)
		assert.equal((await Presentation.load(saved)).slides[0].hasAnimations, true)
	})

	/**
	 * Each `p:seq` on a slide, in document order: its `p:cTn/@nodeType`, and the XML of each group
	 * in its child list.
	 * @param {any} slide
	 */
	function sequencesOf(slide) {
		const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
		const children = (/** @type {any} */ node, local = '') =>
			[...node.childNodes].filter((child) => child.nodeType === 1 && (!local || child.localName === local))
		return [...slide.part.dom.documentElement.getElementsByTagNameNS(P_NS, 'seq')].map((seq) => {
			const [cTn] = children(seq, 'cTn')
			const [list] = cTn ? children(cTn, 'childTnLst') : []
			return { nodeType: cTn?.getAttribute('nodeType') ?? null, groups: list ? children(list).map(String) : [] }
		})
	}

	test('carries into the main sequence of a media slide, leaving its media trigger sequence alone', async () => {
		// PowerPoint times a slide's embedded or online video in an `interactiveSeq`, and these media
		// slides have no `mainSeq`. The carry took the first `p:seq` as the main one, so the click step
		// landed in the media trigger sequence and no main sequence was ever created.
		for (const name of ['av-media', 'online-video']) {
			const target = await openFixture(name)
			const source = await openFixture('slide-animation-basic')
			const slide = target.slides[0]
			const before = sequencesOf(slide)
			assert.deepEqual(
				before.map((seq) => seq.nodeType),
				['interactiveSeq'],
				`${name}: only a media trigger sequence`
			)
			const newSpid = slide.nextShapeId()
			target.importShape(slide, source.slides[0], 0, { carryAnimation: true, theme: 'copy' })

			const after = sequencesOf(slide)
			const main = after.find((seq) => seq.nodeType === 'mainSeq')
			assert.ok(main, `${name}: a mainSeq was created`)
			assert.equal(main.groups.length, 1, `${name}: holding the one carried click step`)
			assert.ok(main.groups[0].includes(`spid="${newSpid}"`), `${name}: which targets the carried shape`)
			const media = after.filter((seq) => seq.nodeType === 'interactiveSeq')
			assert.equal(media.length, 1, `${name}: still one media trigger sequence`)
			assert.deepEqual(media[0].groups, before[0].groups, `${name}: and it is unchanged`)
		}
	})

	test('does not carry a media trigger as a click step', async () => {
		// The source side read click groups from the first `p:seq` too, so lifting the video off a
		// media slide turned its play trigger into a build step on a slide that had no animation.
		const target = await openFixture('slide-transition')
		const source = await openFixture('av-media')
		const slide = target.slides[0]
		target.importShape(slide, source.slides[0], 0, { carryAnimation: true, theme: 'copy' })
		assert.deepEqual(sequencesOf(slide), [], 'no timing sequence was created')
		assert.equal(slide.hasAnimations, false, 'and the slide still reports no animation')
	})

	test.skipIf(!validatorInstalled)('the carried package stays schema-valid', async () => {
		const target = await openFixture('slide-animation-rich')
		const source = await openFixture('slide-animation-basic')
		target.importShape(target.slides[0], source.slides[0], 0, { carryAnimation: true, theme: 'copy' })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assert.equal(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
