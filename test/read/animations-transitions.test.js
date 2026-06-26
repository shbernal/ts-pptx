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
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

async function loadOracle(name) {
	return JSON.parse(await readFile(path.join(__dirname, 'fixtures', `${name}.oracle.json`), 'utf8'))
}

async function slidePartXml(pptxBytes, slideNumber) {
	const zip = await JSZip.loadAsync(pptxBytes)
	return zip.file(`ppt/slides/slide${slideNumber}.xml`).async('string')
}

async function partBodies(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	const bodies = new Map()
	for (const entry of Object.values(zip.files)) {
		if (entry.dir) continue
		bodies.set(entry.name, await entry.async('string'))
	}
	return bodies
}

describe('slide.transition (read)', () => {
	test('decodes every PowerPoint-authored transition form', async () => {
		const oracle = await loadOracle('slide-transition')
		const pres = await open('slide-transition')
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

	test('untouched slides round-trip byte-identically', async () => {
		const original = await readFile(fixturePath('slide-transition'))
		const pres = await Presentation.load(original)
		// touch nothing
		const before = await partBodies(original)
		const after = await partBodies(await pres.save())
		for (const [name, body] of before) {
			if (name.startsWith('ppt/slides/slide')) assert.equal(after.get(name), body, `${name} unchanged`)
		}
	})
})

describe('slide.transition (write/edit)', () => {
	test('sets a bare transition when no duration is given', async () => {
		const pres = await open('slide-transition')
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
		const pres = await open('slide-transition')
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
		const pres = await open('slide-transition')
		pres.slides[0].transition = { type: 'fade', durationMs: 1500 }
		const reopened = await Presentation.load(await pres.save())
		assert.equal(reopened.slides[0].transition.speed, 'slow')
	})

	test('round-trips advTm / advClick auto-advance', async () => {
		const pres = await open('slide-transition')
		pres.slides[0].transition = { type: 'fade', speed: 'med', advanceOnClick: false, advanceAfterMs: 3000 }
		const reopened = await Presentation.load(await pres.save())
		const info = reopened.slides[0].transition
		assert.equal(info.advanceOnClick, false)
		assert.equal(info.advanceAfterMs, 3000)
	})

	test('assigning null clears the transition', async () => {
		const pres = await open('slide-transition')
		pres.slides[0].transition = null
		const saved = await pres.save()
		const xml = await slidePartXml(saved, 1)
		assert.ok(!xml.includes('<p:transition'), 'no p:transition element remains')
		const reopened = await Presentation.load(saved)
		assert.equal(reopened.slides[0].transition, null)
	})
})

describe('slide animations (opaque, spid-aware)', () => {
	test('hasAnimations + animationSpids on the basic fixture', async () => {
		const oracle = await loadOracle('slide-animation-basic')
		const pres = await open('slide-animation-basic')
		const slide = pres.slides[0]
		assert.equal(slide.hasAnimations, true)
		assert.deepEqual(slide.animationSpids(), oracle.animationSpids)
	})

	test('hasAnimations + animationSpids on the rich fixture', async () => {
		const oracle = await loadOracle('slide-animation-rich')
		const pres = await open('slide-animation-rich')
		const slide = pres.slides[0]
		assert.equal(slide.hasAnimations, true)
		assert.deepEqual(slide.animationSpids(), oracle.animationSpids)
	})

	test('an unanimated fixture reports no animations', async () => {
		const pres = await open('slide-transition')
		assert.equal(pres.slides[0].hasAnimations, false)
		assert.deepEqual(pres.slides[0].animationSpids(), [])
	})

	test('untouched animated slides round-trip byte-identically', async () => {
		const original = await readFile(fixturePath('slide-animation-rich'))
		const pres = await Presentation.load(original)
		const before = await partBodies(original)
		const after = await partBodies(await pres.save())
		for (const [name, body] of before) {
			if (name.startsWith('ppt/slides/slide')) assert.equal(after.get(name), body, `${name} unchanged`)
		}
	})

	test('remapAnimationSpids rewrites every spTgt and bldP reference', async () => {
		const pres = await open('slide-animation-rich')
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
		const pres = await open('slide-animation-rich')
		const slide = pres.slides[0]
		slide.pruneAnimationSpids([3])
		assert.deepEqual(slide.animationSpids(), [2, 4, 5])
		const reopened = await Presentation.load(await pres.save())
		assert.deepEqual(reopened.slides[0].animationSpids(), [2, 4, 5])
	})

	test('pruning a click-effect collapses its emptied wrapper', async () => {
		const pres = await open('slide-animation-rich')
		const slide = pres.slides[0]
		slide.pruneAnimationSpids([2])
		assert.deepEqual(slide.animationSpids(), [3, 4, 5])
	})
})

describe('slide.flattenAnimations (whole-slide flatten pass)', () => {
	test('strips the whole timing block, leaving every shape in place', async () => {
		const pres = await open('slide-animation-rich')
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
		const pres = await open('slide-animation-basic')
		const slide = pres.slides[0]
		assert.equal(slide.flattenAnimations(), true)
		assert.equal(slide.hasAnimations, false)
		assert.deepEqual(slide.animationSpids(), [])
	})

	test('is a no-op on an unanimated slide and is idempotent', async () => {
		const pres = await open('slide-transition')
		const slide = pres.slides[0]
		assert.equal(slide.hasAnimations, false)
		assert.equal(slide.flattenAnimations(), false, 'nothing to flatten')
		// the slide-show transition is untouched by an animation flatten
		assert.notEqual(slide.transition, null)

		const rich = (await open('slide-animation-rich')).slides[0]
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
		const oracle = await loadOracle('slide-animation-presets')
		const slide = (await open('slide-animation-presets')).slides[0]
		assert.equal(slide.hasAnimations, true)
		assert.deepEqual(slide.animationSpids(), oracle.animationSpids)
	})

	test('every preset template appears verbatim in the slide', async () => {
		const oracle = await loadOracle('slide-animation-presets')
		const xml = await slidePartXml(await readFile(fixturePath('slide-animation-presets')), 1)
		for (const [name, t] of Object.entries(oracle.presetTemplates)) {
			assert.ok(xml.includes(t.effectParXml), `${name} effect node present verbatim`)
			assert.ok(xml.includes(t.behaviorsXml), `${name} behaviors present verbatim`)
			assert.ok(xml.includes(t.bldPXml), `${name} bldP present`)
		}
	})

	test('untouched presets slide round-trips byte-identically', async () => {
		const original = await readFile(fixturePath('slide-animation-presets'))
		const before = await partBodies(original)
		const after = await partBodies(await (await Presentation.load(original)).save())
		for (const [name, body] of before) {
			if (name.startsWith('ppt/slides/slide')) assert.equal(after.get(name), body, `${name} unchanged`)
		}
	})
})

describe('slide-transition-sound (read fixture, Phase 2 gate C)', () => {
	test('decodes the fade transition on every slide', async () => {
		const oracle = await loadOracle('slide-transition-sound')
		const pres = await open('slide-transition-sound')
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
		const oracle = await loadOracle('slide-transition-sound')
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

	test('untouched transition-sound slides round-trip byte-identically', async () => {
		const original = await readFile(fixturePath('slide-transition-sound'))
		const before = await partBodies(original)
		const after = await partBodies(await (await Presentation.load(original)).save())
		for (const [name, body] of before) {
			if (name.startsWith('ppt/slides/slide') || name.startsWith('ppt/media/')) {
				assert.equal(after.get(name), body, `${name} unchanged`)
			}
		}
	})
})

describe('import-animation-merge (read fixture, Phase 2 gate A)', () => {
	test('enumerates spids on both slides per the oracle', async () => {
		const oracle = await loadOracle('import-animation-merge')
		const pres = await open('import-animation-merge')
		assert.deepEqual(pres.slides[0].animationSpids(), oracle.source.animationSpids)
		assert.deepEqual(pres.slides[1].animationSpids(), oracle.merged.animationSpids)
	})

	test('the merged slide matches the oracle timing verbatim', async () => {
		const oracle = await loadOracle('import-animation-merge')
		const xml = await slidePartXml(await readFile(fixturePath('import-animation-merge')), 2)
		assert.ok(xml.includes(oracle.merged.timingXml), 'merged timing tree present verbatim')
		assert.ok(xml.includes(oracle.merged.bldList.xml), 'merged bldLst present verbatim')
	})

	test('remapAnimationSpids stays coherent across the merged build', async () => {
		const oracle = await loadOracle('import-animation-merge')
		const pres = await open('import-animation-merge')
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
// oracle. The destination timing is PptxGenJS's own construction (not byte-equal to
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
		const target = await open('slide-transition')
		const source = await open('slide-animation-basic')
		target.importShape(target.slides[0], source.slides[0], 0, { theme: 'copy' })
		assert.equal(target.slides[0].hasAnimations, false, 'no animation carried without the flag')
	})

	test('appends the carried build after the host build, remapped to the new spid', async () => {
		// Host already animates spids 2..5; the lifted basic shape takes the next id (6).
		const target = await open('slide-animation-rich')
		const source = await open('slide-animation-basic')
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

	test('creates a fresh timing scaffold when the host has no animation', async () => {
		// Target slide carries a transition but no animation; carry must build the
		// tmRoot/mainSeq/bldLst scaffold and leave the transition intact.
		const target = await open('slide-transition')
		const source = await open('slide-animation-basic')
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

	test.skipIf(!validatorInstalled)('the carried package stays schema-valid', async () => {
		const target = await open('slide-animation-rich')
		const source = await open('slide-animation-basic')
		target.importShape(target.slides[0], source.slides[0], 0, { carryAnimation: true, theme: 'copy' })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assert.equal(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
