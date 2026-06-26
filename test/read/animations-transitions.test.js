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

const __dirname = path.dirname(fileURLToPath(import.meta.url))

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
