// Read-model coverage for src/read/api/transition.ts — findTransition /
// parseTransition / parseSound / prefixFor / variantAttrs on the getter side and
// buildTransition / buildTransitionElement / speedForDuration / removeTransition
// on the setter side. The parse/build helpers are not exported; they run only
// through the Slide.transition accessor. The real fixture decks carry one
// well-formed transition apiece, so the wrapped-vs-degenerate AlternateContent
// forms, the p15/p159/unknown/no-namespace type prefixes, the sound sub-forms,
// and the fast/med speed buckets never all fire. Here a synthetic Slide (a
// hand-authored `p:sld` fed through `new Part`, wrapped in
// `new Slide(null, part, id, idx)` — the accessor touches only `this.part.dom`,
// never the package) drives each branch directly, including malformed trees.

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

describe('parseTransition — AlternateContent forms', () => {
	test('a bare <p:transition> decodes its type, speed and variant', () => {
		const info = slide('<p:transition spd="med"><p:push dir="d"/></p:transition>').transition
		assert(info, 'the bare transition is decoded')
		assertEqual(info.type, 'push', 'type from the single non-sound child')
		assertEqual(info.namespace, 'p', 'a base ECMA-376 type uses the p prefix')
		assertEqual(info.speed, 'med', 'spd bucket carried through')
		assertEqual(info.durationMs, null, 'no p14:dur on the bare form')
		assertEqual(JSON.stringify(info.variant), '{"dir":"d"}', 'type attributes become the variant')
	})

	test('an mc:AlternateContent p14 Choice is preferred so the exact duration is recovered', () => {
		const info = slide(
			`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ` +
				`xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">` +
				`<mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2000"><p:fade/></p:transition></mc:Choice>` +
				`<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>` +
				`</mc:AlternateContent>`
		).transition
		assert(info, 'the wrapped transition is decoded')
		assertEqual(info.type, 'fade', 'the Choice transition is read')
		assertEqual(info.durationMs, 2000, 'p14:dur is recovered from the Choice')
		assertEqual(info.speed, 'slow', 'the coarse bucket is preserved alongside the exact duration')
	})

	test('a degenerate AlternateContent with only a Fallback is still surfaced', () => {
		const info = slide(
			`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
				`<mc:Fallback><p:transition><p:wipe dir="u"/></p:transition></mc:Fallback>` +
				`</mc:AlternateContent>`
		).transition
		assert(info, 'the Fallback-only form is not lost')
		assertEqual(info.type, 'wipe', 'the Fallback transition is read')
		assertEqual(info.durationMs, null, 'the Fallback carries no exact duration')
	})

	test('a transition with no type child (sound only) decodes to null', () => {
		const info = slide('<p:transition><p:sndAc><p:endSnd/></p:sndAc></p:transition>').transition
		assertEqual(info, null, 'a transition without a type element is malformed → null')
	})

	test('a slide with no transition reports null', () => {
		assertEqual(slide('').transition, null, 'no p:transition and no AlternateContent → null')
	})
})

describe('prefixFor — modern type namespaces', () => {
	/** A transition whose type element is `<localName …decl spokes="8"/>`. */
	const typedSlide = (opening) => slide(`<p:transition><${opening} spokes="8"/></p:transition>`)

	test('a p15 type element reports the p15 prefix, with the xmlns declaration excluded from the variant', () => {
		const info = typedSlide(`p15:flash xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"`).transition
		assertEqual(info.namespace, 'p15', 'the 2012 namespace maps to p15')
		assertEqual(JSON.stringify(info.variant), '{"spokes":"8"}', 'the xmlns:p15 declaration is not a variant attribute')
	})

	test('a p159 type element reports the p159 prefix', () => {
		const info = typedSlide(
			`p159:morph xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/9/main"`
		).transition
		assertEqual(info.namespace, 'p159', 'the 2015/9 namespace maps to p159')
	})

	test('an unrecognised type namespace falls back to the p prefix', () => {
		const info = slide('<p:transition><x:glitz xmlns:x="urn:example:unknown"/></p:transition>').transition
		assertEqual(info.namespace, 'p', 'an unknown namespace URI defaults to p')
	})

	test('a type element with no namespace falls back to the p prefix', () => {
		const info = slide('<p:transition><cut/></p:transition>').transition
		assertEqual(info.type, 'cut', 'the no-namespace child is still the type')
		assertEqual(info.namespace, 'p', 'a null namespace URI defaults to p')
	})
})

describe('parseSound — start / stop / empty', () => {
	test('a start sound (p:stSnd) is decoded, and a sound child before the type is skipped', () => {
		// The sndAc sits before the type child, so typeElement must skip it (a NON_TYPE child)
		// before finding p:fade.
		const info = slide(
			`<p:transition xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
				`<p:sndAc><p:stSnd loop="1"><p:snd r:embed="rId5" name="chime.wav"/></p:stSnd></p:sndAc>` +
				`<p:fade/></p:transition>`
		).transition
		assertEqual(info.type, 'fade', 'the type is found past the leading sound child')
		assert(info.sound, 'a sound is decoded')
		assertEqual(info.sound.form, 'start', 'p:stSnd → start form')
		assertEqual(info.sound.loop, true, 'loop="1" → looping')
		assertEqual(info.sound.embedRid, 'rId5', 'the embed relationship id is read')
		assertEqual(info.sound.name, 'chime.wav', 'the sound name is read')
	})

	test('a stop sound (p:endSnd) is the stop-previous form', () => {
		const info = slide('<p:transition><p:fade/><p:sndAc><p:endSnd/></p:sndAc></p:transition>').transition
		assert(info.sound, 'a stop sound is decoded')
		assertEqual(info.sound.form, 'stop', 'p:endSnd → stop form')
		assertEqual(info.sound.loop, false, 'the stop form never loops')
		assertEqual(info.sound.embedRid, null, 'the stop form has no embed')
	})

	test('a sndAc with neither start nor stop sound decodes to no sound', () => {
		const info = slide('<p:transition><p:fade/><p:sndAc/></p:transition>').transition
		assertEqual(info.sound, null, 'an empty sound action is not a sound')
	})

	test('a start sound with no p:snd child and no loop attribute decodes to bare defaults', () => {
		// stSnd present but empty: loop falls back to false, and embedRid/name have no p:snd to read.
		const info = slide('<p:transition><p:fade/><p:sndAc><p:stSnd/></p:sndAc></p:transition>').transition
		assert(info.sound, 'the start form is still recognised')
		assertEqual(info.sound.form, 'start', 'p:stSnd → start form')
		assertEqual(info.sound.loop, false, 'a missing loop attribute defaults to non-looping')
		assertEqual(info.sound.embedRid, null, 'no p:snd → no embed id')
		assertEqual(info.sound.name, null, 'no p:snd → no name')
	})
})

describe('findTransition — degenerate AlternateContent', () => {
	test('an AlternateContent with neither a Choice transition nor a Fallback yields no transition', () => {
		const info = slide(
			`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
				`<mc:Choice Requires="p99"/></mc:AlternateContent>`
		).transition
		assertEqual(info, null, 'no usable transition inside the wrapper → null')
	})
})

describe('buildTransition + speedForDuration (setter) — buckets and forms', () => {
	test('setting a bare transition writes no spd and defaults the read-back speed to fast', () => {
		const s = slide('') // no existing transition → removeTransition finds nothing to remove
		s.transition = { type: 'cut' }
		const info = s.transition
		assertEqual(info.type, 'cut', 'the type round-trips')
		assertEqual(info.speed, 'fast', 'an absent spd reads as the fast default')
		assertEqual(info.durationMs, null, 'no duration was requested')
	})

	test('a sub-500ms duration derives the fast bucket', () => {
		const s = slide('')
		s.transition = { type: 'fade', durationMs: 400 }
		const info = s.transition
		assertEqual(info.durationMs, 400, 'the exact duration is written to the Choice')
		assertEqual(info.speed, 'fast', '400ms → fast bucket')
	})

	test('a 500–1000ms duration derives the med bucket', () => {
		const s = slide('')
		s.transition = { type: 'fade', durationMs: 800 }
		const info = s.transition
		assertEqual(info.durationMs, 800, 'the exact duration is written to the Choice')
		assertEqual(info.speed, 'med', '800ms → med bucket')
	})

	test('assigning null removes an existing transition', () => {
		const s = slide('<p:transition spd="med"><p:wipe dir="u"/></p:transition>')
		assert(s.transition, 'the transition is present to begin with')
		s.transition = null
		assertEqual(s.transition, null, 'clearing removes the transition node')
	})
})
