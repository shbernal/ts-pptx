// A background that belongs to the SLIDE, not to its layout or master.
//
// The converter had two background mappers and only one of them could carry a picture. The
// layout/master arm handled `solid`, `image` and `themeRef`; the slide arm handled `solid` and
// `none` and recorded everything else as "not expressible through the write API's background
// option" — a claim that was false for `image`, since `SlideProps.background` takes the same
// `BackgroundProps` the layout arm authors and `BackgroundIr.data` was a declared, documented
// field nothing could produce for a slide.
//
// Neither gate could see it. The round trip excludes exactly the *declared* losses, and
// `slide.background` was declared — so a dropped picture background was excused by the very
// note that misdescribed it. The byte-identity gate reaches no read fixture at all. So these
// cases assert what the read model and the IR say, against a deck whose three slide-scoped
// backgrounds are the corpus' only instances of each.

import JSZip from 'jszip'
import { describe, expect, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'
import { irFor, openFixture } from './corpus.js'

const FIXTURE = 'slide-background.pptx'

const GRADIENT = {
	type: /** @type {const} */ ('gradient'),
	gradient: {
		kind: /** @type {const} */ ('linear'),
		angle: 90,
		stops: [
			{ color: 'FF0000', position: 0 },
			{ color: '0000FF', position: 100 },
		],
	},
}
const PATTERN = {
	type: /** @type {const} */ ('pattern'),
	pattern: { preset: /** @type {const} */ ('dkHorz'), fgColor: 'accent1', bgColor: 'FFFFFF' },
}

describe('the read model resolves a slide-scoped background of every kind', () => {
	test('a slide’s own picture, theme reference and transparent solid are all read', async () => {
		const pres = await openFixture(FIXTURE)
		const [picture, themeRef, translucent] = pres.slides.map((slide) => slide.background)

		// `throw` rather than `assertEqual` on the three discriminants: it is what narrows the
		// union, so the reads below are checked rather than cast.
		if (picture?.type !== 'image') throw new Error(`slide 1 must read as an image background, got ${picture?.type}`)
		if (themeRef?.type !== 'themeRef') throw new Error(`slide 2 must read as a theme reference, got ${themeRef?.type}`)
		if (translucent?.type !== 'solid') throw new Error(`slide 3 must read as a solid, got ${translucent?.type}`)

		assertEqual(picture.source, 'slide', 'the picture background belongs to the slide, not the layout')
		assert(picture.partName?.endsWith('.png'), `the image resolves to a media part (got ${picture.partName})`)

		assertEqual(themeRef.source, 'slide', 'a p:bgRef can be slide-scoped, not only layout-scoped')
		assertEqual(themeRef.idx, 1001, 'the raw idx is kept for fidelity')
		assertEqual(themeRef.resolvedFill?.type, 'solid', 'bgFillStyleLst[0] is a solid fill in the stock theme')

		assertEqual(translucent.color?.alpha, 0.6, '<a:alpha val="60000"/> reads as 0.6 opacity')
	})
})

describe('the converter carries a slide-scoped background instead of dropping it', () => {
	test('a picture background becomes an asset reference rather than a fidelity note', async () => {
		const ir = await irFor(FIXTURE)
		const background = ir.slides[0].background
		assert(background, 'slide 1 keeps its own background')
		assert(
			background.data,
			`a slide-scoped picture background reaches BackgroundIr.data (got ${JSON.stringify(background)})`
		)
		const asset = ir.assets.find((a) => a.name === background.data.$asset)
		assert(asset, 'the referenced asset is registered, so the emitted script can carry the bytes')
		assertEqual(asset.contentType, 'image/png', 'with the source part’s content type')
	})

	test('a theme reference bakes its resolved colour and says so, at slide scope', async () => {
		const ir = await irFor(FIXTURE)
		const background = ir.slides[1].background
		assert(background, 'slide 2 keeps its own background')
		assertEqual(background.color, 'E97132', 'the colour the p:bgRef currently resolves to is baked in')
		const note = ir.fidelity.find((n) => n.construct === 'slide.background')
		assert(note, 'the loss of the reference itself is recorded')
		assertEqual(
			note.disposition,
			'flattened',
			'flattened rather than dropped — the colour survives, the reference does not'
		)
		// The construct stays tier-scoped even though one mapper now serves both tiers: a
		// layout losing its background is `master.background` and says something different.
		assert(note.detail.includes('slide'), `the slide tier’s note names the slide (got ${JSON.stringify(note.detail)})`)
	})

	test('a background transparency reaches the IR, which nothing used to produce', async () => {
		const ir = await irFor(FIXTURE)
		const background = ir.slides[2].background
		assert(background, 'slide 3 keeps its own background')
		assertEqual(background.color, 'C83C28', 'the colour survives')
		// The read model reports opacity as a 0-1 fraction and the write option takes
		// transparency as a 0-100 percent. `BackgroundIr.transparency` was declared,
		// documented and compared by the round-trip verifier with no producer on either arm.
		assertEqual(background.transparency, 40, '0.6 opacity is 40% transparency')
	})

	// Both were noted "not expressible through the write API's background option", which was false:
	// `BackgroundProps` extends `ShapeFillProps` and the background emitter writes either.
	test('a gradient and a pattern background carry, on a slide and on a layout', async () => {
		const { presentation } = await authorRead((pptx) => {
			pptx.defineSlideMaster({ title: 'GRADIENT', background: GRADIENT })
			pptx.addSlide().background = GRADIENT
			pptx.addSlide().background = PATTERN
			pptx.addSlide({ masterTitle: 'GRADIENT' })
		})
		const ir = readModelToIr(presentation)
		expect(ir.slides[0].background).toEqual(GRADIENT)
		expect(ir.slides[1].background).toEqual(PATTERN)
		assertEqual(ir.slides[2].background, undefined, 'the third slide inherits its layout’s')
		const layout = ir.chrome.masters.find((master) => master.props.title === 'GRADIENT')
		expect(layout?.props.background).toEqual(GRADIENT)
		const losses = ir.fidelity.filter((note) => note.construct.includes('background') && note.disposition === 'dropped')
		assertEqual(losses.length, 0, `no background is dropped (got ${JSON.stringify(losses)})`)
	})

	test('a theme reference resolving to a gradient carries the gradient and says it is baked in', async () => {
		const { buf } = await authorRead((pptx) => {
			pptx.addSlide()
		})
		const zip = await JSZip.loadAsync(buf)
		const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')
		// The write path's theme holds a gradient in the third `bgFillStyleLst` slot.
		const bg = '<p:bg><p:bgRef idx="1003"><a:schemeClr val="accent1"/></p:bgRef></p:bg>'
		zip.file('ppt/slides/slide1.xml', slideXml.replace(/(<p:cSld[^>]*>)/, `$1${bg}`))
		const pres = await Presentation.load(await zip.generateAsync({ type: 'uint8array' }))
		assertEqual(pres.slides[0].background?.type, 'themeRef', 'the patched slide reads a theme reference')

		const ir = readModelToIr(pres)
		assertEqual(ir.slides[0].background?.type, 'gradient', 'the gradient it resolves to is carried')
		const note = ir.fidelity.find((n) => n.construct === 'slide.background')
		assertEqual(note?.disposition, 'flattened', 'and the reference itself is noted as flattened')
		assert(note.detail.includes('gradient it currently resolves to'), `the note names the fill (got ${note.detail})`)
	})

	test('an inherited background is still not re-authored onto the slide', async () => {
		// The one thing that must NOT change: a background that comes from the layout stays
		// there, so it keeps following the layout instead of being pinned onto every slide.
		const ir = await irFor('empty.pptx')
		assert(
			ir.slides.every((slide) => slide.background === undefined),
			'a slide with no background of its own states none'
		)
	})
})
