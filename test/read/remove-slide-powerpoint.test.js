// Presentation.removeSlide against PowerPoint's own deletion. slide-jump-link.pptx is the deck
// before: a target slide (id 256), a referrer (257) whose LinkedText run and LinkedShape both jump
// to the target and whose ControlText run jumps to slide 3 (258), a section per side, and a custom
// show "Tour" listing the target and the referrer. slide-jump-link-target-deleted.pptx is the same
// deck after PowerPoint deleted the target and saved.
//
// PowerPoint keeps the "First" section with an empty p14:sldIdLst and drops the target from "Tour",
// and removeSlide has to leave both the same way. The jump links are where the two part company on
// purpose: PowerPoint keeps each a:hlinkClick and its relationship, and once it renumbers the
// slides that relationship names slide1.xml, which is the referrer itself. A link that silently
// starts jumping somewhere else is the failure removeSlide unlinks to avoid.

import { describe, test } from 'vitest'

import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const PRESENTATION = '/ppt/presentation.xml'

/** The presentation part's XML as the deck currently holds it. */
function presentationXml(deck) {
	return new TextDecoder().decode(deck.opc.part(PRESENTATION).serialize())
}

/** Each section as `name: id id …`, in order. */
function sections(deck) {
	return [...presentationXml(deck).matchAll(/<p14:section name="([^"]*)"[^>]*>([\s\S]*?)<\/p14:section>/g)].map(
		([, name, body]) => `${name}: ${[...body.matchAll(/<p14:sldId id="(\d+)"/g)].map(([, id]) => id).join(' ')}`
	)
}

/** The slide ids each custom show lists, as `name: id id …`, resolved through the relationships. */
function customShows(deck) {
	const rels = deck.opc.relationshipsFor(PRESENTATION)
	const idOf = (relId) => deck.slides.find((slide) => slide.partName === rels.resolveTarget(relId))?.slideId ?? '?'
	return [...presentationXml(deck).matchAll(/<p:custShow name="([^"]*)"[^>]*>([\s\S]*?)<\/p:custShow>/g)].map(
		([, name, body]) =>
			`${name}: ${[...body.matchAll(/<p:sld r:id="([^"]+)"/g)].map(([, relId]) => idOf(relId)).join(' ')}`
	)
}

/** Where the click action on the shape named `name` jumps: the target slide's id, or `null` for no link. */
function jumpOf(deck, slide, name) {
	const shape = slide.shapes.find((s) => s.name === name)
	assert(shape, `expected a shape named ${name}`)
	const link = shape.element_.getElementsByTagNameNS(
		'http://schemas.openxmlformats.org/drawingml/2006/main',
		'hlinkClick'
	)[0]
	if (!link) return null
	const relId = link.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
	const target = deck.opc.relationshipsFor(slide.partName).resolveTarget(relId)
	return deck.slides.find((s) => s.partName === target)?.slideId ?? target
}

describe("Presentation.removeSlide against PowerPoint's own deletion (slide-jump-link.pptx)", () => {
	test('the removed slide leaves its section, which stays, and the custom show', async () => {
		const deck = await openFixture('slide-jump-link')
		assertEqual(sections(deck).join(' | '), 'First: 256 | Rest: 257 258', 'the before deck')
		assertEqual(customShows(deck).join(' | '), 'Tour: 256 257', 'the before deck')

		deck.removeSlide(0)
		const powerPoint = await openFixture('slide-jump-link-target-deleted')
		assertEqual(sections(deck).join(' | '), sections(powerPoint).join(' | '), 'sections as PowerPoint leaves them')
		assertEqual(
			customShows(deck).join(' | '),
			customShows(powerPoint).join(' | '),
			'custom shows as PowerPoint leaves them'
		)
	})

	test('a jump link to the removed slide is dropped where PowerPoint leaves it jumping to another slide', async () => {
		const powerPoint = await openFixture('slide-jump-link-target-deleted')
		const theirReferrer = powerPoint.slides[0]
		assertEqual(
			jumpOf(powerPoint, theirReferrer, 'LinkedText'),
			theirReferrer.slideId,
			'PowerPoint: the text now jumps to its own slide'
		)
		assertEqual(jumpOf(powerPoint, theirReferrer, 'LinkedShape'), theirReferrer.slideId, 'and so does the shape')

		const deck = await openFixture('slide-jump-link')
		deck.removeSlide(0)
		const referrer = deck.slides[0]
		assertEqual(jumpOf(deck, referrer, 'LinkedText'), null, 'removeSlide: the text link is gone')
		assertEqual(jumpOf(deck, referrer, 'LinkedShape'), null, 'and the shape link')
		assertEqual(jumpOf(deck, referrer, 'ControlText'), 258, 'while a link to a slide that stays is untouched')
		assertEqual(jumpOf(powerPoint, theirReferrer, 'ControlText'), 258, 'as PowerPoint leaves it')
	})
})
