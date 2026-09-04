// Phase 4 OPC-mutation tests: adding a picture (new media part + content-type
// registration + image relationship) through the read model.
//
// Contract under test: Slide.addPicture(bytes, ...) creates a /ppt/media part,
// registers its content type, wires an image relationship from the slide, and
// appends a p:pic. The edit survives a save → reopen round-trip (the picture
// resolves its image part), leaves parts it does not touch byte-identical, and
// keeps the package schema-valid.

import { readFile } from 'node:fs/promises'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import {
	assert,
	assertEqual,
	assertRejects,
	assertUnchangedExcept,
	bytesEqual,
	captureDiagnostics,
	partBodies,
	throws,
} from '../helpers.js'
import { validateBuf, validatorInstalled } from '../validator.js'
import { fixturePath, openFixture } from './corpus.js'

// A 1×1 transparent PNG.
const PNG_1X1 = new Uint8Array(
	Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC', 'base64')
)

/** A 2x1 PNG: twice as wide as it is tall, so a square frame has to crop or letterbox it. */
const PNG_2X1 = new Uint8Array(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAYAAAD0In+KAAAAEklEQVR4nGP8z8DAwMDAxAADAA8sAQdb4A2UAAAAAElFTkSuQmCC',
		'base64'
	)
)

/** The picture's `a:srcRect` attributes, as a plain record (absent attributes are absent). */
function srcRectOf(picture) {
	const match = /<a:srcRect\b([^>]*)\/?>/.exec(picture.element_.toString())
	assert(match, 'expected an a:srcRect; got: ' + picture.element_.toString())
	const attrs = {}
	for (const [, name, value] of match[1].matchAll(/([a-zA-Z]+)="([^"]*)"/g)) attrs[name] = value
	return attrs
}

describe('Slide.addPicture', () => {
	test('adds a media part, relationship, and p:pic that reload correctly', async () => {
		const presentation = await openFixture('empty')
		const slide = presentation.slides[0]
		const picture = slide.addPicture(PNG_1X1, {
			left: 914400,
			top: 457200,
			width: 1828800,
			height: 1828800,
			name: 'Logo',
		})
		assertEqual(picture.shapeType, 'picture', 'returns a Picture')
		assert(picture.imageRelId, 'picture has an embed rel id')

		const saved = await presentation.save()
		const reopened = await Presentation.load(saved)
		const reloaded = reopened.slides[0].shapes.find((shape) => shape.shapeType === 'picture' && shape.name === 'Logo')
		assert(reloaded?.shapeType === 'picture', 'added picture reloads')
		assertEqual(reloaded.width, 1828800, 'geometry reloads')

		const mediaPartName = reloaded.imagePartName
		assert(mediaPartName && mediaPartName.startsWith('/ppt/media/'), `image partname resolves: ${mediaPartName}`)
		const mediaPart = reopened.opc.part(mediaPartName)
		assert(mediaPart, 'media part exists in the reopened package')
		assert(bytesEqual(mediaPart.bytes, PNG_1X1), 'media bytes round-trip unchanged')
		assertEqual(
			reopened.opc.contentTypes.contentTypeFor(mediaPartName),
			'image/png',
			'content type registered as image/png'
		)
	})

	test('leaves parts it does not touch byte-identical and appends the media part', async () => {
		const input = await readFile(fixturePath('empty'))
		const presentation = await Presentation.load(input)
		presentation.slides[0].addPicture(PNG_1X1, { left: 0, top: 0, width: 100000, height: 100000 })
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await presentation.save())

		// The slide, its rels, and the content-types map are allowed to change.
		assertUnchangedExcept(inputBodies, outputBodies, [
			'ppt/slides/slide1.xml',
			'ppt/slides/_rels/slide1.xml.rels',
			'[Content_Types].xml',
		])
		const newMedia = [...outputBodies.keys()].filter((name) => name.startsWith('ppt/media/') && !inputBodies.has(name))
		assertEqual(newMedia.length, 1, `exactly one new media part, got ${JSON.stringify(newMedia)}`)
	})

	test('reserveMediaPartName does not collide with an existing image', async () => {
		const presentation = await openFixture('image')
		const slide =
			presentation.slides.find((s) => s.shapes.some((shape) => shape.shapeType === 'picture')) ?? presentation.slides[0]
		const before = new Set(presentation.opc.parts.keys())
		slide.addPicture(PNG_1X1, {
			left: 0,
			top: 0,
			width: 100000,
			height: 100000,
			extension: 'png',
			contentType: 'image/png',
		})
		const added = [...presentation.opc.parts.keys()].filter((name) => !before.has(name))
		assertEqual(added.length, 1, 'one media part added')
		assert(!before.has(added[0]), `new media partname ${added[0]} did not collide`)
	})

	test('throws when the image type cannot be determined', async () => {
		const slide = (await openFixture('empty')).slides[0]
		const notAnImage = new Uint8Array([1, 2, 3, 4])
		assert(
			throws(() => slide.addPicture(notAnImage, { left: 0, top: 0, width: 100000, height: 100000 })),
			'unsniffable bytes without {extension, contentType} should throw'
		)
	})

	test.skipIf(!validatorInstalled)('a deck with an added picture stays schema-valid', async () => {
		const presentation = await openFixture('empty')
		presentation.slides[0].addPicture(PNG_1X1, { left: 914400, top: 457200, width: 1828800, height: 1828800 })
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

describe('Picture.setImage', () => {
	test('mints a new media part, repoints the blip, and leaves the old part untouched', async () => {
		const presentation = await openFixture('image')
		const slide = presentation.slides[0]
		const picture = slide.shapes.find((shape) => shape.shapeType === 'picture')
		assert(picture, 'fixture slide 1 has a picture')

		const oldRelId = picture.imageRelId
		const oldPartName = picture.imagePartName
		assert(oldPartName, 'picture resolves its original media part')
		const oldBytes = Uint8Array.from(presentation.opc.part(oldPartName).bytes)

		picture.setImage(PNG_1X1, { contentType: 'image/png' })
		assert(picture.imageRelId && picture.imageRelId !== oldRelId, 'blip repointed to a fresh rel id')

		const saved = await presentation.save()
		const reopened = await Presentation.load(saved)
		const reloaded = reopened.slides[0].shapes.find((shape) => shape.shapeType === 'picture')

		const newPartName = reloaded.imagePartName
		assert(
			newPartName && newPartName.startsWith('/ppt/media/') && newPartName.endsWith('.png'),
			`new image partname: ${newPartName}`
		)
		assert(newPartName !== oldPartName, 'blip points at a different media part than before')
		assert(bytesEqual(reopened.opc.part(newPartName).bytes, PNG_1X1), 'new media part holds the supplied bytes')
		assertEqual(reopened.opc.contentTypes.contentTypeFor(newPartName), 'image/png', 'new media content type registered')

		// Copy-on-write fidelity: the original media part survives byte-identical.
		assert(bytesEqual(reopened.opc.part(oldPartName).bytes, oldBytes), 'original media part is untouched')
	})

	test('defaults the media extension from the content type', async () => {
		const presentation = await openFixture('image')
		const picture = presentation.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		picture.setImage(PNG_1X1, { contentType: 'image/gif' })
		const partName = picture.imagePartName
		assert(partName && partName.endsWith('.gif'), `extension derived from content type: ${partName}`)
		assertEqual(presentation.opc.contentTypes.contentTypeFor(partName), 'image/gif', 'gif content type registered')
	})

	// `fit` is the half of `setImage` a swap actually needs: an inherited `a:srcRect` was tuned
	// to the PREVIOUS image's aspect ratio, so a new image of a different ratio reuses a crop
	// that no longer fits. A quarter of this module was unreached and all of it was here.
	test("fit 'stretch' drops the inherited crop", async () => {
		const presentation = await openFixture('image')
		const picture = presentation.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		// Give it a crop to drop, through the same accessor a caller would.
		picture.setImage(PNG_1X1, { contentType: 'image/png', fit: 'cover' })
		assert(/<a:srcRect\b/.test(picture.element_.toString()), 'cover leaves a crop behind')

		picture.setImage(PNG_1X1, { contentType: 'image/png', fit: 'stretch' })
		assert(!/<a:srcRect\b/.test(picture.element_.toString()), 'stretch removes it again')
	})

	test("fit 'cover' and 'contain' crop opposite axes", async () => {
		// A 2x1 image in a square frame: `cover` fills the frame and crops the wide axis,
		// `contain` fits the whole image and letterboxes the tall one. Which axis carries the
		// inset is the whole content of the fit decision.
		const presentation = await openFixture('image')
		const picture = presentation.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		picture.width = 914400
		picture.height = 914400

		picture.setImage(PNG_2X1, { contentType: 'image/png', fit: 'cover' })
		const cover = srcRectOf(picture)
		assert(Number(cover.l) > 0 && Number(cover.r) > 0, `cover crops the wide axis; got ${JSON.stringify(cover)}`)
		assertEqual(cover.t ?? '0', '0', 'and leaves the tall one alone')

		picture.setImage(PNG_2X1, { contentType: 'image/png', fit: 'contain' })
		const contain = srcRectOf(picture)
		assert(
			Number(contain.t) < 0 && Number(contain.b) < 0,
			`contain insets the tall axis; got ${JSON.stringify(contain)}`
		)
		assertEqual(contain.l ?? '0', '0', 'and leaves the wide one alone')
	})

	test('an unmeasurable image leaves the crop alone and says so', async () => {
		// The warn-rather-than-degrade arm: the alternative is a silently stretched picture.
		const presentation = await openFixture('image')
		const picture = presentation.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		const { codes } = await captureDiagnostics(() => {
			picture.setImage(new Uint8Array([1, 2, 3, 4]), { contentType: 'image/png', fit: 'cover' })
		})
		assert(
			codes.includes('image/unmeasurable-natural-size'),
			'expected the unmeasurable warning; got ' + JSON.stringify(codes)
		)
		assert(!/<a:srcRect\b/.test(picture.element_.toString()), 'and no crop was invented')
	})

	test('fit needs a frame extent, and says which picture has none', async () => {
		const presentation = await openFixture('image')
		const picture = presentation.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		// Drop the transform: a picture inheriting its box from a placeholder has no `a:ext` to
		// measure the crop against, and guessing one would be a crop nobody asked for.
		const xfrm = picture.element_.getElementsByTagName('a:xfrm')[0]
		xfrm.parentNode.removeChild(xfrm)
		await assertRejects(
			() => picture.setImage(PNG_1X1, { contentType: 'image/png', fit: 'contain' }),
			/needs a frame extent/,
			'setImage with a fit on a picture with no transform'
		)
	})

	test('throws when no content type is supplied', async () => {
		const picture = (await openFixture('image')).slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		assert(
			throws(() => picture.setImage(PNG_1X1, { contentType: '' })),
			'empty content type should throw'
		)
	})

	test('a sibling picture sharing the old media part is unaffected (copy-on-write)', async () => {
		const presentation = await openFixture('image')
		// On fixture slide 2, two pictures embed the same rel (image2.png).
		const slide = presentation.slides[1]
		const pictures = slide.shapes.filter((shape) => shape.shapeType === 'picture')
		const counts = new Map()
		for (const pic of pictures) counts.set(pic.imageRelId, (counts.get(pic.imageRelId) ?? 0) + 1)
		const sharedRelId = [...counts].find(([, n]) => n >= 2)?.[0]
		assert(sharedRelId, 'fixture slide 2 has two pictures sharing one image rel')

		const shared = pictures.filter((pic) => pic.imageRelId === sharedRelId)
		const sharedPartName = shared[0].imagePartName
		const sharedBytes = Uint8Array.from(presentation.opc.part(sharedPartName).bytes)

		shared[0].setImage(PNG_1X1, { contentType: 'image/png' })

		assert(shared[1].imageRelId === sharedRelId, 'the sibling picture still points at the shared rel')
		assertEqual(shared[1].imagePartName, sharedPartName, 'the sibling still resolves the original media part')
		assert(
			bytesEqual(presentation.opc.part(sharedPartName).bytes, sharedBytes),
			'the shared media part bytes are unchanged'
		)
	})

	test('imageRelId setter repoints the blip without adding a media part', async () => {
		const presentation = await openFixture('image')
		const pictures = presentation.slides[1].shapes.filter((shape) => shape.shapeType === 'picture')
		const [first, second] = pictures
		assert(first && second && first.imageRelId !== second.imageRelId, 'two pictures with distinct rels')

		const before = presentation.opc.parts.size
		first.imageRelId = second.imageRelId
		assertEqual(first.imageRelId, second.imageRelId, 'blip repointed to the chosen rel id')
		assertEqual(presentation.opc.parts.size, before, 'no media part added by the rel-id setter')
	})

	test.skipIf(!validatorInstalled)('a deck with a swapped image stays schema-valid', async () => {
		const presentation = await openFixture('image')
		presentation.slides[0].shapes
			.find((shape) => shape.shapeType === 'picture')
			.setImage(PNG_1X1, { contentType: 'image/png' })
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
