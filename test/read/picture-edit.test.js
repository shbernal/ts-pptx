// Phase 4 OPC-mutation tests: adding a picture (new media part + content-type
// registration + image relationship) through the read model.
//
// Contract under test: Slide.addPicture(bytes, ...) creates a /ppt/media part,
// registers its content type, wires an image relationship from the slide, and
// appends a p:pic. The edit survives a save → reopen round-trip (the picture
// resolves its image part), leaves parts it does not touch byte-identical, and
// keeps the package schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const validatorInstalled = await isInstalled()

// A 1×1 transparent PNG.
const PNG_1X1 = new Uint8Array(
	Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC', 'base64')
)

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

async function partBodies(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	const bodies = new Map()
	for (const entry of Object.values(zip.files)) {
		if (entry.dir) continue
		bodies.set(entry.name, await entry.async('uint8array'))
	}
	return bodies
}

function bytesEqual(a, b) {
	return a && b && a.length === b.length && a.every((value, index) => value === b[index])
}

function throws(fn) {
	try {
		fn()
		return false
	} catch {
		return true
	}
}

describe('Slide.addPicture', () => {
	test('adds a media part, relationship, and p:pic that reload correctly', async () => {
		const presentation = await open('empty')
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
		assert(reloaded, 'added picture reloads')
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
		const allowedToChange = new Set([
			'ppt/slides/slide1.xml',
			'ppt/slides/_rels/slide1.xml.rels',
			'[Content_Types].xml',
		])
		for (const [name, body] of inputBodies) {
			if (allowedToChange.has(name)) continue
			assert(bytesEqual(body, outputBodies.get(name)), `${name} should be untouched`)
		}
		const newMedia = [...outputBodies.keys()].filter((name) => name.startsWith('ppt/media/') && !inputBodies.has(name))
		assertEqual(newMedia.length, 1, `exactly one new media part, got ${JSON.stringify(newMedia)}`)
	})

	test('reserveMediaPartName does not collide with an existing image', async () => {
		const presentation = await open('image')
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
		const slide = (await open('empty')).slides[0]
		const notAnImage = new Uint8Array([1, 2, 3, 4])
		assert(
			throws(() => slide.addPicture(notAnImage, { left: 0, top: 0, width: 100000, height: 100000 })),
			'unsniffable bytes without {extension, contentType} should throw'
		)
	})

	test.skipIf(!validatorInstalled)('a deck with an added picture stays schema-valid', async () => {
		const presentation = await open('empty')
		presentation.slides[0].addPicture(PNG_1X1, { left: 914400, top: 457200, width: 1828800, height: 1828800 })
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

describe('Picture.setImage', () => {
	test('mints a new media part, repoints the blip, and leaves the old part untouched', async () => {
		const presentation = await open('image')
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
		const presentation = await open('image')
		const picture = presentation.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		picture.setImage(PNG_1X1, { contentType: 'image/gif' })
		const partName = picture.imagePartName
		assert(partName && partName.endsWith('.gif'), `extension derived from content type: ${partName}`)
		assertEqual(presentation.opc.contentTypes.contentTypeFor(partName), 'image/gif', 'gif content type registered')
	})

	test('throws when no content type is supplied', async () => {
		const picture = (await open('image')).slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		assert(
			throws(() => picture.setImage(PNG_1X1, { contentType: '' })),
			'empty content type should throw'
		)
	})

	test('a sibling picture sharing the old media part is unaffected (copy-on-write)', async () => {
		const presentation = await open('image')
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
		const presentation = await open('image')
		const pictures = presentation.slides[1].shapes.filter((shape) => shape.shapeType === 'picture')
		const [first, second] = pictures
		assert(first && second && first.imageRelId !== second.imageRelId, 'two pictures with distinct rels')

		const before = presentation.opc.parts.size
		first.imageRelId = second.imageRelId
		assertEqual(first.imageRelId, second.imageRelId, 'blip repointed to the chosen rel id')
		assertEqual(presentation.opc.parts.size, before, 'no media part added by the rel-id setter')
	})

	test.skipIf(!validatorInstalled)('a deck with a swapped image stays schema-valid', async () => {
		const presentation = await open('image')
		presentation.slides[0].shapes
			.find((shape) => shape.shapeType === 'picture')
			.setImage(PNG_1X1, { contentType: 'image/png' })
		const errors = await validateBuf(Buffer.from(await presentation.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
