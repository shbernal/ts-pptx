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
