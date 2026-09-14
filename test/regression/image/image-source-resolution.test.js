// How the definers read an image source: the media type an inline payload states, and what each
// kind of image does with a source it cannot use. A picture has nothing to degrade to and throws, a
// cover or preview falls back to the gray placeholder, and a background is dropped, each with a
// diagnostic. Before, three of those sites checked no header at all and wrote the caller's text
// into an image part.
//
// `dataUriMediaType` and `imageExtensionForSource` are imported from `src/`, as the format registry
// suite imports its tables: they are internal, and what they decide is asserted through the package
// in the cases below.

import { describe, test } from 'vitest'
import { dataUriMediaType, imageExtensionForSource } from '../../../src/media/content-type.ts'
import {
	PNG_1X1,
	assert,
	assertEqual,
	assertIncludes,
	build,
	captureDiagnostics,
	listEntries,
	readEntry,
} from '../../helpers.js'

/** A few bytes standing in for a metafile. Nothing reads them as one; only the declared type matters. */
const METAFILE = 'AQAAAGwAAAAAAAAAAAAAAA=='

/** The package's media parts, by name, with their bytes. */
async function mediaParts(zip) {
	const names = listEntries(zip).filter((name) => name.startsWith('ppt/media/'))
	return Promise.all(names.map(async (name) => ({ name, bytes: await zip.file(name).async('uint8array') })))
}

describe('dataUriMediaType', () => {
	test('reads the whole media type, with or without the data: scheme', () => {
		/** @type {Array<[string, string | null]>} */
		const cases = [
			['data:image/x-emf;base64,AA', 'image/x-emf'],
			['image/x-wmf;base64,AA', 'image/x-wmf'],
			['data:image/svg+xml;base64,AA', 'image/svg+xml'],
			['data:image/vnd.microsoft.icon;base64,AA', 'image/vnd.microsoft.icon'],
			['data:image/JPEG;base64,AA', 'image/jpeg'],
			['data:audio/x-wav;base64,AA', 'audio/x-wav'],
			['data:image/svg+xml,%3Csvg%3E', 'image/svg+xml'],
			['iVBORw0KGgo=', null],
			['', null],
		]
		for (const [data, expected] of cases) assertEqual(dataUriMediaType(data), expected, data)
	})

	test('names an image part by the registry extension, and by the subtype when the registry has no row', () => {
		/** @type {Array<[string, string, string]>} */
		const cases = [
			['', 'data:image/x-emf;base64,AA', 'emf'],
			['', 'data:image/x-wmf;base64,AA', 'wmf'],
			['', 'data:image/svg+xml;base64,AA', 'svg'],
			['', 'data:image/jpeg;base64,AA', 'jpeg'],
			['', 'image/jpg;base64,AA', 'jpg'],
			['photo.PNG', '', 'png'],
			['photo.jpg', 'data:image/x-emf;base64,AA', 'emf'],
		]
		for (const [path, data, expected] of cases)
			assertEqual(imageExtensionForSource(path, data), expected, `path ${path}, data ${data}`)
	})
})

describe('image sources through the definers', () => {
	test('a metafile data URI is written as an emf or wmf part with its own content type', async () => {
		// The type used to be read with a pattern that stopped at the `-`, so both came out as `.png`
		// parts declared `image/png`.
		const { zip } = await build((p) => {
			p.addSlide().addImage({ data: `data:image/x-emf;base64,${METAFILE}`, x: 1, y: 1, w: 1, h: 1 })
			p.addSlide().background = { data: `data:image/x-wmf;base64,${METAFILE}` }
		})
		const names = listEntries(zip).filter((name) => name.startsWith('ppt/media/'))
		assert(
			names.some((name) => name.endsWith('.emf')),
			`an emf part: ${names.join(', ')}`
		)
		assert(
			names.some((name) => name.endsWith('.wmf')),
			`a wmf part: ${names.join(', ')}`
		)
		assert(!names.some((name) => name.endsWith('.png')), `no png part: ${names.join(', ')}`)
		const types = await readEntry(zip, '[Content_Types].xml')
		assert(/Extension="emf" ContentType="image\/x-emf"/.test(types), 'emf is declared image/x-emf')
		assert(/Extension="wmf" ContentType="image\/x-wmf"/.test(types), 'wmf is declared image/x-wmf')
	})

	test('a cover or preview whose data has no base64 header warns and embeds the placeholder', async () => {
		const { result, codes } = await captureDiagnostics(() =>
			build((p) => {
				const host = p.addSlide()
				p.addSlide()
				host.addSlideZoom({ target: 2, x: 1, y: 1, w: 3, h: 2, coverImage: { data: 'hello world' } })
				host.addOleObject({
					data: 'UEsDBA==',
					progId: 'Excel.Sheet.12',
					cover: { data: 'nope' },
					x: 5,
					y: 1,
					w: 2,
					h: 2,
				})
				host.addModel3d({
					data: 'data:model/gltf-binary;base64,Z2xURg==',
					preview: { data: 'nope' },
					x: 1,
					y: 4,
					w: 2,
					h: 2,
				})
			})
		)
		assertEqual(
			codes.filter((code) => code === 'preview-image/missing-base64-header').length,
			3,
			`one warning per cover: ${codes.join(', ')}`
		)
		const pngs = (await mediaParts(result.zip)).filter(({ name }) => name.endsWith('.png'))
		assert(pngs.length > 0, 'the placeholder is embedded')
		for (const { name, bytes } of pngs)
			assert(bytes[0] === 0x89 && bytes[1] === 0x50, `${name} is a PNG, not the caller's text`)
	})

	test('a background whose data has no base64 header warns and paints no image', async () => {
		const { result, codes } = await captureDiagnostics(() =>
			build((p) => {
				p.addSlide().background = { data: 'not base64 at all' }
			})
		)
		assertIncludes(codes, 'background/missing-base64-header')
		assertEqual((await mediaParts(result.zip)).length, 0, 'no media part')
		const rels = await readEntry(result.zip, 'ppt/slides/_rels/slide1.xml.rels')
		assert(!rels.includes('../media/'), `no image relationship: ${rels}`)
	})

	test('a refused addImage or addText leaves no relationship, media part, name or object behind', async () => {
		// Both used to register first and validate last: the image's media part and relationship, and
		// the first run's link, stayed on the slide with nothing to use them.
		const conflicting = { url: 'https://refused.example', slide: 1 }
		/** @type {string[]} */
		const refusals = []
		const { zip } = await build((p) => {
			const slide = p.addSlide()
			const attempts = [
				() => slide.addImage({ data: PNG_1X1, x: 1, y: 1, w: 1, h: 1, hyperlink: conflicting }),
				() =>
					slide.addText(
						[
							{ text: 'first', options: { hyperlink: { url: 'https://first.example' } } },
							{ text: 'refused', options: { hyperlink: conflicting } },
						],
						{ x: 1, y: 3, w: 4, h: 1 }
					),
			]
			for (const attempt of attempts) {
				try {
					attempt()
					refusals.push('none')
				} catch (err) {
					refusals.push(/** @type {any} */ (err).code)
				}
			}
			slide.addImage({ data: PNG_1X1, x: 5, y: 1, w: 1, h: 1 })
		})
		assertEqual(refusals.join(' '), 'hyperlink/conflicting-targets hyperlink/conflicting-targets', 'both calls refused')
		const rels = await readEntry(zip, 'ppt/slides/_rels/slide1.xml.rels')
		assert(!rels.includes('.example'), `no hyperlink relationship survives: ${rels}`)
		assertEqual(
			(rels.match(/\.\.\/media\//g) ?? []).length,
			1,
			`only the added image has a media relationship: ${rels}`
		)
		const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
		assertEqual((xml.match(/<p:pic>/g) ?? []).length, 1, 'one picture')
		assertIncludes(xml, 'name="Image 1"', 'the refused image took no name')
	})
})
