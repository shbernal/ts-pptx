// Edit tests for shape fill / line colour setters (src/read/api/shapes/base.ts).
//
// Contract under test: setting fillColor / fillSchemeColor / noFill() / lineColor
// on a Shape mutates the shape's p:spPr (or p:grpSpPr) in document order, marks
// only the owning slide part dirty, survives a save → reopen round-trip, leaves
// untouched parts byte-identical, and keeps the package schema-valid. Kinds that
// have no own fill model (picture fill, graphicFrame) reject the setter.

import { readFile } from 'node:fs/promises'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { throws, bytesEqual, assert, assertEqual, partBodies, assertUnchangedExcept } from '../helpers.js'
import { validateBuf, validatorInstalled } from '../validator.js'
import { fixturePath, openFixture } from './corpus.js'

async function editAndReopen(name, edit) {
	const presentation = await openFixture(name)
	await edit(presentation)
	const saved = await presentation.save()
	return { presentation, saved, reopened: await Presentation.load(saved) }
}

function replaceTextShape(presentation) {
	return presentation.slides[0].shapes.find((shape) => shape.name === 'replaceText')
}

function findByKind(presentation, shapeType) {
	return presentation.slides.flatMap((slide) => slide.shapes).find((shape) => shape.shapeType === shapeType)
}

/** Slide XML restricted to the p:spPr of the named shape, for ordering assertions. */
function spPrXml(slideXml, name) {
	const nameIdx = slideXml.indexOf(`name="${name}"`)
	const start = slideXml.indexOf('<p:spPr>', nameIdx)
	const end = slideXml.indexOf('</p:spPr>', start)
	return slideXml.slice(start, end)
}

describe('Shape fill editing', () => {
	test('fillColor setter replaces the prior fill and reloads', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			const shape = replaceTextShape(presentation)
			assertEqual(shape.fillColor, null, 'precondition: replaceText starts with a:noFill, no srgb fill')
			shape.fillColor = '#FF0000'
		})
		const shape = replaceTextShape(reopened)
		assertEqual(shape.fillColor, 'FF0000', 'explicit srgb fill reloads (hash + case normalized)')
		assertEqual(shape.fillSchemeColor, null, 'no scheme fill remains')
	})

	test('fillSchemeColor setter writes a theme token and reloads', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			replaceTextShape(presentation).fillSchemeColor = 'accent2'
		})
		const shape = replaceTextShape(reopened)
		assertEqual(shape.fillSchemeColor, 'accent2', 'scheme fill reloads')
		assertEqual(shape.fillColor, null, 'no explicit srgb fill remains')
	})

	test('fillSchemeColor and lineSchemeColor refuse a token outside the schema, writing nothing', async () => {
		const presentation = await openFixture('textbox')
		const shape = replaceTextShape(presentation)
		const before = await presentation.save()
		const refusals = [
			() => {
				shape.fillSchemeColor = 'notAToken'
			},
			() => {
				shape.lineSchemeColor = 'notAToken'
			},
		]
		for (const refuse of refusals) {
			let code = null
			try {
				refuse()
			} catch (err) {
				code = err.code
			}
			assertEqual(code, 'color/invalid-scheme-token', 'the token is refused')
		}
		assert(bytesEqual(before, await presentation.save()), 'and nothing reached the part')
	})

	test('clearing fillColor (= null) removes the solidFill, restoring inheritance', async () => {
		const { saved, reopened } = await editAndReopen('textbox', (presentation) => {
			const shape = replaceTextShape(presentation)
			shape.fillColor = '112233'
			shape.fillColor = null
		})
		assertEqual(replaceTextShape(reopened).fillColor, null, 'fill is cleared after reload')
		const slideXml = new TextDecoder().decode((await partBodies(saved)).get('ppt/slides/slide1.xml'))
		assert(!spPrXml(slideXml, 'replaceText').includes('<a:solidFill>'), 'no a:solidFill left in the shape')
	})

	test('noFill() emits an explicit <a:noFill/>, distinct from clearing', async () => {
		const { saved, reopened } = await editAndReopen('textbox', (presentation) => {
			const shape = replaceTextShape(presentation)
			shape.fillColor = 'AABBCC'
			shape.noFill()
		})
		const shape = replaceTextShape(reopened)
		assertEqual(shape.fillColor, null, 'no srgb fill after noFill()')
		const slideXml = new TextDecoder().decode((await partBodies(saved)).get('ppt/slides/slide1.xml'))
		assert(spPrXml(slideXml, 'replaceText').includes('<a:noFill/>'), 'explicit a:noFill emitted')
	})

	test('inserts a:solidFill after a:prstGeom and before a:ln', async () => {
		const { saved } = await editAndReopen('textbox', (presentation) => {
			const shape = replaceTextShape(presentation)
			shape.lineColor = '00FF00' // creates a:ln (a later sibling than the fill)
			shape.fillColor = 'FF0000'
		})
		const xml = spPrXml(new TextDecoder().decode((await partBodies(saved)).get('ppt/slides/slide1.xml')), 'replaceText')
		const geom = xml.indexOf('<a:prstGeom')
		const fill = xml.indexOf('<a:solidFill>')
		const line = xml.indexOf('<a:ln>')
		assert(geom >= 0 && fill >= 0 && line >= 0, 'prstGeom, solidFill, and ln are all present')
		assert(geom < fill && fill < line, `document order is prstGeom < solidFill < ln (got ${geom}, ${fill}, ${line})`)
	})

	test('rejects a malformed hex colour', async () => {
		const shape = replaceTextShape(await openFixture('textbox'))
		assert(
			throws(() => (shape.fillColor = 'nothex')),
			'malformed fill colour should throw'
		)
	})
})

describe('Shape line editing', () => {
	test('lineColor and lineSchemeColor round-trip on an auto shape', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			replaceTextShape(presentation).lineColor = '#1A2B3C'
		})
		assertEqual(replaceTextShape(reopened).lineColor, '1A2B3C', 'line srgb colour reloads')

		const scheme = await editAndReopen('textbox', (presentation) => {
			replaceTextShape(presentation).lineSchemeColor = 'accent4'
		})
		const shape = replaceTextShape(scheme.reopened)
		assertEqual(shape.lineSchemeColor, 'accent4', 'line scheme colour reloads')
		assertEqual(shape.lineColor, null, 'no explicit srgb line colour remains')
	})

	test('clearing lineColor removes only the line solidFill', async () => {
		const { reopened } = await editAndReopen('textbox', (presentation) => {
			const shape = replaceTextShape(presentation)
			shape.lineColor = '445566'
			shape.lineColor = null
		})
		assertEqual(replaceTextShape(reopened).lineColor, null, 'line colour cleared after reload')
	})

	test('sets lineColor on a connector', async () => {
		const { reopened } = await editAndReopen('mixed', (presentation) => {
			findByKind(presentation, 'connector').lineColor = 'FF8800'
		})
		assertEqual(findByKind(reopened, 'connector').lineColor, 'FF8800', 'connector line colour reloads')
	})
})

describe('Per-kind fill / line support', () => {
	test('group shape fill writes p:grpSpPr/a:solidFill and reloads', async () => {
		const { reopened } = await editAndReopen('mixed', (presentation) => {
			findByKind(presentation, 'group').fillColor = '123456'
		})
		assertEqual(findByKind(reopened, 'group').fillColor, '123456', 'group fill reloads')
	})

	test('group shape rejects a line colour (grpSpPr has no a:ln)', async () => {
		const group = findByKind(await openFixture('mixed'), 'group')
		assert(
			throws(() => (group.lineColor = '000000')),
			'group lineColor should throw'
		)
	})

	test('picture rejects fill but accepts a border (lineColor)', async () => {
		const picture = findByKind(await openFixture('image'), 'picture')
		assert(
			throws(() => (picture.fillColor = '000000')),
			'picture fillColor should throw'
		)
		assert(
			throws(() => picture.noFill()),
			'picture noFill() should throw'
		)
		const { reopened } = await editAndReopen('image', (presentation) => {
			findByKind(presentation, 'picture').lineColor = '00AAFF'
		})
		assertEqual(findByKind(reopened, 'picture').lineColor, '00AAFF', 'picture border colour reloads')
	})

	test('each kind refuses a paint it cannot carry with that paint’s code, and clearing never throws', async () => {
		// A frame threw a properties code for both paints, a picture a fill code, a group a line code:
		// three codes for one condition.
		const mixed = await openFixture('mixed')
		const frame = findByKind(mixed, 'graphicFrame')
		const group = findByKind(mixed, 'group')
		const picture = findByKind(await openFixture('image'), 'picture')
		const codeOf = (/** @type {() => unknown} */ fn) => {
			try {
				fn()
				return null
			} catch (err) {
				return err.code
			}
		}
		assertEqual(
			codeOf(() => (frame.fillColor = '000000')),
			'shape/fill-unsupported',
			'a frame fill'
		)
		assertEqual(
			codeOf(() => frame.noFill()),
			'shape/fill-unsupported',
			'a frame no-fill'
		)
		assertEqual(
			codeOf(() => (frame.lineSchemeColor = 'accent1')),
			'shape/line-unsupported',
			'a frame line'
		)
		assertEqual(
			codeOf(() => (picture.fillSchemeColor = 'accent1')),
			'shape/fill-unsupported',
			'a picture fill'
		)
		assertEqual(
			codeOf(() => (group.lineColor = '000000')),
			'shape/line-unsupported',
			'a group line'
		)
		for (const clear of [
			() => (frame.fillColor = null),
			() => (frame.lineColor = null),
			() => (picture.fillColor = null),
			() => (group.lineSchemeColor = null),
		]) {
			assertEqual(codeOf(clear), null, 'clearing never throws')
		}
	})

	test('a fill a deck wrote on a picture still clears', async () => {
		const presentation = await openFixture('image')
		const picture = findByKind(presentation, 'picture')
		const spPr = picture.element_.getElementsByTagNameNS(
			'http://schemas.openxmlformats.org/presentationml/2006/main',
			'spPr'
		)[0]
		const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
		const fill = spPr.ownerDocument.createElementNS(A, 'a:solidFill')
		const clr = spPr.ownerDocument.createElementNS(A, 'a:srgbClr')
		clr.setAttribute('val', '123456')
		fill.appendChild(clr)
		spPr.appendChild(fill)
		assertEqual(picture.fillColor, '123456', 'the planted fill reads back')
		picture.fillColor = null
		assertEqual(picture.fillColor, null, 'and clears')
	})

	test('graphic frame rejects both fill and line colours', async () => {
		const frame = findByKind(await openFixture('mixed'), 'graphicFrame')
		assert(
			throws(() => (frame.fillColor = '000000')),
			'graphicFrame fillColor should throw'
		)
		assert(
			throws(() => (frame.lineColor = '000000')),
			'graphicFrame lineColor should throw'
		)
	})
})

describe('fidelity and schema validity', () => {
	test('a fill edit leaves every other part byte-identical', async () => {
		const input = await readFile(fixturePath('textbox'))
		const presentation = await Presentation.load(input)
		replaceTextShape(presentation).fillColor = 'FF0000'
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(await presentation.save())
		const dirty = 'ppt/slides/slide1.xml'
		assert(!bytesEqual(inputBodies.get(dirty), outputBodies.get(dirty)), 'edited slide differs')
		assertUnchangedExcept(inputBodies, outputBodies, [dirty])
	})

	test.skipIf(!validatorInstalled)('fill + line + noFill edits stay schema-valid', async () => {
		const { saved } = await editAndReopen('mixed', (presentation) => {
			const auto = findByKind(presentation, 'autoShape')
			auto.fillColor = '1A2B3C'
			auto.lineColor = 'D4D4D4'
			findByKind(presentation, 'group').fillSchemeColor = 'accent3'
			findByKind(presentation, 'connector').lineColor = '003366'
		})
		const errors = await validateBuf(Buffer.from(saved))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
