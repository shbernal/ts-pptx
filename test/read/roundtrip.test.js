// Phase 1 round-trip fidelity harness for `pptxgenjs/read` (src/read/).
//
// Contract under test: OpcPackage.load(buf).save() preserves the part-name
// set and writes every untouched part body byte-identically; dirty parts
// reserialize from their DOM and stay schema-valid.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { ContentTypes, OpcPackage, Relationships, resolveRelativePartName, relsPartNameFor } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const FIXTURES = ['empty', 'textbox', 'image', 'table', 'mixed']

const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

const validatorInstalled = await isInstalled()

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function loadFixture(name) {
	return readFile(fixturePath(name))
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
	return a.length === b.length && a.every((value, index) => value === b[index])
}

for (const name of FIXTURES) {
	describe(`round-trip: ${name}.pptx`, () => {
		test('part-set stability: load → save → reload keeps the same partnames', async () => {
			const input = await loadFixture(name)
			const pkg = await OpcPackage.load(input)
			const saved = await pkg.save()
			const inputBodies = await partBodies(input)
			const outputBodies = await partBodies(saved)
			assertEqual(
				[...outputBodies.keys()].sort().join('\n'),
				[...inputBodies.keys()].sort().join('\n'),
				`${name}: part-name set after round-trip`
			)
		})

		test('byte-identity: every untouched part body is identical to the input', async () => {
			const input = await loadFixture(name)
			const saved = await (await OpcPackage.load(input)).save()
			const inputBodies = await partBodies(input)
			const outputBodies = await partBodies(saved)
			for (const [entryName, inputBody] of inputBodies) {
				const outputBody = outputBodies.get(entryName)
				assert(outputBody, `${name}: ${entryName} missing from output`)
				assert(bytesEqual(inputBody, outputBody), `${name}: ${entryName} body differs after round-trip`)
			}
		})

		test('laziness: no part is parsed as XML during load/save', async () => {
			const pkg = await OpcPackage.load(await loadFixture(name))
			await pkg.save()
			for (const part of pkg.parts.values()) {
				assert(!part.isParsed, `${part.partName} was parsed without any DOM access`)
				assert(!part.isDirty, `${part.partName} was marked dirty without any mutation`)
			}
		})

		test('idempotence: saving twice yields identical part bodies', async () => {
			const pkg = await OpcPackage.load(await loadFixture(name))
			const first = await partBodies(await pkg.save())
			const second = await partBodies(await pkg.save())
			assertEqual([...second.keys()].join('\n'), [...first.keys()].join('\n'), `${name}: partnames across saves`)
			for (const [entryName, firstBody] of first) {
				assert(bytesEqual(firstBody, second.get(entryName)), `${name}: ${entryName} differs between saves`)
			}
		})

		test('content-type and relationship resolution', async () => {
			const pkg = await OpcPackage.load(await loadFixture(name))
			const slides = pkg.partsByContentType(SLIDE_CONTENT_TYPE)
			assert(slides.length >= 1, `${name}: expected at least one slide part`)
			assertEqual(pkg.contentTypes.contentTypeFor(slides[0].partName), SLIDE_CONTENT_TYPE, 'slide Override')

			const packageRels = pkg.relationshipsFor('/')
			const officeDocument = packageRels.byType(OFFICE_DOCUMENT_REL)
			assertEqual(officeDocument.length, 1, `${name}: officeDocument relationship count`)
			assertEqual(packageRels.resolveTarget(officeDocument[0].id), '/ppt/presentation.xml', 'officeDocument target')

			const presentationRels = pkg.relationshipsFor('/ppt/presentation.xml')
			assert(presentationRels.size > 0, `${name}: presentation part should have relationships`)
			for (const relationship of presentationRels) {
				if (relationship.targetMode === 'External') continue
				const target = presentationRels.resolveTarget(relationship.id)
				assert(pkg.part(target), `${name}: relationship ${relationship.id} target ${target} is not a part`)
			}
		})

		test.skipIf(!validatorInstalled)('schema validity: saved output passes the OOXML validator', async () => {
			const saved = await (await OpcPackage.load(await loadFixture(name))).save()
			const errors = await validateBuf(Buffer.from(saved))
			assertEqual(errors.length, 0, `${name}: validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
		})
	})
}

describe('dirty path: mutate one slide, save', () => {
	async function mutateFirstTextRun() {
		const input = await loadFixture('textbox')
		const pkg = await OpcPackage.load(input)
		const slide = pkg.partsByContentType(SLIDE_CONTENT_TYPE)[0]
		const textNode = slide.dom.getElementsByTagName('a:t')[0]
		assert(textNode, 'textbox slide should contain an <a:t> run')
		textNode.textContent = 'EDITED BY ROUNDTRIP TEST'
		slide.markDirty()
		return { input, pkg, slide }
	}

	test('dirty part body changes, untouched parts stay byte-identical', async () => {
		const { input, pkg, slide } = await mutateFirstTextRun()
		const saved = await pkg.save()
		const inputBodies = await partBodies(input)
		const outputBodies = await partBodies(saved)
		const dirtyEntry = slide.partName.slice(1)
		assert(!bytesEqual(inputBodies.get(dirtyEntry), outputBodies.get(dirtyEntry)), 'dirty part body should differ')
		for (const [entryName, inputBody] of inputBodies) {
			if (entryName === dirtyEntry) continue
			assert(bytesEqual(inputBody, outputBodies.get(entryName)), `${entryName} should be untouched`)
		}
	})

	test('the edit survives a reload', async () => {
		const { pkg, slide } = await mutateFirstTextRun()
		const reloaded = await OpcPackage.load(await pkg.save())
		const xml = new TextDecoder().decode(reloaded.part(slide.partName).bytes)
		assert(xml.includes('EDITED BY ROUNDTRIP TEST'), 'mutated text should be present after reload')
		assert(xml.startsWith('<?xml'), 'dirty part should keep an XML declaration')
	})

	test.skipIf(!validatorInstalled)('mutated package is still schema-valid', async () => {
		const { pkg } = await mutateFirstTextRun()
		const errors = await validateBuf(Buffer.from(await pkg.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

describe('partname and overlay units', () => {
	test('resolveRelativePartName handles relative, parent, and absolute targets', () => {
		assertEqual(resolveRelativePartName('/', 'ppt/presentation.xml'), '/ppt/presentation.xml')
		assertEqual(resolveRelativePartName('/ppt/presentation.xml', 'slides/slide1.xml'), '/ppt/slides/slide1.xml')
		assertEqual(resolveRelativePartName('/ppt/slides/slide1.xml', '../media/image1.png'), '/ppt/media/image1.png')
		assertEqual(
			resolveRelativePartName('/ppt/slides/slide1.xml', '/docProps/thumbnail.jpeg'),
			'/docProps/thumbnail.jpeg'
		)
	})

	test('relsPartNameFor maps package root and nested parts', () => {
		assertEqual(relsPartNameFor('/'), '/_rels/.rels')
		assertEqual(relsPartNameFor('/ppt/presentation.xml'), '/ppt/_rels/presentation.xml.rels')
		assertEqual(relsPartNameFor('/ppt/slides/slide1.xml'), '/ppt/slides/_rels/slide1.xml.rels')
	})

	test('ContentTypes resolves Override before Default and round-trips', () => {
		const xml =
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
			'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
			'<Default Extension="png" ContentType="image/png"/>' +
			'<Default Extension="xml" ContentType="application/xml"/>' +
			'<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
			'</Types>'
		const contentTypes = ContentTypes.parse(xml)
		assertEqual(
			contentTypes.contentTypeFor('/ppt/slides/slide1.xml'),
			'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
			'Override wins'
		)
		assertEqual(contentTypes.contentTypeFor('/ppt/media/IMAGE1.PNG'), 'image/png', 'Default by lowercased extension')
		assertEqual(contentTypes.contentTypeFor('/ppt/slides/slide2.xml'), 'application/xml', 'Default fallback for xml')
		assertEqual(contentTypes.contentTypeFor('/ppt/media/movie.mp4'), undefined, 'unknown extension')
		assertEqual(
			ContentTypes.parse(contentTypes.serialize()).contentTypeFor('/ppt/media/a.png'),
			'image/png',
			'serialize round-trips'
		)
	})

	test('Relationships parses targets, modes, and external rels', () => {
		const xml =
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
			'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>' +
			'</Relationships>'
		const relationships = Relationships.parse(xml, '/ppt/slides/slide1.xml')
		assertEqual(relationships.resolveTarget('rId1'), '/ppt/slideLayouts/slideLayout1.xml', 'relative target')
		assertEqual(relationships.get('rId2').targetMode, 'External', 'external mode')
		let threw = false
		try {
			relationships.resolveTarget('rId2')
		} catch {
			threw = true
		}
		assert(threw, 'resolveTarget on an External rel should throw')
		const reparsed = Relationships.parse(relationships.serialize(), '/ppt/slides/slide1.xml')
		assertEqual(reparsed.resolveTarget('rId1'), '/ppt/slideLayouts/slideLayout1.xml', 'serialize round-trips')
	})

	test('binary parts refuse DOM access but serialize their original bytes', async () => {
		const pkg = await OpcPackage.load(await loadFixture('image'))
		const media = [...pkg.parts.values()].find((part) => part.contentType === 'image/png')
		assert(media, 'image fixture should contain a png part')
		assert(!media.isXmlPart, 'png part is not an XML part')
		let threw = false
		try {
			void media.dom
		} catch {
			threw = true
		}
		assert(threw, 'dom access on a binary part should throw')
		assert(bytesEqual(media.serialize(), media.bytes), 'binary serialize returns original bytes')
	})

	test('load rejects a part with no resolvable content type', async () => {
		const zip = await JSZip.loadAsync(await loadFixture('empty'))
		zip.file('ppt/media/orphan.zzz', 'not a known type')
		const broken = await zip.generateAsync({ type: 'uint8array' })
		let message = ''
		try {
			await OpcPackage.load(broken)
		} catch (error) {
			message = String(error)
		}
		assert(message.includes('/ppt/media/orphan.zzz'), `load should name the offending part; got: ${message}`)
	})
})
