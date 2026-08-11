// Phase 1 round-trip fidelity harness for `ts-pptx/read` (src/read/).
//
// Contract under test: OpcPackage.load(buf).save() preserves the part-name
// set and writes every untouched part body byte-identically; dirty parts
// reserialize from their DOM and stay schema-valid.

import { readFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { ContentTypes, OpcPackage, Relationships, resolveRelativePartName, relsPartNameFor } from '../../dist/read.js'
import { bytesEqual, assert, assertEqual, partBodies, assertUnchangedExcept } from '../helpers.js'
import { validatorAvailable, validateBuf } from '../validator.js'
import { fixtureNames, fixturePath } from './corpus.js'

const OFFICE_DOCUMENT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

const validatorInstalled = await validatorAvailable()

async function loadFixture(name) {
	return readFile(fixturePath(name))
}

/**
 * `{ input, saved }` for a fixture — its committed bytes, and what one load→save produces.
 *
 * Memoized because six contracts below each want the same pair and a save is deterministic
 * (which `idempotence` asserts separately, so this does not assume what it is checking).
 * Only *bytes* are cached, never an `OpcPackage`: `laziness` asserts no part was parsed, and
 * a package shared with a test that read a DOM would fail on the neighbour's access rather
 * than its own.
 *
 * @type {Map<string, Promise<{ input: Buffer, saved: Uint8Array }>>}
 */
const roundTripped = new Map()
function roundTrip(name) {
	let pending = roundTripped.get(name)
	if (!pending) {
		pending = (async () => {
			const input = await loadFixture(name)
			return { input, saved: await (await OpcPackage.load(input)).save() }
		})()
		roundTripped.set(name, pending)
	}
	return pending
}

// Every committed fixture, not a hand-picked five.
//
// This list used to be `['empty', 'textbox', 'image', 'table', 'mixed']`, and
// docs/testing.md told you to add to it by hand when promoting a deck into the corpus.
// Nobody did: the corpus grew to 44 and the OPC contract kept being proved against the
// same 5, so the decks that actually stress it — chartEx, model3d, math-omml, embedded
// fonts, av-media, modern comments — were carried without ever being round-tripped here.
// Reading `fixtureNames` makes promotion the only step there is.
//
// One case per fixture rather than one loop over all of them, for the reason
// script-ir.test.js gives: a loop stops at the first offender, so a change that breaks
// half the corpus reports the same single failure as one that breaks a single deck.
describe('OPC round-trip — corpus invariants', () => {
	test.for(fixtureNames)('%s: load → save keeps the same part-name set', async (name) => {
		const { input, saved } = await roundTrip(name)
		assertEqual(
			[...(await partBodies(saved)).keys()].sort().join('\n'),
			[...(await partBodies(input)).keys()].sort().join('\n'),
			`${name}: part-name set after round-trip`
		)
	})

	test.for(fixtureNames)('%s: every untouched part body is byte-identical', async (name) => {
		const { input, saved } = await roundTrip(name)
		assertUnchangedExcept(await partBodies(input), await partBodies(saved), [], name)
	})

	test.for(fixtureNames)('%s: no part is parsed as XML during load/save', async (name) => {
		const pkg = await OpcPackage.load(await loadFixture(name))
		await pkg.save()
		for (const part of pkg.parts.values()) {
			assert(!part.isParsed, `${name}: ${part.partName} was parsed without any DOM access`)
			assert(!part.isDirty, `${name}: ${part.partName} was marked dirty without any mutation`)
		}
	})

	test.for(fixtureNames)('%s: saving twice yields identical part bodies', async (name) => {
		const pkg = await OpcPackage.load(await loadFixture(name))
		const first = await partBodies(await pkg.save())
		const second = await partBodies(await pkg.save())
		assertEqual([...second.keys()].join('\n'), [...first.keys()].join('\n'), `${name}: partnames across saves`)
		assertUnchangedExcept(first, second, [], `${name}: between saves`)
	})

	test.for(fixtureNames)('%s: content types and relationships resolve', async (name) => {
		const pkg = await OpcPackage.load(await loadFixture(name))
		const slides = pkg.partsByContentType(SLIDE_CONTENT_TYPE)
		assert(slides.length >= 1, `${name}: expected at least one slide part`)
		assertEqual(pkg.contentTypes.contentTypeFor(slides[0].partName), SLIDE_CONTENT_TYPE, `${name}: slide Override`)

		const packageRels = pkg.relationshipsFor('/')
		const officeDocument = packageRels.byType(OFFICE_DOCUMENT_REL)
		assertEqual(officeDocument.length, 1, `${name}: officeDocument relationship count`)
		assertEqual(
			packageRels.resolveTarget(officeDocument[0].id),
			'/ppt/presentation.xml',
			`${name}: officeDocument target`
		)

		const presentationRels = pkg.relationshipsFor('/ppt/presentation.xml')
		assert(presentationRels.size > 0, `${name}: presentation part should have relationships`)
		for (const relationship of presentationRels) {
			if (relationship.targetMode === 'External') continue
			const target = presentationRels.resolveTarget(relationship.id)
			assert(pkg.part(target), `${name}: relationship ${relationship.id} target ${target} is not a part`)
		}
	})
})

/** A stable, order-independent identity for a validator verdict. */
function errorFingerprint(errors) {
	return errors
		.map((e) => `${e.Id} ${e.Path?.PartUri ?? ''} ${e.Path?.XPath ?? ''}`)
		.sort()
		.join('\n')
}

// The claim is "a round-trip introduces no NEW errors", not "the output is clean".
//
// Those were the same assertion while this ran against five hand-picked decks that all
// happened to validate. Across the whole corpus they are not: `bar-chart-data-labels.pptx`
// carries three Microsoft365 errors *as committed*, in PowerPoint's own chart `c:extLst` —
// an undeclared `uri` on `c:ext`, a `chart:dataDisplayOptions16` where the SDK's schema
// models only `dispNaAsBlank`, and a 2012-namespace `chart:leaderLines` under `c:dLbls`.
// The SDK does not model those extension namespaces; PowerPoint wrote them anyway. Nothing
// of ours produced them and nothing of ours can fix them.
//
// So validate both sides and compare. That is strictly stronger than the old assertion for
// the 43 clean decks — an empty verdict before still demands an empty verdict after — and
// it is the only form that says anything true about the 44th. Asserting "clean" would have
// forced the choice between excluding that fixture and pretending the library caused it.
//
// Concurrent, and its own block for that reason: the validator batches whatever is in
// flight (test/validator.js), so 88 sequential validations would pay the ~0.4s .NET startup
// 88 times where concurrent ones go out in a handful of invocations. Nothing here touches a
// process global, which is what makes that safe.
describe.concurrent('OPC round-trip — schema validity', () => {
	test.skipIf(!validatorInstalled).for(fixtureNames)('%s: a round-trip introduces no new errors', async (name) => {
		const { input, saved } = await roundTrip(name)
		const [before, after] = await Promise.all([validateBuf(Buffer.from(input)), validateBuf(Buffer.from(saved))])
		assertEqual(
			errorFingerprint(after),
			errorFingerprint(before),
			`${name}: the saved package's validator verdict differs from the committed fixture's. ` +
				`after: ${JSON.stringify(after).slice(0, 1500)}`
		)
	})
})

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
		assertUnchangedExcept(inputBodies, outputBodies, [dirtyEntry])
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

	test('load drops PowerPoint [trash] parts instead of failing on their missing content type', async () => {
		const zip = await JSZip.loadAsync(await loadFixture('empty'))
		// PowerPoint parks deleted parts under /[trash]/ without registering them
		// in [Content_Types].xml. Such a part has no resolvable content type, so it
		// would trip the rejection above if it were not skipped on load.
		zip.file('[trash]/3681', 'inert deleted-part bytes')
		zip.file('[trash]/ppt/slides/slide99.xml', '<p:sld/>')
		const withTrash = await zip.generateAsync({ type: 'uint8array' })

		const pkg = await OpcPackage.load(withTrash)
		assertEqual(pkg.part('/[trash]/3681'), undefined, 'top-level trash part is not loaded')
		assertEqual(pkg.part('/[trash]/ppt/slides/slide99.xml'), undefined, 'nested trash part is not loaded')
		assert(
			[...pkg.parts.keys()].every((name) => !name.startsWith('/[trash]/')),
			`no [trash] parts survive load; got: ${[...pkg.parts.keys()].join(', ')}`
		)
		// Live parts are unaffected: the presentation part still loads.
		assert(pkg.part('/ppt/presentation.xml'), 'live presentation part still loaded alongside trash')
	})
})
