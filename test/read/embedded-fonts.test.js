// Feature A (import-carry) for embedded fonts: importSlide(source, i, { embedFonts: true })
// brings the source deck's presentation-level embedded fonts across — the binary
// `.fntdata` parts, the `application/x-fontdata` content-type Default, the font
// relationships, and a merged `p:embeddedFontLst` — while the default (flag off)
// leaves the deck unchanged. Oracle: test/read/fixtures/embedded-fonts.pptx
// (PowerPoint-authored) + embedded-fonts.oracle.json.

import JSZip from 'jszip'
import { describe, test } from 'vitest'

import { Presentation } from '../../dist/read.js'
import { assert, assertEqual, bytesEqual } from '../helpers.js'
import { validateBuf, validatorInstalled } from '../validator.js'
import { openFixture, readFixture } from './corpus.js'

async function entries(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	return zip
}

describe('Presentation.importSlide({ embedFonts })', () => {
	test('carries font parts, content-type Default, rels, and a merged embeddedFontLst', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('embedded-fonts')
		target.importSlide(source, 0, { embedFonts: true })

		const zip = await entries(await target.save())
		const names = Object.keys(zip.files)

		const fontParts = names.filter((n) => /^ppt\/fonts\/font\d+\.fntdata$/.test(n)).sort()
		assertEqual(fontParts.length, 2, `two font parts carried (got ${JSON.stringify(fontParts)})`)

		const ct = await zip.file('[Content_Types].xml').async('string')
		assert(/<Default Extension="fntdata" ContentType="application\/x-fontdata"\/>/.test(ct), 'fntdata Default added')
		// One Default, no per-part Override (ensureDefault ran before the part was copied).
		assertEqual((ct.match(/x-fontdata/g) || []).length, 1, 'content type registered once (Default only)')

		const rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
		const fontRels = [...rels.matchAll(/<Relationship[^>]*\/relationships\/font"[^>]*\/>/g)].map((m) => m[0])
		assertEqual(fontRels.length, 2, 'two font relationships')
		assert(
			/Target="fonts\/font1\.fntdata"/.test(rels) && /Target="fonts\/font2\.fntdata"/.test(rels),
			'font rels target the carried parts'
		)

		const pres = await zip.file('ppt/presentation.xml').async('string')
		const lst = pres.match(/<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/)?.[0]
		assert(lst, 'embeddedFontLst present')
		// Typeface identity (typeface + pitchFamily + charset) is cloned from the source p:font.
		assert(
			/<p:font typeface="Silkscreen" pitchFamily="2" charset="0"\/>/.test(lst),
			`p:font identity carried; got ${lst}`
		)
		assert(
			/<p:regular r:id="[^"]+"\/>/.test(lst) && /<p:bold r:id="[^"]+"\/>/.test(lst),
			'regular + bold faces carried'
		)
		// embeddedFontLst sits before defaultTextStyle (CT_Presentation index 7).
		assert(
			pres.indexOf('<p:embeddedFontLst>') < pres.indexOf('<p:defaultTextStyle'),
			'embeddedFontLst precedes defaultTextStyle'
		)
	})

	test('is idempotent: importing the same slide twice carries each face once', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('embedded-fonts')
		target.importSlide(source, 0, { embedFonts: true })
		target.importSlide(source, 0, { embedFonts: true })

		const zip = await entries(await target.save())
		const fontParts = Object.keys(zip.files).filter((n) => /^ppt\/fonts\/font\d+\.fntdata$/.test(n))
		assertEqual(fontParts.length, 2, 'each face copied exactly once across repeated imports')

		const pres = await zip.file('ppt/presentation.xml').async('string')
		assertEqual(
			(pres.match(/<p:embeddedFont>/g) || []).length,
			1,
			'a single embeddedFont entry for the shared typeface'
		)
		assertEqual((pres.match(/<p:regular /g) || []).length, 1, 'regular face not duplicated')
		assertEqual((pres.match(/<p:bold /g) || []).length, 1, 'bold face not duplicated')
	})

	test('default (flag off) carries no fonts — unchanged behaviour', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('embedded-fonts')
		target.importSlide(source, 0)

		const zip = await entries(await target.save())
		assert(!Object.keys(zip.files).some((n) => /fntdata/.test(n)), 'no font parts without embedFonts')
		const pres = await zip.file('ppt/presentation.xml').async('string')
		assert(!/embeddedFontLst/.test(pres), 'no embeddedFontLst without embedFonts')
	})

	test.skipIf(!validatorInstalled)('a deck with carried embedded fonts stays schema-valid', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('embedded-fonts')
		target.importSlide(source, 0, { embedFonts: true })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})

// One embedded-font list, read one way. The getter skipped a face whose `r:id` names no
// relationship, while the import check and the carry resolved it and threw an option error,
// `relationship/not-found`. `importSlides` threw it from its dry run, but `importSlide` and
// `importSlideMasters` carry fonts last and threw it after the slide or the masters were already
// in the deck.
describe('an embedded font face whose r:id names no relationship', () => {
	/** `embedded-fonts.pptx` with the relationship behind its regular face removed. */
	async function danglingFaceSource() {
		const zip = await JSZip.loadAsync(await readFixture('embedded-fonts.pptx'))
		const presentation = await zip.file('ppt/presentation.xml').async('string')
		const relId = /<p:regular r:id="([^"]+)"/.exec(presentation)?.[1]
		assert(relId, 'the fixture embeds a regular face')
		const relsPath = 'ppt/_rels/presentation.xml.rels'
		const rels = await zip.file(relsPath).async('string')
		const pruned = rels.replace(new RegExp(`<Relationship [^>]*Id="${relId}"[^>]*/>`), '')
		assert(pruned !== rels, 'the relationship behind the regular face was removed')
		zip.file(relsPath, pruned)
		return Presentation.load(await zip.generateAsync({ type: 'uint8array' }))
	}

	test('the getter lists only the faces that resolve', async () => {
		const [font] = (await danglingFaceSource()).embeddedFonts
		assertEqual(font.typeface, 'Silkscreen', 'the typeface is still read')
		assertEqual(font.faces.map((face) => face.slot).join(','), 'bold', 'the dangling regular face is skipped')
	})

	/** @type {Record<string, (target: any, source: any) => unknown>} */
	const imports = {
		importSlide: (target, source) => target.importSlide(source, 0, { embedFonts: true }),
		importSlides: (target, source) =>
			target.importSlides([{ source, sourceIndex: 0, outputIndex: 0, embedFonts: true }]),
		importSlideMasters: (target, source) => target.importSlideMasters(source, { embedFonts: true }),
	}

	test.for(Object.keys(imports))('%s refuses it as a package error and leaves the deck unchanged', async (name) => {
		const source = await danglingFaceSource()
		const target = await openFixture('empty')
		const before = await target.save()
		/** @type {any} */
		let thrown = null
		try {
			imports[name](target, source)
		} catch (error) {
			thrown = error
		}
		assert(thrown, `${name} throws`)
		assertEqual(thrown.code, 'package/part-missing', `${name} names the missing font part`)
		assertEqual(thrown.name, 'PackageReadError', `${name} raises it as a package error`)
		assert(bytesEqual(before, await target.save()), `${name} changed no byte of the deck`)
	})
})
