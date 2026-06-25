// Feature A (import-carry) for embedded fonts: importSlide(source, i, { embedFonts: true })
// brings the source deck's presentation-level embedded fonts across — the binary
// `.fntdata` parts, the `application/x-fontdata` content-type Default, the font
// relationships, and a merged `p:embeddedFontLst` — while the default (flag off)
// leaves the deck unchanged. Oracle: test/read/fixtures/embedded-fonts.pptx
// (PowerPoint-authored) + embedded-fonts.oracle.json.

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

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}
async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}
async function entries(pptxBytes) {
	const zip = await JSZip.loadAsync(pptxBytes)
	return zip
}

describe('Presentation.importSlide({ embedFonts })', () => {
	test('carries font parts, content-type Default, rels, and a merged embeddedFontLst', async () => {
		const target = await open('empty')
		const source = await open('embedded-fonts')
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
		const target = await open('empty')
		const source = await open('embedded-fonts')
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
		const target = await open('empty')
		const source = await open('embedded-fonts')
		target.importSlide(source, 0)

		const zip = await entries(await target.save())
		assert(!Object.keys(zip.files).some((n) => /fntdata/.test(n)), 'no font parts without embedFonts')
		const pres = await zip.file('ppt/presentation.xml').async('string')
		assert(!/embeddedFontLst/.test(pres), 'no embeddedFontLst without embedFonts')
	})

	test.skipIf(!validatorInstalled)('a deck with carried embedded fonts stays schema-valid', async () => {
		const target = await open('empty')
		const source = await open('embedded-fonts')
		target.importSlide(source, 0, { embedFonts: true })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
