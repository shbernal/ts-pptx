// Embedded-font carry for appendSlides: a generator's presentation-level embedded
// fonts (pptx.embedFont) are carried into the destination deck and merged into its
// p:embeddedFontLst — the author-onto-template counterpart of the importSlide carry
// in embedded-fonts.test.js. Asserts the font parts, the application/x-fontdata
// content-type Default, the presentation font rels, and the merged embeddedFontLst,
// plus de-dupe across repeated appends and schema validity. The generator side
// (pptx.embedFont emit) is covered in test/regression/embed-font.test.js.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test, beforeAll } from 'vitest'
import PptxGenJS from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { isInstalled, validateBuf } from '../validator.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const fontsDir = path.join(__dirname, 'fixtures', 'fonts')
const validatorInstalled = await isInstalled()

let regular
let bold

beforeAll(async () => {
	regular = new Uint8Array(await readFile(path.join(fontsDir, 'Silkscreen-Regular.ttf')))
	bold = new Uint8Array(await readFile(path.join(fontsDir, 'Silkscreen-Bold.ttf')))
})

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

/** Load the LAYOUT_WIDE-sized theme-colors deck (has a "Blank" layout to bind to). */
async function openTarget() {
	return Presentation.load(await readFile(fixturePath('theme-colors')))
}

/** A generator deck sized to LAYOUT_WIDE (matches theme-colors), embedding Silkscreen. */
async function fontGenerator() {
	const pptx = new PptxGenJS()
	pptx.layout = 'LAYOUT_WIDE'
	await pptx.embedFont({ data: regular, typeface: 'Silkscreen' })
	await pptx.embedFont({ data: bold, typeface: 'Silkscreen', style: 'bold' })
	pptx.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Silkscreen' })
	return pptx
}

async function zipOf(pptxBytes) {
	return JSZip.loadAsync(pptxBytes)
}

describe('Presentation.appendSlides — embedded fonts', () => {
	test('carries font parts, content-type Default, rels, and a merged embeddedFontLst', async () => {
		const target = await openTarget()
		await target.appendSlides(await fontGenerator(), { layout: 'Blank' })

		const zip = await zipOf(await target.save())
		const names = Object.keys(zip.files)

		const fontParts = names.filter((n) => /^ppt\/fonts\/font\d+\.fntdata$/.test(n)).sort()
		assertEqual(fontParts.length, 2, `two font parts carried (got ${JSON.stringify(fontParts)})`)

		const ct = await zip.file('[Content_Types].xml').async('string')
		assert(/<Default Extension="fntdata" ContentType="application\/x-fontdata"\/>/.test(ct), 'fntdata Default added')
		// One Default, no per-part Override (ensureDefault ran before the part was created).
		assertEqual((ct.match(/x-fontdata/g) || []).length, 1, 'content type registered once (Default only)')

		const rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
		const fontRels = [...rels.matchAll(/<Relationship[^>]*\/relationships\/font"[^>]*\/>/g)].map((m) => m[0])
		assertEqual(fontRels.length, 2, 'two font relationships')
		assert(
			fontRels.every((r) => /Target="fonts\/font\d+\.fntdata"/.test(r)),
			'font rels target the carried parts'
		)

		const pres = await zip.file('ppt/presentation.xml').async('string')
		const lst = pres.match(/<p:embeddedFontLst>[\s\S]*?<\/p:embeddedFontLst>/)?.[0]
		assert(lst, 'embeddedFontLst present')
		assert(/<p:font typeface="Silkscreen"\/>/.test(lst), `p:font identity carried; got ${lst}`)
		assert(
			/<p:regular r:id="[^"]+"\/>/.test(lst) && /<p:bold r:id="[^"]+"\/>/.test(lst),
			'regular + bold faces carried'
		)
		// embeddedFontLst sits before defaultTextStyle (CT_Presentation index 7).
		assert(
			pres.indexOf('<p:embeddedFontLst') < pres.indexOf('<p:defaultTextStyle'),
			'embeddedFontLst precedes defaultTextStyle'
		)
	})

	test('is idempotent: appending the same generator twice carries each face once', async () => {
		const target = await openTarget()
		await target.appendSlides(await fontGenerator(), { layout: 'Blank' })
		await target.appendSlides(await fontGenerator(), { layout: 'Blank' })

		const zip = await zipOf(await target.save())
		const fontParts = Object.keys(zip.files).filter((n) => /^ppt\/fonts\/font\d+\.fntdata$/.test(n))
		assertEqual(fontParts.length, 2, 'each face written exactly once across repeated appends')

		const pres = await zip.file('ppt/presentation.xml').async('string')
		assertEqual(
			(pres.match(/<p:embeddedFont>/g) || []).length,
			1,
			'a single embeddedFont entry for the shared typeface'
		)
		assertEqual((pres.match(/<p:regular /g) || []).length, 1, 'regular face not duplicated')
		assertEqual((pres.match(/<p:bold /g) || []).length, 1, 'bold face not duplicated')
	})

	test('no embedded fonts on the generator leaves the deck font-free', async () => {
		const target = await openTarget()
		const pptx = new PptxGenJS()
		pptx.layout = 'LAYOUT_WIDE'
		pptx.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1 })
		await target.appendSlides(pptx, { layout: 'Blank' })

		const zip = await zipOf(await target.save())
		assert(!Object.keys(zip.files).some((n) => /fntdata/.test(n)), 'no font parts when the generator embeds none')
		const pres = await zip.file('ppt/presentation.xml').async('string')
		assert(!/embeddedFontLst/.test(pres), 'no embeddedFontLst when the generator embeds none')
	})

	test.skipIf(!validatorInstalled)('a deck with appended embedded fonts stays schema-valid', async () => {
		const target = await openTarget()
		await target.appendSlides(await fontGenerator(), { layout: 'Blank' })
		const errors = await validateBuf(Buffer.from(await target.save()))
		assertEqual(errors.length, 0, `validator errors: ${JSON.stringify(errors).slice(0, 2000)}`)
	})
})
