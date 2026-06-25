// Feature B (author-side embedding): pptx.embedFont() embeds a whole font face so
// the deck renders with it on machines lacking the font. Asserts the public API
// shape — byte sources (Uint8Array / ArrayBuffer / base64), multi-face
// accumulation under one typeface, and input validation — plus the emitted package
// pieces (font parts, content-type Default, presentation rels, embeddedFontLst,
// embedTrueTypeFonts/saveSubsetFonts). Schema validity is covered in
// test/schema.test.js; structural emit is covered here without the validator.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test, beforeAll } from 'vitest'
import PptxGenJS from '../../dist/node.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const fontsDir = path.join(__dirname, '..', 'read', 'fixtures', 'fonts')

let regular
let bold

beforeAll(async () => {
	regular = new Uint8Array(await readFile(path.join(fontsDir, 'Silkscreen-Regular.ttf')))
	bold = new Uint8Array(await readFile(path.join(fontsDir, 'Silkscreen-Bold.ttf')))
})

async function zipOf(pres) {
	return JSZip.loadAsync(await pres.stream())
}

describe('PptxGenJS.embedFont', () => {
	test('embeds a single regular face: part + Default + rel + list + flags', async () => {
		const p = new PptxGenJS()
		await p.embedFont({ data: regular, typeface: 'Silkscreen' })
		p.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1, fontFace: 'Silkscreen' })
		const zip = await zipOf(p)

		const font1 = await zip.file('ppt/fonts/font1.fntdata')?.async('uint8array')
		assert(font1, 'font1.fntdata written')
		assertEqual(font1.length, regular.length, 'whole face embedded (not subset)')

		const ct = await zip.file('[Content_Types].xml').async('string')
		assert(/<Default Extension="fntdata" ContentType="application\/x-fontdata"\/>/.test(ct), 'fntdata Default')

		const rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string')
		assert(/relationships\/font" Target="fonts\/font1\.fntdata"/.test(rels), 'font rel targets the part')

		const pres = await zip.file('ppt/presentation.xml').async('string')
		assert(/embedTrueTypeFonts="1"/.test(pres), 'embedTrueTypeFonts on')
		assert(/saveSubsetFonts="0"/.test(pres), 'saveSubsetFonts off (whole faces)')
		assert(
			/<p:embeddedFontLst><p:embeddedFont><p:font typeface="Silkscreen"\/><p:regular r:id="rId\d+"\/><\/p:embeddedFont><\/p:embeddedFontLst>/.test(
				pres
			),
			`embeddedFontLst entry; got ${pres.match(/<p:embeddedFontLst>.*?<\/p:embeddedFontLst>/)}`
		)
	})

	test('accumulates multiple styles of one typeface into a single embeddedFont entry', async () => {
		const p = new PptxGenJS()
		await p.embedFont({ data: regular, typeface: 'Silkscreen' })
		await p.embedFont({ data: bold, typeface: 'Silkscreen', style: 'bold' })
		p.addSlide()
		const zip = await zipOf(p)

		assert(await zip.file('ppt/fonts/font1.fntdata'), 'first face part')
		assert(await zip.file('ppt/fonts/font2.fntdata'), 'second face part')

		const pres = await zip.file('ppt/presentation.xml').async('string')
		assertEqual((pres.match(/<p:embeddedFont>/g) || []).length, 1, 'one embeddedFont entry for the family')
		assert(/<p:regular r:id="rId\d+"\/><p:bold r:id="rId\d+"\/>/.test(pres), 'regular then bold, in schema order')
	})

	test('two distinct typefaces yield two embeddedFont entries', async () => {
		const p = new PptxGenJS()
		await p.embedFont({ data: regular, typeface: 'Silkscreen' })
		await p.embedFont({ data: bold, typeface: 'Other Face' })
		p.addSlide()
		const zip = await zipOf(p)
		const pres = await zip.file('ppt/presentation.xml').async('string')
		assertEqual((pres.match(/<p:embeddedFont>/g) || []).length, 2, 'two entries')
	})

	test('accepts ArrayBuffer and base64-string byte sources', async () => {
		const p = new PptxGenJS()
		// ArrayBuffer slice (exact backing buffer for the face bytes)
		await p.embedFont({
			data: regular.buffer.slice(regular.byteOffset, regular.byteOffset + regular.byteLength),
			typeface: 'AB',
		})
		// raw base64 (no data-URL prefix)
		await p.embedFont({ data: Buffer.from(bold).toString('base64'), typeface: 'B64' })
		p.addSlide()
		const zip = await zipOf(p)
		const f1 = await zip.file('ppt/fonts/font1.fntdata').async('uint8array')
		const f2 = await zip.file('ppt/fonts/font2.fntdata').async('uint8array')
		assertEqual(f1.length, regular.length, 'ArrayBuffer face length')
		assertEqual(f2.length, bold.length, 'base64 face length')
	})

	test('a repeat of the same typeface+style replaces the prior bytes (last wins)', async () => {
		const p = new PptxGenJS()
		await p.embedFont({ data: bold, typeface: 'Silkscreen' }) // wrong bytes first
		await p.embedFont({ data: regular, typeface: 'Silkscreen' }) // corrected
		p.addSlide()
		const zip = await zipOf(p)
		const names = Object.keys(zip.files).filter((n) => /fntdata/.test(n))
		assertEqual(names.length, 1, 'still a single face part')
		const f1 = await zip.file('ppt/fonts/font1.fntdata').async('uint8array')
		assertEqual(f1.length, regular.length, 'last call bytes win')
	})

	test('validates input: missing typeface and missing source throw', async () => {
		const p = new PptxGenJS()
		await assertRejects(() => p.embedFont({ data: regular }), 'missing typeface throws')
		await assertRejects(() => p.embedFont({ typeface: 'X' }), 'missing source throws')
		await assertRejects(() => p.embedFont({ data: regular, typeface: 'X', style: 'heavy' }), 'invalid style throws')
	})

	test('no embedFont calls: deck is unchanged (no font parts/list, historical flags)', async () => {
		const p = new PptxGenJS()
		p.addSlide()
		const zip = await zipOf(p)
		assert(!Object.keys(zip.files).some((n) => /fntdata/.test(n)), 'no font parts')
		const pres = await zip.file('ppt/presentation.xml').async('string')
		assert(!/embeddedFontLst/.test(pres), 'no embeddedFontLst')
		assert(!/embedTrueTypeFonts/.test(pres), 'embedTrueTypeFonts not emitted')
		assert(/saveSubsetFonts="1"/.test(pres), 'historical saveSubsetFonts="1" preserved')
	})
})

async function assertRejects(fn, label) {
	let threw = false
	try {
		await fn()
	} catch {
		threw = true
	}
	assert(threw, label)
}
