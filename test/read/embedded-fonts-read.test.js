// Read accessor for embedded fonts: Presentation.embeddedFonts enumerates the
// deck's p:embeddedFontLst entries and resolves each face's r:id to the absolute
// partname of its `.fntdata` binary. Oracle: the PowerPoint-authored fixture
// test/read/fixtures/embedded-fonts.pptx + embedded-fonts.oracle.json (one
// p:embeddedFont, typeface 'Silkscreen', regular→font1 / bold→font2).

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}
async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

describe('Presentation.embeddedFonts', () => {
	test('enumerates the embedded typeface and resolves each face to its font part', async () => {
		const pres = await open('embedded-fonts')
		const fonts = pres.embeddedFonts

		assertEqual(fonts.length, 1, 'one embedded typeface')
		const font = fonts[0]
		assertEqual(font.typeface, 'Silkscreen', 'typeface read from p:font/@typeface')
		assertEqual(font.panose, null, 'no @panose on this entry → null (fixture has none)')

		assertEqual(font.faces.length, 2, 'two faces embedded')
		const bySlot = Object.fromEntries(font.faces.map((f) => [f.slot, f.partName]))
		assert('regular' in bySlot, 'regular face present')
		assert('bold' in bySlot, 'bold face present')
		assert(!('italic' in bySlot) && !('boldItalic' in bySlot), 'only the two authored faces')

		// Faces are ordered by the schema slot order (regular before bold).
		assertEqual(font.faces[0].slot, 'regular', 'regular first')
		assertEqual(font.faces[1].slot, 'bold', 'bold second')

		// Each partName resolves to a real .fntdata part in the package.
		for (const face of font.faces) {
			assert(/\/ppt\/fonts\/font\d+\.fntdata$/.test(face.partName), `absolute font partname (${face.partName})`)
			assert(pres.opc.part(face.partName), `partName resolves to a real part (${face.partName})`)
		}
		// regular → font1, bold → font2 (oracle rId3 → font1.fntdata, rId4 → font2.fntdata).
		assert(bySlot.regular.endsWith('/font1.fntdata'), `regular → font1 (${bySlot.regular})`)
		assert(bySlot.bold.endsWith('/font2.fntdata'), `bold → font2 (${bySlot.bold})`)
	})

	test('is [] for a deck that embeds no fonts', async () => {
		const pres = await open('empty')
		assertEqual(pres.embeddedFonts.length, 0, 'no embeddedFontLst → []')
	})

	test('round-trips through a write→read carry (importSlide embedFonts)', async () => {
		// The write side carries a source deck's embedded fonts into a target; the
		// read accessor then sees the merged list on the target — a genuine
		// write→read round-trip, not just a fixture read.
		const target = await open('empty')
		const source = await open('embedded-fonts')
		target.importSlide(source, 0, { embedFonts: true })

		const reopened = await Presentation.load(await target.save())
		const fonts = reopened.embeddedFonts
		assertEqual(fonts.length, 1, 'carried typeface reads back')
		assertEqual(fonts[0].typeface, 'Silkscreen', 'typeface survives the carry')
		assertEqual(fonts[0].faces.length, 2, 'both faces survive')
		for (const face of fonts[0].faces) {
			assert(reopened.opc.part(face.partName), `carried face part resolves (${face.partName})`)
		}
	})
})
