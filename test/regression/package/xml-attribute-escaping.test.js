// One XML escaper, reached from all three sides.
//
// Three escapers with three notions of what is dangerous used to live in three modules, and the
// weakest was the one on the least-checked input. `serializeEmbeddedFontLst` put `font.typeface`
// — straight from `pptx.embedFont({ typeface })`, unvalidated — through a local escaper that
// handled `& < > "` and nothing else. So a control character XML 1.0 forbids went into
// `<p:font typeface="…">` verbatim, and a newline was silently normalised to a space by any
// parser that read the file back, while every other attribute in the package had both handled.
//
// No showcase deck embeds a font with an exotic typeface, so byte identity cannot see this.
//
// The read/edit path (`read/opc/*`) had the middle escaper: tab/CR/LF handled, control
// characters not, `'` not. Its writers now share the same one.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import JSZip from 'jszip'
import { describe, test, expect, beforeAll } from 'vitest'
import TsPptx from '../../../dist/node.js'
import { FIXTURES } from '../../read/corpus.js'

/** A vertical tab: legal in a JavaScript string, forbidden in XML 1.0 in any spelling. */
const VERTICAL_TAB = String.fromCharCode(11)

/** The XML 1.0 illegal control set, as a detector — the same set the write side strips. */
const ILLEGAL_XML_CHARS = new RegExp(
	((cc) => `[${cc(0)}-${cc(8)}${cc(11)}${cc(12)}${cc(14)}-${cc(31)}${cc(127)}]`)(String.fromCharCode)
)

let regular

beforeAll(async () => {
	regular = new Uint8Array(await readFile(path.join(FIXTURES, 'fonts', 'Silkscreen-Regular.ttf')))
})

/** `ppt/presentation.xml` of a deck with one embedded face under `typeface`. */
async function presentationXml(typeface) {
	const pres = new TsPptx()
	await pres.embedFont({ data: regular, typeface })
	pres.addSlide().addText('hi', { x: 1, y: 1, w: 4, h: 1, fontFace: typeface })
	const zip = await JSZip.loadAsync(await pres.toBytes())
	return zip.file('ppt/presentation.xml').async('string')
}

describe('embedded font typeface escaping', () => {
	test('a control character is stripped rather than written into the attribute', async () => {
		// A vertical tab is not representable in XML 1.0 at all — not literally and not as a
		// character reference — so a document carrying one is malformed and PowerPoint offers to
		// repair it. Every other emission site in the library strips this set; this one did not.
		const xml = await presentationXml(`Silk${VERTICAL_TAB}screen`)

		expect(ILLEGAL_XML_CHARS.test(xml)).toBe(false)
		expect(xml).toContain('<p:font typeface="Silkscreen"')
	})

	test('a newline survives as a character reference instead of collapsing to a space', async () => {
		// XML 1.0 section 3.3.3: a parser normalises a literal tab, CR or LF inside an attribute
		// value to a single space before any consumer sees it. Carrying one across therefore
		// requires the reference, which is what the write side's own escaper has always emitted.
		const xml = await presentationXml('Silk\nscreen')

		expect(xml).toContain('<p:font typeface="Silk&#10;screen"')
	})

	test('an apostrophe is escaped, as it is everywhere else in the package', async () => {
		const xml = await presentationXml("Bob's Screen")

		expect(xml).toContain('<p:font typeface="Bob&apos;s Screen"')
	})
})
