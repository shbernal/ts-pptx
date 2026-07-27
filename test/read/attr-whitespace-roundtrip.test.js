// Write→read fidelity for whitespace inside an XML ATTRIBUTE value (dn-xml-attr-whitespace).
//
// XML 1.0 section 3.3.3 requires a parser to normalise a literal tab, carriage return or line feed
// inside an attribute value to a single space *before any consumer sees it*. So a caller-supplied
// string carrying a line break — a layout title, an objectName, alt text, a section title — used to
// come back with that break silently flattened to a space. Preserving one requires a character
// reference (`&#9;` / `&#10;` / `&#13;`), which `encodeXmlAttrValue` (src/gen-utils.ts) now emits
// for every attribute-emitting path, wired through `src/gen/oxml/el.ts` and `cNvPrOpen`.
//
// The oracle is deliberately a real parser, not a substring match on the emitted bytes: the defect
// is precisely that the bytes look fine and the *parse* is what loses information. A byte assertion
// alone would pass against a writer that emitted `&#10;` inside an element's text, where it is not
// needed, so each case is asserted through `Presentation.load` (@xmldom/xmldom).
//
// The layout case is not hypothetical: PowerPoint's built-in German layout set ships a layout named
// across two lines ("Abschnitts-<LF>überschrift"), so a deck created from it already hits this.

import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { assert, assertEqual } from '../helpers.js'
import { authorRead, firstShape, schemaErrors, validatorInstalled } from './authored.js'

/** The German built-in layout name, split across two lines exactly as PowerPoint ships it. */
const LAYOUT_TITLE = 'Abschnitts-\nüberschrift'
const OBJECT_NAME = 'Kapitel\nEins'
const ALT_TEXT = 'Zeile eins\nZeile zwei\tmit Tabulator'
const SECTION_TITLE = 'Teil\nEins'
const TOOLTIP = 'Zeile\nZwei'

/** Author the whitespace-bearing deck. */
async function authorWhitespaceDeck() {
	return authorRead((pres) => {
		pres.defineSlideMaster({ title: LAYOUT_TITLE, background: { color: 'FFFFFF' } })
		pres.addSection({ title: SECTION_TITLE })
		const slide = pres.addSlide({ masterTitle: LAYOUT_TITLE, sectionTitle: SECTION_TITLE })
		slide.addText('Kapitel', {
			x: 1,
			y: 1,
			w: 4,
			h: 1,
			objectName: OBJECT_NAME,
			altText: ALT_TEXT,
			hyperlink: { url: 'https://example.com/', tooltip: TOOLTIP },
		})
	})
}

/** The first element named `tag` anywhere in `part`, or null. */
function firstElement(part, tag) {
	return part.dom.getElementsByTagName(tag)[0] ?? null
}

describe('XML attribute whitespace — write→read fidelity', () => {
	test('an objectName and alt text keep their line feed and tab', async () => {
		const { presentation } = await authorWhitespaceDeck()
		const shape = firstShape(presentation, (s) => s.name === OBJECT_NAME)
		assert(shape, `the shape named ${JSON.stringify(OBJECT_NAME)} is found by its exact name`)
		assertEqual(shape.name, OBJECT_NAME, 'p:cNvPr/@name round-trips')
		assertEqual(shape.description, ALT_TEXT, 'p:cNvPr/@descr round-trips (line feed AND tab)')
	})

	test('a layout title spanning two lines round-trips', async () => {
		const { presentation } = await authorWhitespaceDeck()
		// `defineSlideMaster({ title })` names the LAYOUT part's `p:cSld/@name`; the master part
		// itself is written unnamed (`name=""`), so the layout gallery is the whole oracle here.
		const named = presentation.layouts().filter((layout) => layout.name === LAYOUT_TITLE)
		assertEqual(named.length, 1, 'exactly one layout carries the two-line title')
	})

	test('a section title round-trips through p14:section/@name', async () => {
		const { presentation } = await authorWhitespaceDeck()
		const section = firstElement(presentation.presentationPart, 'p14:section')
		assert(section, 'the authored section is present in presentation.xml')
		assertEqual(section.getAttribute('name'), SECTION_TITLE, 'p14:section/@name round-trips')
	})

	test('a hyperlink tooltip round-trips through a:hlinkClick/@tooltip', async () => {
		const { presentation } = await authorWhitespaceDeck()
		const hlink = firstElement(presentation.slides[0].part, 'a:hlinkClick')
		assert(hlink, 'the authored hyperlink is present on the slide')
		assertEqual(hlink.getAttribute('tooltip'), TOOLTIP, 'a:hlinkClick/@tooltip round-trips')
	})

	test('the whitespace is emitted as a character reference, never a literal', async () => {
		// The complement to the parse-level assertions: proves the round-trip is carried by `&#10;`
		// rather than by a parser that happens to be lenient about literal newlines in attributes.
		const { buf } = await authorWhitespaceDeck()
		const zip = await JSZip.loadAsync(buf)
		const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')
		const cNvPr = slideXml.slice(slideXml.indexOf('<p:cNvPr'), slideXml.indexOf('</p:nvSpPr>'))
		assert(cNvPr.includes('name="Kapitel&#10;Eins"'), `objectName uses a character reference: ${cNvPr}`)
		assert(cNvPr.includes('&#9;'), 'the alt text tab uses a character reference')
		assert(!/name="[^"]*\n[^"]*"/.test(cNvPr), 'no literal newline survives inside an attribute value')
	})

	test('the deck stays schema-valid with character references in attributes', async () => {
		if (!validatorInstalled) return
		const { buf } = await authorWhitespaceDeck()
		const errors = await schemaErrors(buf)
		assertEqual(errors.length, 0, `schema errors: ${errors.join('\n')}`)
	})
})
