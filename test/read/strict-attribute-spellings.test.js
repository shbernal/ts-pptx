// The read model against the *other* schema-legal spelling of an attribute it already reads.
//
// Two OOXML types have two written forms, and Office only ever writes one of them:
//
//   - `xsd:boolean` is `1`/`0` *or* `true`/`false`. PowerPoint writes the digits.
//   - `a:ST_Percentage` is a union of the fixed-point integer (`100%` → `100000`) and a
//     decimal string with a literal `%`, and the `%` form is the only one the Strict
//     profile has. PowerPoint (Transitional) writes the integer.
//
// So no fixture in the corpus carries either alternative, and a getter that reads only the
// Office form passes every deck the suite owns while silently dropping a schema-legal value
// from any other producer's file. The cases below take an authored deck and rewrite exactly
// those attributes into the other form, which is the only way to reach that branch.
//
// Rewriting the bytes rather than authoring them is deliberate: this library writes the
// Office form on purpose and should keep doing so, so the alternative can only ever arrive
// as input.

import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { TableStyle } from '../../dist/node.js'
import { PNG_1X1, assert, assertEqual } from '../helpers.js'
import { authorRead, authorReadWithFixtureStyles, firstTable } from './authored.js'

/** Apply `rewrite` to every slide part of `buf` and reload the result. */
async function reloadWithSlideXml(buf, rewrite) {
	const zip = await JSZip.loadAsync(buf)
	for (const name of Object.keys(zip.files)) {
		if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue
		zip.file(name, rewrite(await zip.file(name).async('string')))
	}
	return Presentation.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
}

/** A deck carrying every construct the cases rewrite. */
function authorSpellingDeck() {
	return authorRead((pres) => {
		const slide = pres.addSlide()
		slide.addTable([[{ text: 'H1' }, { text: 'H2' }], [{ text: 'wide', options: { colspan: 2 } }]], {
			x: 0.5,
			y: 0.5,
			w: 6,
			h: 2,
			hasHeader: true,
		})
		slide.addImage({ data: PNG_1X1, x: 0.5, y: 3, w: 2, h: 2, crop: { l: 20, t: 0, r: 0, b: 0 } })
		slide.addShape('rect', {
			x: 4,
			y: 3,
			w: 2,
			h: 2,
			fill: {
				type: 'gradient',
				gradient: {
					kind: 'linear',
					angle: 0,
					stops: [
						{ position: 0, color: '0088CC' },
						{ position: 50, color: 'FFFFFF' },
					],
				},
			},
		})
	})
}

describe('xsd:boolean spelled true/false', () => {
	test('Table.firstRowHeader reads a:tblPr/@firstRow="true"', async () => {
		const { buf } = await authorSpellingDeck()
		const reopened = await reloadWithSlideXml(buf, (xml) => {
			assert(xml.includes('firstRow="1"'), 'the authored deck writes the digit form')
			return xml.replaceAll('firstRow="1"', 'firstRow="true"')
		})
		assertEqual(firstTable(reopened).firstRowHeader, true, 'firstRow="true" is a header row')
	})

	test('TableCell.isMergeContinuation reads a:tc/@hMerge="true"', async () => {
		const { buf } = await authorSpellingDeck()
		const reopened = await reloadWithSlideXml(buf, (xml) => {
			assert(xml.includes('hMerge="1"'), 'the colspan authors a digit-form hMerge')
			return xml.replaceAll('hMerge="1"', 'hMerge="true"')
		})
		assertEqual(firstTable(reopened).cell(1, 1).isMergeContinuation, true, 'hMerge="true" covers the cell')
	})

	test('Slide.hidden reads p:sld/@show="false"', async () => {
		const { buf } = await authorSpellingDeck()
		const hiddenPresentation = await Presentation.load(buf)
		hiddenPresentation.slides[0].hidden = true
		const reopened = await reloadWithSlideXml(await hiddenPresentation.save(), (xml) => {
			assert(xml.includes('show="0"'), 'the setter writes the digit form')
			return xml.replaceAll('show="0"', 'show="false"')
		})
		assertEqual(reopened.slides[0].hidden, true, 'show="false" is a hidden slide')
	})
})

describe('a:ST_Percentage spelled with a literal %', () => {
	test('Picture.crop reads a:srcRect/@l="10%"', async () => {
		const { buf } = await authorSpellingDeck()
		const reopened = await reloadWithSlideXml(buf, (xml) => {
			assert(/<a:srcRect l="\d+"/.test(xml), 'the authored deck writes the fixed-point form')
			return xml.replace(/<a:srcRect l="\d+"/, '<a:srcRect l="10%"')
		})
		const picture = reopened.slides[0].shapes.find((shape) => shape.shapeType === 'picture')
		assertEqual(picture.crop.left, 0.1, 'l="10%" is a tenth of the source width')
	})

	test('gradient stop positions read a:gs/@pos="50%"', async () => {
		const { buf } = await authorSpellingDeck()
		const reopened = await reloadWithSlideXml(buf, (xml) => {
			assert(xml.includes('pos="50000"'), 'the authored deck writes the fixed-point form')
			return xml.replaceAll('pos="50000"', 'pos="50%"')
		})
		const shape = reopened.slides[0].shapes.find((s) => s.gradientFill)
		const positions = shape.gradientFill.stops.map((stop) => stop.position)
		assert(positions.includes(0.5), `a 50% stop reads as 0.5, got ${JSON.stringify(positions)}`)
	})
})

describe('what the digit-only readings cost downstream', () => {
	test('a table style resolves its header shading through firstRow="true"', async () => {
		// `#tblPrFlag` read `=== '1'`, so a producer that spells the flag `true` (LibreOffice
		// does) read `firstRowHeader` as false. The style context's flags then came back all
		// false, `cellStyleParts` never emitted `a:firstRow`, and a header cell reported the
		// `wholeTbl` shading for a row PowerPoint paints in the accent colour.
		const { buf } = await authorReadWithFixtureStyles((pres) => {
			pres.addSlide().addTable(
				[
					[{ text: 'H1' }, { text: 'H2' }],
					[{ text: 'a' }, { text: 'b' }],
				],
				{ x: 0.5, y: 0.5, w: 6, h: 2, hasHeader: true, tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1 }
			)
		})
		const digits = await Presentation.load(buf)
		const words = await reloadWithSlideXml(buf, (xml) => xml.replaceAll('firstRow="1"', 'firstRow="true"'))
		const fillOf = (presentation) => firstTable(presentation).cell(0, 0).resolvedFill?.effectiveHex ?? null
		assert(fillOf(digits), 'the digit spelling resolves a header fill')
		assertEqual(fillOf(words), fillOf(digits), 'and so does the word spelling')
	})

	test('a cropped picture spelled in percent reports the same crop as the fixed-point form', async () => {
		// `intValue('10%')` was `null`, so `Picture.crop` reported zeros and the script converter
		// wrote a crop of nothing.
		const { buf } = await authorSpellingDeck()
		const fixed = await Presentation.load(buf)
		const percent = await reloadWithSlideXml(buf, (xml) => xml.replace(/<a:srcRect l="\d+"/, '<a:srcRect l="20%"'))
		const cropOf = (presentation) =>
			presentation.slides[0].shapes.find((shape) => shape.shapeType === 'picture').crop.left
		assertEqual(cropOf(fixed), 0.2, 'the authored deck crops a fifth off the left')
		assertEqual(cropOf(percent), cropOf(fixed), 'and `20%` is the same fifth')
	})
})
