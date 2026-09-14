// importSlide({ theme: 'preserve' }) and the source deck's p:defaultTextStyle.
//
// default-text-style.pptx's PlainBox is a text box whose bare run takes 18pt, tx1 and the minor font
// from the presentation's default text style; its StyledRect takes its size from there and its
// colour, white, from a p:style fontRef. table-text-inheritance.pptx's default text style is 28pt,
// C00000 and Courier New. Pasting the one slide into the other in PowerPoint
// (test/read/fixtures/authoring/probe-default-text-style-paste.ps1) with "Keep Source Formatting"
// leaves both runs at 18pt, PlainBox black and StyledRect white, because PowerPoint writes those
// values onto the runs; with "Use Destination Theme" both turn 28pt. `preserve` is the first.

import { describe, test } from 'vitest'

import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

/** The first run of the shape named `name` on `slide`. */
function runOf(slide, name) {
	const shape = slide.shapes.find((s) => s.name === name)
	assert(shape?.textFrame, `expected a text shape named ${name}`)
	return shape.textFrame.paragraphs[0].runs[0]
}

describe("importSlide({ theme: 'preserve' }) keeps what runs took from the source p:defaultTextStyle", () => {
	test('a text box keeps the source size and colour in a deck whose default text style differs', async () => {
		const target = await openFixture('table-text-inheritance')
		const imported = target.importSlide(await openFixture('default-text-style'), 0, { theme: 'preserve' })
		const slide = (await Presentation.load(await target.save())).slides.find((s) => s.partName === imported.partName)
		assert(slide, 'the imported slide is found after a save')

		const plain = runOf(slide, 'PlainBox')
		assertEqual(plain.resolvedSizePt, 18, "the source default text style's 18pt, not the destination's 28pt")
		assertEqual(plain.resolvedColor?.effectiveHex, '000000', "the source tx1, not the destination's C00000")

		const styled = runOf(slide, 'StyledRect')
		assertEqual(styled.resolvedSizePt, 18, 'the size a fontRef never carries comes from the source default text style')
		assertEqual(
			styled.resolvedColor?.effectiveHex,
			'FFFFFF',
			'the fontRef colour still wins over the default text style'
		)
	})

	test('a run that states its own size keeps it', async () => {
		const target = await openFixture('table-text-inheritance')
		const source = await openFixture('default-text-style')
		const run = runOf(source.slides[0], 'PlainBox')
		run.fontSizePt = 40
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const slide = (await Presentation.load(await target.save())).slides.find((s) => s.partName === imported.partName)
		assertEqual(runOf(slide, 'PlainBox').resolvedSizePt, 40, 'an own size is not baked over')
	})
})
