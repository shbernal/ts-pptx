// Copies and imports carry a part as it is now, not the bytes it was loaded with.
//
// A part edited through its DOM reflects the edit only through `serialize()`; its
// `originalBytes` are what the zip held. Every copy operation used to read the loaded bytes, so
// an edit made earlier in the session was silently missing from the copy: a cloned slide came
// back with the old text, an import from an edited deck imported the unedited page, and a clone
// of a slide `importSlide({ theme: 'preserve' })` had just flattened came back with the source's
// scheme colours. Each case is checked after a save and a reload, which is what a caller sees.

import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const EDITED = 'EDITED IN SESSION'

/** `textbox.pptx` with the first text shape on slide 1 rewritten, and not yet saved. */
async function editedTextbox() {
	const deck = await openFixture('textbox')
	deck.slides[0].shapes.find((shape) => shape.hasTextFrame).text = EDITED
	return deck
}

const reload = async (deck) => Presentation.load(await deck.save())
const firstText = (slide) => slide.shapes.find((shape) => shape.hasTextFrame)?.text

describe('copies carry the part as it is now', () => {
	test('cloneSlide copies an edit made earlier in the session', async () => {
		const deck = await editedTextbox()
		const clone = deck.cloneSlide(0)
		const reopened = await reload(deck)
		assertEqual(firstText(reopened.slides[0]), EDITED, 'the source slide kept its edit')
		assertEqual(firstText(reopened.slides[clone.index]), EDITED, 'and the clone carries it')
	})

	test.for(/** @type {const} */ (['copy', 'preserve', 'restyle']))(
		'importSlide in %s mode copies an edited source slide',
		async (theme) => {
			const source = await editedTextbox()
			const target = await openFixture('textbox')
			const imported = target.importSlide(source, 0, { theme })
			const reopened = await reload(target)
			assertEqual(firstText(reopened.slides[imported.index]), EDITED, `the ${theme} import carries the edit`)
		}
	)

	test('importSlides copies an edited source slide', async () => {
		const source = await editedTextbox()
		const target = await openFixture('textbox')
		target.importSlides([{ source, sourceIndex: 0, outputIndex: 0 }])
		const reopened = await reload(target)
		assertEqual(firstText(reopened.slides[0]), EDITED, 'the batch import carries the edit')
	})

	test('a clone of a preserve-imported slide keeps the colours the import baked in', async () => {
		const target = await openFixture('empty')
		const source = await openFixture('multi-theme')
		const imported = target.importSlide(source, 0, { theme: 'preserve' })
		const clone = target.cloneSlide(imported.index)
		const zip = await JSZip.loadAsync(await target.save())
		const xml = await zip.file(clone.partName.slice(1)).async('string')
		assert(!/schemeClr/.test(xml), 'no scheme token came back into the clone')
		assert(/<a:srgbClr val="B01513"/.test(xml), 'the clone keeps the flattened accent1')
	})

	test('cloneSlide copies a shape imported onto the slide this session', async () => {
		const target = await openFixture('textbox')
		const source = await openFixture('textbox')
		const before = target.slides[0].shapes.length
		target.importShape(target.slides[0], source.slides[0], 0)
		const clone = target.cloneSlide(0)
		const reopened = await reload(target)
		assertEqual(reopened.slides[0].shapes.length, before + 1, 'the host slide kept the imported shape')
		assertEqual(reopened.slides[clone.index].shapes.length, before + 1, 'and so does its clone')
	})
})
