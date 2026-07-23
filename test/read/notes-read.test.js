// Write→read fidelity for the speaker-notes rich text frame (src/read/api/slide.ts).
//
// `slide.addNotes([{ text, options }])` authors the notes-slide body placeholder,
// serializing each run with the standard text-run generator — so bold/italic/colour
// and an external hyperlink land in `<p:txBody>` just as on any shape. The read side
// previously exposed only `Slide.notesText`, a flattened string that dropped all of
// it. `Slide.notesTextFrame` now hands back the same body as a navigable TextFrame
// (paragraphs → runs), threaded with the notes part's own rels so a notes hyperlink
// resolves its url.
//
// The writer splits a run on `\n` into separate `<a:p>` paragraphs, so a multi-line
// note round-trips as multiple paragraphs. `notesText` stays as the flattened
// convenience and must keep agreeing with the frame's joined text.

import { describe, test } from 'vitest'
import { authorRead, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** The first slide of `presentation`. */
function firstSlide(presentation) {
	const slide = presentation.slides[0]
	assert(slide, 'the authored slide is read back')
	return slide
}

describe('Slide.notesTextFrame — write→read fidelity', () => {
	test('a bold/coloured notes run survives as a navigable run', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes([{ text: 'Loud', options: { bold: true, color: 'C00000' } }, { text: ' and plain' }])
		})

		const frame = firstSlide(presentation).notesTextFrame
		assert(frame, 'the notes body is exposed as a text frame')
		const runs = frame.paragraphs[0].runs
		assertEqual(runs.length, 2, 'both notes runs read back')

		assertEqual(runs[0].text, 'Loud', 'first run text round-trips')
		assertEqual(runs[0].bold, true, 'first run stays bold')
		assertEqual(runs[0].color?.toUpperCase(), 'C00000', 'first run colour round-trips')

		assertEqual(runs[1].text, ' and plain', 'second run text round-trips')
		assertEqual(runs[1].bold, null, 'the second run carries no bold of its own')
		assertEqual(runs[1].color, null, 'the second run carries no colour of its own')
	})

	test('a newline in a note splits into separate paragraphs', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes('line one\nline two')
		})

		const slide = firstSlide(presentation)
		const frame = slide.notesTextFrame
		assert(frame, 'the notes body is exposed as a text frame')
		assertEqual(frame.paragraphs.length, 2, 'the newline starts a second paragraph')
		assertEqual(frame.paragraphs[0].text, 'line one', 'first paragraph text')
		assertEqual(frame.paragraphs[1].text, 'line two', 'second paragraph text')

		// The flattened convenience keeps agreeing with the frame (paragraphs joined by \n).
		assertEqual(slide.notesText, 'line one\nline two', 'notesText stays the flattened view')
	})

	test('a notes hyperlink resolves its url through the notes part rels', async () => {
		const { presentation } = await authorRead((pres) => {
			pres
				.addSlide()
				.addNotes([
					{ text: 'see ' },
					{ text: 'the site', options: { hyperlink: { url: 'https://example.com/notes' } } },
				])
		})

		const frame = firstSlide(presentation).notesTextFrame
		assert(frame, 'the notes body is exposed as a text frame')
		const linked = frame.paragraphs[0].runs.find((r) => r.hyperlink !== null)
		assert(linked, 'the linked run reads back with a hyperlink')
		assertEqual(
			linked.hyperlink.url,
			'https://example.com/notes',
			'the notes hyperlink url resolves via the notes rels'
		)
	})

	test('a slide with no addNotes still exposes an empty notes frame', async () => {
		// The writer attaches an empty notes part to every slide (to keep the notesSlide
		// rel/_rels bookkeeping uniform), so an authored slide never hits the true
		// no-notes-part path — that null branch is import-only. Here the notes part
		// exists but its body is empty: a frame with a single empty paragraph, and
		// `notesText === ''` (the '' vs null distinction the getters document).
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('body only', { x: 1, y: 1, w: 3, h: 1 })
		})
		const slide = firstSlide(presentation)
		assertEqual(slide.notesText, '', 'an empty notes part flattens to the empty string, not null')
		const frame = slide.notesTextFrame
		assert(frame, 'the empty notes body is still a frame')
		assertEqual(frame.paragraphs.length, 1, 'the empty body carries a single empty paragraph')
		assertEqual(frame.paragraphs[0].text, '', 'that paragraph is empty')
	})

	test.skipIf(!validatorInstalled)('the authored rich-notes deck is schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			pres
				.addSlide()
				.addNotes([
					{ text: 'Loud', options: { bold: true, color: 'C00000' } },
					{ text: ' and a ' },
					{ text: 'link', options: { hyperlink: { url: 'https://example.com/notes' } } },
				])
		})
		assertEqual((await schemaErrors(buf)).length, 0, 'rich-notes deck validates')
	})
})
