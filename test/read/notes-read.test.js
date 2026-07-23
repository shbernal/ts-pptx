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
//
// T1.2 threads a notes theme context (resolved through the notesMaster → theme2.xml
// chain) into the frame, so a notes run authored with a *scheme* colour resolves to
// a literal hex via `Run.resolvedColor` — previously inert (null) because the frame
// was given no theme context.
//
// T2.1 models the notes slide as its own object: `Slide.notesSlide` → a `NotesSlide`
// exposing the three placeholders (`slideImage`/`body`/`slideNumber`), each with its
// geometry and (where present) a text frame. The writer authors all three but leaves
// `sldImg`/`sldNum` with an empty `p:spPr`, so their geometry reads `null` on an
// authored deck (import-only); the body text and the slide-number field round-trip.
// `notesText`/`notesTextFrame` are now thin delegates over the body placeholder.

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

	test('a scheme-coloured notes run resolves to a theme hex via resolvedColor (T1.2)', async () => {
		const { presentation } = await authorRead((pres) => {
			// A scheme colour authors <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
			// in the notes run — the trigger for the theme-resolved getter.
			pres.addSlide().addNotes([{ text: 'Themed', options: { color: 'accent1' } }])
		})

		const run = firstSlide(presentation).notesTextFrame.paragraphs[0].runs[0]
		assert(run, 'the themed notes run reads back')
		// Own-attribute getters: the scheme token is surfaced, and `color` (hex only) is null.
		assertEqual(run.schemeColor, 'accent1', 'the scheme token is read as the own attribute')
		assertEqual(run.color, null, 'a scheme fill has no explicit hex')
		// The fix: the notes theme context resolves accent1 through notesMaster → theme2.xml.
		assert(run.resolvedColor, 'the scheme colour now resolves against the notes theme (was null)')
		assertEqual(run.resolvedColor.hex, '4472C4', 'accent1 resolves to the default Office theme hex')
		assertEqual(run.resolvedColor.effectiveHex, '4472C4', 'no transforms, so effectiveHex equals the base hex')
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

	test('notesSlide models the three placeholders (T2.1)', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes('a note')
		})

		const notes = firstSlide(presentation).notesSlide
		assert(notes, 'the notes slide is exposed as a modeled object')

		// The writer authors exactly the slide-thumbnail, body, and slide-number placeholders.
		const types = notes.placeholders.map((ph) => ph.type)
		assertEqual(types.length, 3, 'the notes slide has three placeholders')
		assert(types.includes('sldImg'), 'a slide-thumbnail placeholder is present')
		assert(types.includes('body'), 'a notes body placeholder is present')
		assert(types.includes('sldNum'), 'a slide-number placeholder is present')

		assert(notes.slideImage, 'slideImage resolves the sldImg placeholder')
		assertEqual(notes.slideImage.type, 'sldImg', 'slideImage is the thumbnail placeholder')
		assert(notes.body, 'body resolves the notes body placeholder')
		assertEqual(notes.body.text, 'a note', 'the body placeholder round-trips the note text')
	})

	test('the slide-number placeholder carries the slide number field (T2.1)', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes('n')
		})

		const notes = firstSlide(presentation).notesSlide
		assert(notes?.slideNumber, 'the slide-number placeholder reads back')
		// The writer emits <a:fld type="slidenum"> whose text is the 1-based slide number;
		// TextFrame.text reads a:fld text, so the field surfaces as '1' for the first slide.
		assertEqual(notes.slideNumber.type, 'sldNum', 'it is the sldNum placeholder')
		assertEqual(notes.slideNumber.text, '1', 'the field shows the first slide number')
		// The thumbnail placeholder has no text body at all.
		assertEqual(notes.slideImage?.textFrame, null, 'the sldImg thumbnail carries no text frame')
	})

	test('authored notes placeholders inherit geometry (null own xfrm) (T2.1)', async () => {
		// The writer leaves sldImg/sldNum p:spPr empty, so their geometry is inherited
		// from the notesMaster — the read model reports null for the placeholder's own
		// transform (a positive number would only come from an imported deck).
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes('n')
		})

		const notes = firstSlide(presentation).notesSlide
		assert(notes, 'the notes slide reads back')
		for (const ph of notes.placeholders) {
			assertEqual(ph.left, null, `${ph.type} has no own left (inherited geometry)`)
			assertEqual(ph.top, null, `${ph.type} has no own top`)
			assertEqual(ph.width, null, `${ph.type} has no own width`)
			assertEqual(ph.height, null, `${ph.type} has no own height`)
		}
	})

	test('a slide with no notes part has a null notesSlide', async () => {
		// The writer attaches a notes part to every authored slide, so this exercises the
		// import-only null boundary via a slide whose notes rel is absent. Here we assert
		// the modeled getter agrees with notesText/notesTextFrame on a real authored slide:
		// all three resolve the same body.
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes('shared body')
		})
		const slide = firstSlide(presentation)
		assertEqual(slide.notesSlide?.text, 'shared body', 'notesSlide.text matches the body')
		assertEqual(slide.notesText, 'shared body', 'notesText delegates to the same body')
		assertEqual(
			slide.notesTextFrame?.paragraphs[0].text,
			'shared body',
			'notesTextFrame delegates to the same body frame'
		)
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
