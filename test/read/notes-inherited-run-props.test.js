// Write→read fidelity for a notes-body run's INHERITED character properties
// (FIDELITY-BACKLOG F2). T1.2 threaded a notesMaster→theme2 *colour* context into
// the notes frame, so a notes run's own `schemeClr` resolves (`Run.resolvedColor`).
// But the notes frame was built without a placeholder context, so the *inherited*
// getters — the ones that read from the notesMaster's `p:notesStyle` when a run sets
// no own `@sz`/`a:latin`/`@b` — stayed inert (null). This suite covers the fix:
// `NotesPlaceholder.textFrame` now gives the body a placeholder context whose
// `FlattenContext.notesStyle` is the notesMaster's `p:notesStyle`, so a body run's
// `resolvedSizePt`/`resolvedFontFace`/`resolvedBold` (and inherited `resolvedColor`)
// walk that chain the same way a slide placeholder run walks layout→master→txStyles.
//
// Gate: AUTHORABLE. The gate check (recorded in FIDELITY-BACKLOG.md) found the
// backlog's "fixture-only" assumption wrong: a plain `addNotes('text')` serializes
// the run through genXmlTextRunProperties with empty options, which emits NO `@sz`
// and NO `<a:latin>` (src/gen/drawingml/text-run.ts) — a genuine inherit trigger —
// and the writer authors its own notesMaster with a full `p:notesStyle`
// (src/gen/slide/notes.ts). So this is a real write→read round-trip, no fixture.
//
// Oracle (independent of the reader): read straight off the writer's known output.
// makeXmlNotesMaster emits every `p:notesStyle` level with `<a:defRPr sz="1200" ...>`
// and `<a:latin typeface="+mn-lt"/>` and NO `@b`; theme2.xml is makeXmlTheme, whose
// default minorFont latin is 'Calibri'; the notesStyle level colour is `schemeClr tx1`
// which maps (clrMap tx1→dk1) to the default theme dk1 (black, 000000).

import { describe, test } from 'vitest'
import { authorRead, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

// From the writer's own notesMaster / theme constants, not the reader under test.
const NOTES_STYLE_SIZE_PT = 12 // p:notesStyle lvl defRPr @sz=1200 -> /100
const NOTES_MINOR_FACE = 'Calibri' // +mn-lt -> theme2 minorFont latin (default bodyFontFace)
const NOTES_STYLE_COLOR = '000000' // schemeClr tx1 -> clrMap dk1 -> default theme dk1

/** The first notes-body run of the first slide. */
function firstNotesRun(presentation) {
	const run = presentation.slides[0]?.notesSlide?.body?.textFrame?.paragraphs[0]?.runs[0]
	assert(run, 'the authored notes body run reads back')
	return run
}

describe('read: notes-body inherited run properties (F2)', () => {
	test('a plain notes run inherits size/face from the notesMaster p:notesStyle', async () => {
		const { presentation } = await authorRead((pres) => {
			// No run options -> the writer emits no @sz and no <a:latin>, so both are inherited.
			pres.addSlide().addNotes('plain note text')
		})

		const run = firstNotesRun(presentation)

		// Own-attribute getters read null -- proving the resolved values below are inherited.
		assertEqual(run.fontSizePt, null, 'the run carries no own size')
		assertEqual(run.fontName, null, 'the run carries no own face')

		assertEqual(run.resolvedSizePt, NOTES_STYLE_SIZE_PT, 'size resolves from the notesStyle level defRPr')
		assertEqual(run.resolvedFontFace, NOTES_MINOR_FACE, '+mn-lt resolves through theme2 minorFont to Calibri')
	})

	test('a plain notes run inherits its colour from the notesStyle level (tx1 -> theme dk1)', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes('plain note text')
		})

		const run = firstNotesRun(presentation)
		assertEqual(run.color, null, 'the run carries no own hex')
		assertEqual(run.schemeColor, null, 'the run carries no own scheme token')

		assert(run.resolvedColor, 'the inherited notesStyle colour now resolves (was null)')
		assertEqual(run.resolvedColor.effectiveHex.toUpperCase(), NOTES_STYLE_COLOR, 'tx1 maps to the default theme dk1')
	})

	test('the notesStyle defines no bold, so an unstyled run resolves bold to null', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes('plain note text')
		})

		const run = firstNotesRun(presentation)
		assertEqual(run.bold, null, 'the run carries no own bold')
		// The notesStyle level defRPr has no @b, so nothing in the chain defines it.
		assertEqual(run.resolvedBold, null, 'inherited bold is null when the notesStyle omits @b')
	})

	test("a run's own size/face/bold win over the inherited notesStyle (negative control)", async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addNotes([{ text: 'styled', options: { fontSize: 20, fontFace: 'Georgia', bold: true } }])
		})

		const run = firstNotesRun(presentation)
		// Own attributes are set, so resolved* must report them, not the notesStyle defaults.
		assertEqual(run.fontSizePt, 20, 'the run sets its own size')
		assertEqual(run.resolvedSizePt, 20, 'own size wins over the notesStyle 12pt')
		assertEqual(run.resolvedFontFace, 'Georgia', 'own face wins over +mn-lt/Calibri')
		assertEqual(run.resolvedBold, true, 'own bold wins over the notesStyle (which defines none)')
	})

	test.skipIf(!validatorInstalled)('the authored inherited-notes deck is schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addNotes('plain note text')
		})
		assertEqual((await schemaErrors(buf)).length, 0, 'inherited-notes deck validates')
	})
})
