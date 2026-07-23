// Write→read fidelity for the shared-chrome property model (src/read/api/chrome.ts).
//
// T2.2 models a deck's masters, layouts, and themes as typed read objects. The
// write API authors all three: `pres.theme = { colorScheme, headFontFace, ... }`
// emits `theme1.xml` (the 12-slot `a:clrScheme` + the major/minor `a:fontScheme`),
// and `defineSlideMaster({ title, background, slideNumber, objects })` emits a
// `slideMaster1.xml` (its `p:clrMap` + `p:txStyles` + placeholder shapes) and a
// `slideLayoutN.xml` (its `p:cSld@name`, background, and placeholders). So the read
// model has a genuine measured round-trip rather than a hand-authored fixture.
//
// The chain is navigable both from the deck (`pres.masters()[i]`) and from a slide
// (`slide.layout.master.theme`). What round-trips: theme colour slots (an `a:sysClr`
// slot via its `lastClr`), the major/minor Latin faces, the master colour map, the
// master/layout placeholder *geometry* (authored explicitly on master/layout, unlike
// a notes placeholder's inherited-and-null geometry), and layout backgrounds. The
// one import-only surface is `p:sldLayout@type` — the writer authors none, so it
// reads `null` on an authored deck (asserted null here, documented, not faked).

import { describe, test } from 'vitest'
import { authorRead, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** The first slide of `presentation`. */
function firstSlide(presentation) {
	const slide = presentation.slides[0]
	assert(slide, 'the authored slide is read back')
	return slide
}

/**
 * Author a deck with a customized theme + a defined master, and return it read back.
 * The theme overrides accent1/dk2 (the rest keep the Office defaults) and sets the
 * heading/body faces; the master `BRANDED` carries a solid background, a rect (a
 * non-placeholder decoration), and a slide-number placeholder with explicit geometry.
 */
function authorBrandedDeck() {
	return authorRead((pres) => {
		pres.theme = { headFontFace: 'Georgia', bodyFontFace: 'Verdana', colorScheme: { accent1: '112233', dk2: 'AABBCC' } }
		pres.defineSlideMaster({
			title: 'BRANDED',
			background: { color: 'F1F2F3' },
			objects: [{ rect: { x: 0, y: 0, w: 1, h: 1, fill: { color: '00FF00' } } }],
			slideNumber: { x: 0.5, y: 7.0 },
		})
		pres.addSlide({ masterName: 'BRANDED' }).addText('hi', { x: 1, y: 1, w: 3, h: 1 })
	})
}

describe('Theme / SlideMaster / SlideLayout — write→read fidelity', () => {
	test('the theme colour scheme resolves each slot to a literal hex (T2.2)', async () => {
		const { presentation } = await authorBrandedDeck()
		const theme = firstSlide(presentation).theme
		assert(theme, 'the slide resolves its theme through layout → master → theme')

		assertEqual(theme.name, 'Office Theme', 'the theme name round-trips')
		assertEqual(theme.colorSchemeName, 'Office', 'the colour-scheme name round-trips')

		// Overrides win; un-overridden slots keep the Office default; a sysClr slot
		// resolves through its lastClr.
		assertEqual(theme.color('accent1'), '112233', 'the accent1 override resolves')
		assertEqual(theme.colorScheme.dk2, 'AABBCC', 'the dk2 override resolves')
		assertEqual(theme.colorScheme.accent2, 'ED7D31', 'an un-overridden slot keeps the Office default')
		assertEqual(theme.colorScheme.dk1, '000000', 'the dk1 sysClr resolves via lastClr')
		assertEqual(theme.colorScheme.lt1, 'FFFFFF', 'the lt1 sysClr resolves via lastClr')
	})

	test('the theme font scheme reads the major/minor faces (T2.2)', async () => {
		const { presentation } = await authorBrandedDeck()
		const fontScheme = firstSlide(presentation).theme.fontScheme
		assert(fontScheme, 'the theme declares a font scheme')

		assertEqual(fontScheme.name, 'Office', 'the font-scheme name round-trips')
		assertEqual(fontScheme.major.latin, 'Georgia', 'the heading (major) Latin face round-trips')
		assertEqual(fontScheme.minor.latin, 'Verdana', 'the body (minor) Latin face round-trips')
		// PowerPoint authors the ea/cs slots empty by default; empty string reads as null.
		assertEqual(fontScheme.major.ea, null, 'an empty ea slot reads null, not ""')
		assertEqual(fontScheme.minor.cs, null, 'an empty cs slot reads null, not ""')
	})

	test('the slide master exposes its colour map and theme (T2.2)', async () => {
		const { presentation } = await authorBrandedDeck()
		const master = firstSlide(presentation).master
		assert(master, 'the slide resolves its master through its layout')

		assertEqual(master.name, '', 'the writer authors the master with no cSld name (empty string, not null)')
		// The colour map is the token → slot indirection (a slide schemeClr val="tx1" → dk1).
		assertEqual(master.colorMap.tx1, 'dk1', 'tx1 maps to the dk1 slot')
		assertEqual(master.colorMap.bg1, 'lt1', 'bg1 maps to the lt1 slot')
		assertEqual(master.colorMap.accent1, 'accent1', 'accent1 maps to its own slot')

		assertEqual(master.theme?.name, 'Office Theme', 'the master resolves the theme part')
		assertEqual(
			master.layouts
				.map((l) => l.name)
				.sort()
				.join(','),
			'BRANDED,DEFAULT',
			'the master lists its layouts'
		)
	})

	test('master and layout placeholders carry their own geometry (T2.2)', async () => {
		const { presentation } = await authorBrandedDeck()
		const master = firstSlide(presentation).master

		// The master's slide-number placeholder carries explicit EMU geometry —
		// authored on the master, so it round-trips (unlike a notes placeholder's null).
		const sldNum = master.placeholders.find((ph) => ph.type === 'sldNum')
		assert(sldNum, 'the master has a slide-number placeholder')
		assertEqual(sldNum.left, 457200, 'placeholder left round-trips (0.5in)')
		assertEqual(sldNum.top, 6400800, 'placeholder top round-trips (7.0in)')
		assertEqual(sldNum.width, 800000, 'placeholder width round-trips')
		assertEqual(sldNum.height, 300000, 'placeholder height round-trips')

		// The BRANDED layout carries the same placeholder; its decorative rect (no p:ph)
		// is filtered out, so `placeholders` lists only the true placeholder.
		const branded = master.layouts.find((l) => l.name === 'BRANDED')
		assert(branded, 'the BRANDED layout reads back')
		assertEqual(branded.placeholders.length, 1, 'the non-placeholder rect is excluded from placeholders')
		assertEqual(branded.placeholders[0].type, 'sldNum', 'the one layout placeholder is the slide-number one')
		assertEqual(branded.placeholders[0].left, 457200, 'the layout placeholder geometry round-trips too')
	})

	test('a layout exposes its name, its (import-only) type, and its own background (T2.2)', async () => {
		const { presentation } = await authorBrandedDeck()
		const master = firstSlide(presentation).master

		const branded = master.layouts.find((l) => l.name === 'BRANDED')
		// @type is import-only — the writer authors none, so it reads null (not faked).
		assertEqual(branded.type, null, 'the writer authors no layout @type (import-only surface)')
		assertEqual(branded.master?.name, master.name, 'the layout resolves back to its master')

		const bg = branded.background
		assertEqual(bg?.type, 'solid', 'the BRANDED layout authors a solid background')
		assertEqual(bg?.source, 'layout', 'the background is sourced from the layout itself')
		assertEqual(bg?.color?.effectiveHex, 'F1F2F3', 'the background colour round-trips')

		// The default layout authors a theme-indexed background instead.
		const def = master.layouts.find((l) => l.name === 'DEFAULT')
		const defBg = def.background
		assertEqual(defBg?.type, 'themeRef', 'the DEFAULT layout authors a theme-indexed background')
		assertEqual(defBg?.idx, 1001, 'the bgRef index round-trips')
		assertEqual(defBg?.color?.hex, 'FFFFFF', 'bg1 resolves through the colour map + scheme to white')
	})

	test('the chrome is navigable from both the deck and a slide (T2.2)', async () => {
		const { presentation } = await authorBrandedDeck()

		// Deck-level enumeration and slide-level navigation agree on the same master.
		const masters = presentation.masters()
		assertEqual(masters.length, 1, 'the deck has one master')
		const slide = firstSlide(presentation)
		assertEqual(slide.master?.partName, masters[0].partName, 'slide.master is the deck master')

		// The full walk slide → layout → master → theme resolves the branded palette.
		assert(slide.layout, 'the slide binds to a layout')
		assertEqual(slide.theme?.color('accent1'), '112233', 'slide.theme resolves the branded accent1')
		assertEqual(slide.layout.master?.theme?.partName, slide.theme?.partName, 'both paths reach the same theme part')
	})

	test.skipIf(!validatorInstalled)('the authored branded deck is schema-valid', async () => {
		const { buf } = await authorBrandedDeck()
		assertEqual((await schemaErrors(buf)).length, 0, 'branded deck validates')
	})
})
