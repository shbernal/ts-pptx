// Edge coverage for src/read/api/slide-background.ts — the `p:bg` shapes the
// writer never authors.
//
// slide-read-edges.test.js covers the three backgrounds the write API emits
// (solid / gradient / image) as a measured write→read round-trip. The decoder
// also handles variants that only reach it from an *imported* deck: `a:pattFill`,
// an explicit `a:noFill`, a `p:bg` carrying neither `p:bgPr` nor `p:bgRef`, a
// `p:bgRef` whose `idx` has no `fmtScheme` entry. Those have no writer to author
// them, so each is patched into an authored deck's `slide1.xml` as the XML an
// imported deck would carry — the same synthetic-input approach
// chart-parse-edge.test.js uses for `c:chartSpace`.
//
// The one case with a genuine PowerPoint oracle — a stock theme whose
// `bgFillStyleLst` third slot is a picture — is asserted against theme-colors.pptx
// (the Ion theme) rather than synthesized, since it pins whose relationships a
// theme-materialized fill resolves against.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

/**
 * Author a one-slide deck, splice `bgXml` in as the slide's `p:cSld/p:bg` (first
 * child, per CT_CommonSlideData's sequence), and return the reloaded
 * `slide.background`. The rest of the package — theme, layout, master, rels — is
 * the writer's own, so colour tokens and `fmtScheme` lookups resolve for real.
 */
async function backgroundFrom(bgXml) {
	const { buf } = await authorRead((pres) => {
		pres.addSlide()
	})
	const zip = await JSZip.loadAsync(buf)
	const slideXml = await zip.file('ppt/slides/slide1.xml').async('string')
	assert(!slideXml.includes('<p:bg'), 'the authored slide carries no p:bg of its own to collide with')
	zip.file('ppt/slides/slide1.xml', slideXml.replace(/(<p:cSld[^>]*>)/, `$1${bgXml}`))
	const patched = await zip.generateAsync({ type: 'uint8array' })
	return (await Presentation.load(patched)).slides[0].background
}

describe('slide background — imported-only p:bg variants', () => {
	test('a pattFill reads its preset plus both resolved colours', async () => {
		const bg = await backgroundFrom(
			'<p:bg><p:bgPr><a:pattFill prst="lgCheck">' +
				'<a:fgClr><a:srgbClr val="C00000"/></a:fgClr>' +
				'<a:bgClr><a:schemeClr val="bg1"/></a:bgClr>' +
				'</a:pattFill><a:effectLst/></p:bgPr></p:bg>'
		)
		assert(bg.type === 'pattern', 'a:pattFill decodes to the pattern variant')
		assertEqual(bg.source, 'slide', 'authored on the slide itself')
		assertEqual(bg.preset, 'lgCheck', 'the ST_PresetPatternVal token is kept verbatim')
		assertEqual(bg.foreground?.effectiveHex, 'C00000', 'the literal fgClr resolves')
		// bgClr is a token, so it exercises the colour-map + scheme walk, not just a literal.
		assertEqual(bg.background?.effectiveHex, 'FFFFFF', 'the bgClr schemeClr resolves through bg1 → lt1 → white')
	})

	test('a pattFill with no prst and no colour wrappers reads all-null, not undefined', async () => {
		const bg = await backgroundFrom('<p:bg><p:bgPr><a:pattFill/></p:bgPr></p:bg>')
		assert(bg.type === 'pattern', 'still the pattern variant')
		assertEqual(bg.preset, null, 'a missing @prst reads null')
		assertEqual(bg.foreground, null, 'a missing a:fgClr reads null')
		assertEqual(bg.background, null, 'a missing a:bgClr reads null')
	})

	test('an explicit a:noFill is a transparent background, not an absent one', async () => {
		const bg = await backgroundFrom('<p:bg><p:bgPr><a:noFill/><a:effectLst/></p:bgPr></p:bg>')
		// The distinction that matters: `type: 'none'` (the slide states it has no fill,
		// stopping inheritance) versus `background === null` (nothing in the chain says).
		assert(bg !== null, 'an explicit noFill is a background, so the chain stops at the slide')
		assertEqual(bg.type, 'none', 'a:noFill decodes to none')
		assertEqual(bg.source, 'slide', 'sourced from the slide')
	})

	test('a bgPr holding no recognized fill falls back to none', async () => {
		const bg = await backgroundFrom('<p:bg><p:bgPr><a:effectLst/></p:bgPr></p:bg>')
		assertEqual(bg.type, 'none', 'an effect-only bgPr has no fill to decode')
	})

	test('a p:bg carrying neither p:bgPr nor p:bgRef reads none', async () => {
		const bg = await backgroundFrom('<p:bg/>')
		assertEqual(bg.type, 'none', 'an empty p:bg still stops the inheritance chain')
		assertEqual(bg.source, 'slide', 'sourced from the slide')
	})

	test('a blipFill with no a:blip reads image with a null rel rather than throwing', async () => {
		const bg = await backgroundFrom(
			'<p:bg><p:bgPr><a:blipFill><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>'
		)
		assert(bg.type === 'image', 'the blipFill still classifies as an image background')
		assertEqual(bg.relId, null, 'no a:blip → no rel id')
		assertEqual(bg.partName, null, 'and so no part to resolve')
	})

	test('a bgRef whose idx has no fmtScheme entry keeps the raw idx and resolves to null', async () => {
		// The Office fmtScheme carries three bgFillStyleLst entries (1001–1003); 1009 is
		// the degraded-import case where the index points past them.
		const bg = await backgroundFrom('<p:bg><p:bgRef idx="1009"><a:schemeClr val="bg1"/></p:bgRef></p:bg>')
		assert(bg.type === 'themeRef', 'still a theme-indexed background')
		assertEqual(bg.idx, 1009, 'the raw idx is kept for fidelity even when unresolvable')
		assertEqual(bg.color?.effectiveHex, 'FFFFFF', "the bgRef's own colour still resolves")
		assertEqual(bg.resolvedFill, null, 'no matching fmtScheme entry → no concrete fill')
	})

	test('a bgRef into a gradient fmtScheme entry resolves the phClr into every stop', async () => {
		const bg = await backgroundFrom('<p:bg><p:bgRef idx="1003"><a:schemeClr val="accent1"/></p:bgRef></p:bg>')
		assert(bg.type === 'themeRef', 'theme-indexed background')
		assert(bg.resolvedFill?.type === 'gradient', 'the third Office bg fill-style entry is a gradient')
		const stops = bg.resolvedFill.gradient.stops
		assertEqual(stops.length, 3, 'all three stops decode')
		// Every stop is `phClr` + its own shade/tint, so each resolves off accent1 (4472C4)
		// to a *different* hex — proving the substitution happened before the transforms.
		assertEqual(stops[0].color, '4472C4', 'the phClr was substituted with the resolved accent1')
		assertEqual(new Set(stops.map((s) => s.effectiveHex)).size, 3, 'each stop applies its own transforms')
	})
})

describe('slide background — a theme-materialized image fill resolves against the theme part', () => {
	// theme-colors.pptx uses the stock Ion theme, whose third bgFillStyleLst entry is
	// an `a:blipFill` with `r:embed="rId1"` — and that id lives in the *theme* part's
	// rels (→ /ppt/media/image1.jpeg). Resolving it against the owning master's rels
	// instead silently returns whatever rId1 means there (a slide layout), so this
	// pins whose relationship table a materialized fill belongs to.
	async function ionDeck() {
		return Presentation.load(await readFile(path.join(__dirname, 'fixtures', 'theme-colors.pptx')))
	}

	test("the master's picture background resolves to the theme's media part", async () => {
		const bg = (await ionDeck()).masters()[0].background
		assert(bg?.type === 'themeRef', 'the Ion master authors p:bgRef idx=1003')
		assertEqual(bg.idx, 1003, 'the picture slot of the bg fill-style list')
		assert(bg.resolvedFill?.type === 'image', 'that slot materializes to a blipFill')
		assertEqual(bg.resolvedFill.relId, 'rId1', "the blip's r:embed, verbatim")
		assertEqual(
			bg.resolvedFill.partName,
			'/ppt/media/image1.jpeg',
			"rId1 resolves through the theme part's rels, not the master's"
		)
	})

	test('a slide inheriting that background resolves the same media part', async () => {
		const bg = (await ionDeck()).slides[0].background
		assert(bg?.type === 'themeRef', 'the slide inherits the theme-indexed background')
		assert(bg.resolvedFill?.type === 'image', 'and the same picture fill')
		assertEqual(bg.resolvedFill.partName, '/ppt/media/image1.jpeg', 'inherited backgrounds resolve identically')
	})
})
