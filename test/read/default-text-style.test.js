// Read-model coverage for the two lowest run-resolution tiers, exercised on a
// minimal desktop-PowerPoint fixture (default-text-style.pptx):
//
//   PlainBox   — a plain text box (no placeholder, no p:style) whose only run is
//                bare (`<a:rPr lang="en-US"/>`). Its size/colour/face therefore come
//                entirely from the presentation's p:defaultTextStyle — PowerPoint's
//                lowest-priority text fallback. Before defaultTextStyle joined the
//                chain every one of these read as null.
//   StyledRect — a rectangle carrying a theme shape style, so it has a
//                `p:style/a:fontRef idx="minor"` with an `a:schemeClr val="lt1"`.
//                Its bare run's colour + face resolve through that fontRef (the tier
//                just above the placeholder/defaultTextStyle chain); its size, which
//                a fontRef never carries, still falls through to defaultTextStyle.
//
// Ground truth read out of the fixture: p:defaultTextStyle lvl1 is sz=1800,
// schemeClr tx1, latin +mn-lt; the theme (default Office) has minorFont "Aptos",
// clrMap tx1->dk1 (windowText = 000000) and the direct slot lt1 = window (FFFFFF).

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function open(name) {
	return Presentation.load(await readFile(path.join(__dirname, 'fixtures', `${name}.pptx`)))
}

function runOf(slide, shapeName) {
	const shape = slide.shapes.find((s) => s.name === shapeName)
	assert(shape, `expected a shape named ${shapeName}`)
	return shape.textFrame.paragraphs[0].runs[0]
}

describe('p:defaultTextStyle is the run-resolution fallback (default-text-style.pptx)', () => {
	test('a bare non-placeholder run resolves size/colour/face from p:defaultTextStyle', async () => {
		const run = runOf((await open('default-text-style')).slides[0], 'PlainBox')
		// The run itself sets nothing — every resolved value is the default text style's.
		assertEqual(run.fontSizePt, null, 'the run sets no own @sz')
		assertEqual(run.fontName, null, 'the run sets no own a:latin')
		assertEqual(run.color, null, 'the run sets no own colour')
		assertEqual(run.resolvedSizePt, 18, 'inherits p:defaultTextStyle sz=1800 → 18pt')
		assertEqual(run.resolvedFontFace, 'Aptos', 'inherits +mn-lt → the theme minor font')
		assert(run.resolvedColor, 'the run resolves a colour')
		assertEqual(run.resolvedColor.effectiveHex, '000000', 'inherits schemeClr tx1 → dk1 windowText')
	})

	test('the default text style sets no bold, so resolvedBold stays null', async () => {
		const run = runOf((await open('default-text-style')).slides[0], 'PlainBox')
		// The default tier must not fabricate a bold value it does not carry — the
		// distinction between an inherited false and "nothing inherits" is preserved.
		assertEqual(run.bold, null, 'the run sets no own @b')
		assertEqual(run.resolvedBold, null, 'p:defaultTextStyle carries no @b')
	})
})

describe('p:style/a:fontRef supplies run colour + face (default-text-style.pptx)', () => {
	test("a styled shape's bare run takes its colour + face from the fontRef", async () => {
		const run = runOf((await open('default-text-style')).slides[0], 'StyledRect')
		assertEqual(run.color, null, 'the run sets no own colour')
		assertEqual(run.fontName, null, 'the run sets no own a:latin')
		assert(run.resolvedColor, 'the run resolves a colour')
		// The fontRef names schemeClr lt1 (white). If the fontRef tier were skipped the
		// run would fall through to p:defaultTextStyle's tx1 (black) — so FFFFFF here is
		// the decisive proof the fontRef colour is consulted and wins over the default.
		assertEqual(run.resolvedColor.effectiveHex, 'FFFFFF', 'fontRef schemeClr lt1 → window white')
		assertEqual(run.resolvedFontFace, 'Aptos', 'fontRef idx="minor" → the theme minor font')
	})

	test('a fontRef carries no size, so size still falls through to the default text style', async () => {
		const run = runOf((await open('default-text-style')).slides[0], 'StyledRect')
		assertEqual(run.fontSizePt, null, 'the run sets no own @sz')
		assertEqual(run.resolvedSizePt, 18, 'no size on the fontRef → p:defaultTextStyle sz=1800')
	})

	test("a run's own colour and face still win over the fontRef", async () => {
		const run = runOf((await open('default-text-style')).slides[0], 'StyledRect')
		// Setting own values (in memory) must override the fontRef tier below them.
		run.color = 'FF0000'
		run.fontName = 'Verdana'
		assert(run.resolvedColor, 'own colour resolves')
		assertEqual(run.resolvedColor.effectiveHex, 'FF0000', "the run's own colour beats the fontRef")
		assertEqual(run.resolvedFontFace, 'Verdana', "the run's own face beats the fontRef")
	})
})
