// Write→read fidelity for the run/paragraph formatting reads added to
// src/read/api/text.ts: strikethrough / caps / baseline (sub-super) / highlight /
// hyperlink on Run, and line spacing on Paragraph. Each is a BLIND SPOT the writer
// already emits from the addText run options — so it is proven by authoring a deck
// with the write API, reading it back through the deep model, and asserting the
// extracted values match what was written. The writer's bytes are the fixture.

import { describe, test } from 'vitest'
import { authorRead, firstShape, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** A text box whose runs each carry one character-formatting property. */
function formattedRuns(pres) {
	pres.addSlide().addText(
		[
			{ text: 'Struck', options: { strike: true } },
			{ text: 'Double', options: { strike: 'dblStrike' } },
			{ text: 'AllCaps', options: { caps: 'all' } },
			{ text: 'Super', options: { superscript: true } },
			{ text: 'Sub', options: { subscript: true } },
			{ text: 'Marked', options: { highlight: 'FFFF00' } },
			{ text: 'Link', options: { hyperlink: { url: 'https://example.com', tooltip: 'Homepage' } } },
			{ text: 'Plain', options: {} },
		],
		{ x: 1, y: 1, w: 8, h: 3 }
	)
}

/** Every run of the first text-bearing shape, flattened across paragraphs. */
function runsOf(presentation) {
	const shape = firstShape(presentation, (s) => s.hasTextFrame)
	assert(shape, 'authored text box is read back')
	return shape.textFrame.paragraphs.flatMap((p) => p.runs)
}

/** The run whose text is exactly `text`. */
function run(runs, text) {
	const found = runs.find((r) => r.text === text)
	assert(found, `run "${text}" is present`)
	return found
}

describe('Run character formatting — a:rPr attributes', () => {
	test('strike reads its token (single and double), null when unset', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		assertEqual(run(runs, 'Struck').strike, 'sngStrike', 'strike:true → sngStrike')
		assertEqual(run(runs, 'Double').strike, 'dblStrike', 'strike:"dblStrike" round-trips')
		assertEqual(run(runs, 'Plain').strike, null, 'an unstruck run reports null')
	})

	test('caps reads its token', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		assertEqual(run(runs, 'AllCaps').caps, 'all', 'caps:"all" round-trips')
		assertEqual(run(runs, 'Plain').caps, null, 'a run with no caps reports null')
	})

	test('baseline reads super/subscript as a percentage', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		assertEqual(run(runs, 'Super').baselinePct, 30, 'superscript → +30%')
		assertEqual(run(runs, 'Sub').baselinePct, -40, 'subscript → -40%')
		assertEqual(run(runs, 'Plain').baselinePct, null, 'a baseline-shifted-none run reports null')
	})

	test('highlight resolves to a literal hex', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		const hl = run(runs, 'Marked').highlight
		assert(hl, 'a highlighted run exposes a highlight colour')
		assertEqual(hl.effectiveHex, 'FFFF00', 'authored highlight hex')
		assertEqual(run(runs, 'Plain').highlight, null, 'an unhighlighted run reports null')
	})
})

describe('Run hyperlink — a:hlinkClick', () => {
	test('a URL link resolves its r:id to the external target', async () => {
		const runs = runsOf((await authorRead(formattedRuns)).presentation)
		const link = run(runs, 'Link').hyperlink
		assert(link, 'a linked run exposes a hyperlink')
		assertEqual(link.url, 'https://example.com', 'r:id resolved to the external URL')
		assertEqual(link.targetPartName, null, 'an external link has no internal target part')
		assertEqual(link.tooltip, 'Homepage', 'authored tooltip')
		assert(link.relId, 'the backing relationship id is surfaced')
		assertEqual(run(runs, 'Plain').hyperlink, null, 'a run with no link reports null')
	})

	test('a slide jump resolves its r:id to the linked slide part', async () => {
		const chart = (
			await authorRead((pres) => {
				const first = pres.addSlide()
				pres.addSlide() // slide 2 — the jump target
				first.addText([{ text: 'Jump', options: { hyperlink: { slide: 2, tooltip: 'Go to 2' } } }], {
					x: 1,
					y: 1,
					w: 6,
					h: 2,
				})
			})
		).presentation
		const link = run(runsOf(chart), 'Jump').hyperlink
		assert(link, 'a jump run exposes a hyperlink')
		assertEqual(link.url, null, 'an internal jump has no external URL')
		assertEqual(link.action, 'ppaction://hlinksldjump', 'slide jump action token')
		assert(link.targetPartName?.endsWith('slide2.xml'), `jump resolves to slide 2 (${link.targetPartName})`)
	})
})

describe('Paragraph line spacing — a:lnSpc', () => {
	test('a percentage multiple reads as percent', async () => {
		const presentation = (
			await authorRead((pres) => {
				pres.addSlide().addText('Spaced', { x: 1, y: 1, w: 6, h: 2, lineSpacingMultiple: 1.5 })
			})
		).presentation
		const shape = firstShape(presentation, (s) => s.hasTextFrame)
		const spacing = shape.textFrame.paragraphs[0].lineSpacing
		assert(spacing, 'the paragraph exposes line spacing')
		assertEqual(spacing.type, 'percent', 'lineSpacingMultiple → percent form')
		assertEqual(spacing.percent, 150, '1.5× → 150%')
	})

	test('an exact point height reads as points', async () => {
		const presentation = (
			await authorRead((pres) => {
				pres.addSlide().addText('Spaced', { x: 1, y: 1, w: 6, h: 2, lineSpacing: 24 })
			})
		).presentation
		const shape = firstShape(presentation, (s) => s.hasTextFrame)
		const spacing = shape.textFrame.paragraphs[0].lineSpacing
		assert(spacing, 'the paragraph exposes line spacing')
		assertEqual(spacing.type, 'points', 'lineSpacing → points form')
		assertEqual(spacing.valuePt, 24, 'authored point height')
	})

	test('a paragraph with default spacing reports null', async () => {
		const presentation = (
			await authorRead((pres) => {
				pres.addSlide().addText('Plain', { x: 1, y: 1, w: 6, h: 2 })
			})
		).presentation
		const shape = firstShape(presentation, (s) => s.hasTextFrame)
		assertEqual(shape.textFrame.paragraphs[0].lineSpacing, null, 'no a:lnSpc → null')
	})
})

describe('Run/paragraph formatting — schema validity', () => {
	test.skipIf(!validatorInstalled)('the authored formatted deck is schema-valid', async () => {
		assertEqual((await schemaErrors((await authorRead(formattedRuns)).buf)).length, 0, 'formatted-runs deck validates')
	})
})
