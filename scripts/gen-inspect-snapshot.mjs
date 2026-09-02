#!/usr/bin/env node
// Regenerate test/read/fixtures/inspect-surface.snapshot.json — the recorded
// output of `inspectPptx()` over every `.pptx`/`.potx` in test/read/fixtures.
//
// This is a CHARACTERIZATION snapshot, not an oracle. The `*.oracle.json` files
// beside it hold what PowerPoint produced; this holds what *we* produce, so
// regenerating it accepts whatever the code does today. Read the diff before
// committing one: a changed line is a changed answer to "what does this deck
// contain?", and only you can say whether that was the point of the change.
//
// Formatting is deliberate — one element per line — so `git diff` names the
// element that moved rather than reflowing a pretty-printed block around it.
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { FIXTURES_DIR, corpusDecks, isMain } from './script-utils.mjs'
import { inspectPptx } from '../dist/inspect.js'
import { setDiagnosticHandler } from '../dist/node.js'

export const SNAPSHOT_PATH = join(FIXTURES_DIR, 'inspect-surface.snapshot.json')

/**
 * Every fixture deck, by filename, in a stable order.
 *
 * `.potx` as well as `.pptx`, which is this corpus's one real divergence from every other
 * enumeration of it: `template.potx` is in the inspect snapshot and nobody else's corpus. That
 * is deliberate -- a template is a deck `inspectPptx` should describe -- so it is spelled as an
 * argument rather than as a private filter.
 */
export function fixtureDecks() {
	return corpusDecks({ extensions: ['.pptx', '.potx'] })
}

/**
 * Inspect one deck, recording the diagnostics it emitted alongside the result.
 * Only the diagnostic `code` is kept: the message is explicitly not API
 * (docs/diagnostics.md), so recording it would make a reworded sentence look
 * like a behaviour change.
 *
 * @typedef {import('../dist/inspect.js').PptxInspection} Inspection
 * @typedef {{slideSize: Inspection['slideSize'], slides: Inspection['slides'], diagnostics: string[]}} OkRecord
 * @typedef {{failed: {error: string, code: string | null}, diagnostics: string[]}} FailedRecord
 * @typedef {OkRecord | FailedRecord} DeckRecord
 *
 * @param {string} deck fixture filename
 * @returns {Promise<DeckRecord>}
 */
export async function inspectForSnapshot(deck) {
	/** @type {string[]} */
	const diagnostics = []
	setDiagnosticHandler((d) => diagnostics.push(d.code))
	try {
		const { slideSize, slides } = await inspectPptx(join(FIXTURES_DIR, deck))
		return { slideSize, slides, diagnostics }
	} catch (err) {
		// A deck inspect cannot read at all is itself part of the contract. Record the
		// class and its code — the two halves of the taxonomy that are API — not the
		// message, for the same reason diagnostics keep only their code.
		const failure = /** @type {{ constructor?: { name?: string }; code?: string }} */ (err)
		return { failed: { error: failure?.constructor?.name ?? 'unknown', code: failure?.code ?? null }, diagnostics }
	} finally {
		setDiagnosticHandler(null)
	}
}

/**
 * Render the snapshot: pretty-printed down to the element level, one line per element.
 * @param {Record<string, DeckRecord>} byDeck
 * @returns {string}
 */
export function formatSnapshot(byDeck) {
	const lines = ['{']
	const decks = Object.keys(byDeck)
	decks.forEach((deck, deckIndex) => {
		const record = byDeck[deck]
		if (!record) return
		lines.push(`\t${JSON.stringify(deck)}: {`)
		if ('failed' in record) lines.push(`\t\t"failed": ${JSON.stringify(record.failed)},`)
		else {
			lines.push(`\t\t"slideSize": ${JSON.stringify(record.slideSize)},`)
			lines.push('\t\t"slides": [')
			record.slides.forEach((slide, slideIndex) => {
				const { elements, ...head } = slide
				lines.push('\t\t\t{')
				// `elements` is rendered last regardless of where it sits in the object
				// so the long lines stay together; JSON object key order is not data.
				for (const [key, value] of Object.entries(head)) {
					lines.push(`\t\t\t\t${JSON.stringify(key)}: ${JSON.stringify(value)},`)
				}
				lines.push('\t\t\t\t"elements": [')
				elements.forEach((element, elementIndex) => {
					lines.push(`\t\t\t\t\t${JSON.stringify(element)}${elementIndex < elements.length - 1 ? ',' : ''}`)
				})
				lines.push('\t\t\t\t]')
				lines.push(`\t\t\t}${slideIndex < record.slides.length - 1 ? ',' : ''}`)
			})
			lines.push('\t\t],')
		}
		lines.push(`\t\t"diagnostics": ${JSON.stringify(record.diagnostics)}`)
		lines.push(`\t}${deckIndex < decks.length - 1 ? ',' : ''}`)
	})
	lines.push('}')
	return lines.join('\n') + '\n'
}

/**
 * Inspect every fixture deck, keyed by filename.
 * @returns {Promise<Record<string, DeckRecord>>}
 */
export async function buildSnapshot() {
	/** @type {Record<string, DeckRecord>} */
	const byDeck = {}
	for (const deck of await fixtureDecks()) byDeck[deck] = await inspectForSnapshot(deck)
	return byDeck
}

// Only write when run directly; the test imports the helpers above to rebuild
// the same structure in memory and compare.
if (isMain(import.meta.url)) {
	const byDeck = await buildSnapshot()
	writeFileSync(SNAPSHOT_PATH, formatSnapshot(byDeck))
	const elements = Object.values(byDeck)
		.flatMap((record) => ('failed' in record ? [] : record.slides))
		.reduce((n, slide) => n + slide.elements.length, 0)
	console.log(`wrote ${SNAPSHOT_PATH}: ${Object.keys(byDeck).length} decks, ${elements} elements`)
}
