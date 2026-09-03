// Every emitted element's children are in the order its complexType declares.
//
// The read path inserts into a live DOM and asks `ooxml/sequence.ts` where a child goes. The
// write path has nothing equivalent: each emitter states its children in order by hand, and the
// only thing checking that order is PowerPoint, which reports a misplaced child as a corrupt
// file rather than as a bad option — the failure lands on a user, far from its cause, with no
// compile-time signal. That is exactly how a chart axis ended up carrying its units under the
// wrong element.
//
// So this reads the declared sequences and compares them against what the emitters actually
// write: for every element the table covers, its children must be a *subsequence* of the
// declared order (children are optional, so gaps are fine; reordering is not) and every child
// must be a member of the sequence at all.
//
// Its reach is the deck below, which is the honest limit — an emitter no fixture here exercises
// is not checked. `sequence.ts` is imported from `src/`: it is an internal declaration table
// with no published entry point, and the assertion is about the table itself.

import { describe, test } from 'vitest'
import { DOMParser } from '@xmldom/xmldom'
import JSZip from 'jszip'
import { CHILD_SEQUENCES } from '../../../src/ooxml/sequence.ts'
import { TsPptx, ChartType } from '../../../dist/node.js'
import { assert } from '../../helpers.js'

const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='

/** Flatten a declared sequence into slots: a choice group's members all occupy one slot. */
function slotsOf(sequence) {
	return sequence.map((step) => (typeof step === 'string' ? [step] : [...step]))
}

/**
 * The first violation of `sequence` in `children`, or `null`.
 *
 * Walks the declared slots forward, never back: a child that matches an earlier slot than one
 * already passed is out of order, and a child in no slot at all is not a member of the type.
 */
function outOfOrder(children, sequence) {
	const slots = slotsOf(sequence)
	let at = 0
	for (const [index, child] of children.entries()) {
		const slot = slots.findIndex((names) => names.includes(child))
		if (slot < 0) return `${child} is not a child of this type at all`
		if (slot < at) return `${child} comes after ${children[index - 1]}, but the schema puts it before`
		// A repeated element stays in its own slot; anything else advances past it.
		at = slot
	}
	return null
}

/** Every element in `xml` whose name the sequence table covers, with its child element names. */
function coveredElements(xml) {
	const doc = new DOMParser().parseFromString(xml, 'text/xml')
	const found = []
	const visit = (node) => {
		for (let child = node.firstChild; child; child = child.nextSibling) {
			if (child.nodeType !== 1) continue
			const sequence = CHILD_SEQUENCES[child.nodeName]
			if (sequence) {
				const names = []
				for (let sub = child.firstChild; sub; sub = sub.nextSibling) {
					if (sub.nodeType === 1) names.push(sub.nodeName)
				}
				found.push({ name: child.nodeName, children: names, sequence })
			}
			visit(child)
		}
	}
	visit(doc)
	return found
}

/**
 * A deck reaching the emitters the table covers: shapes with every `spPr` child the write API
 * can produce, runs carrying a link and a typeface, a table with cell borders and fills, a
 * picture, a group, and charts, whose `c:spPr` is the same complexType a slide's is.
 */
async function richDeck() {
	const pres = new TsPptx()
	const slide = pres.addSlide()
	slide.addShape('rect', {
		x: 0.5,
		y: 0.5,
		w: 2,
		h: 1,
		fill: { color: 'accent1', transparency: 20 },
		line: { color: '203050', width: 2, dashType: 'dash', beginArrowType: 'triangle', endArrowType: 'oval' },
		shadow: { type: 'outer', color: '000000', blur: 3, offset: 2, angle: 45, transparency: 40 },
		rotate: 15,
	})
	slide.addText(
		[
			{ text: 'linked ', options: { hyperlink: { url: 'https://example.invalid' }, fontFace: 'Georgia' } },
			{ text: 'and not', options: { bold: true, underline: { style: 'sng' }, highlight: 'FFFF00' } },
		],
		{ x: 0.5, y: 2, w: 4, h: 1, valign: 'middle', fill: { color: 'F0F0F0' }, line: { color: '888888' } }
	)
	slide.addImage({ data: PNG_DATA, x: 5, y: 0.5, w: 1, h: 1, rounding: true })
	slide.addTable(
		[
			[
				{ text: 'a', options: { fill: { color: 'DDEEFF' }, border: { type: 'solid', color: '333333', width: 1 } } },
				{ text: 'b', options: { valign: 'bottom' } },
			],
		],
		{ x: 0.5, y: 3.2, w: 6, h: 1, border: { type: 'solid', color: '999999', width: 1 } }
	)
	slide.addGroup([
		{ rect: { x: 5, y: 2, w: 1, h: 1, fill: { color: 'CC0000' } } },
		{ text: { text: 'in a group', options: { x: 5, y: 3, w: 2, h: 0.5 } } },
	])
	// One chart per type rather than a loop over them: the chart data shape is per type, and
	// `c:spPr` -- the same complexType a slide's `p:spPr` is -- is exactly what these are here for.
	const labels = ['a', 'b', 'c']
	const axes = { catAxisLineColor: '333333', valAxisLineColor: '333333' }
	const box = { x: 0.5, y: 0.5, w: 8, h: 4 }
	pres.addSlide().addChart([{ name: 'S', labels, values: [1, 2, 3] }], {
		type: ChartType.bar,
		...box,
		...axes,
		showTitle: true,
		title: 'chart',
		showLegend: true,
		showValue: true,
	})
	pres.addSlide().addChart([{ name: 'S', labels, values: [3, 2, 1] }], { type: ChartType.line, ...box, ...axes })
	pres
		.addSlide()
		.addChart([{ name: 'S', labels, values: [1, 1, 2] }], { type: ChartType.pie, ...box, showPercent: true })
	pres.addSlide().addChart(
		[
			{ name: 'X', labels, values: [1, 2, 3] },
			{ name: 'Y', labels, values: [4, 5, 6] },
		],
		{ type: ChartType.scatter, ...box, ...axes }
	)
	return JSZip.loadAsync(await pres.toBytes())
}

/**
 * The covered element names this deck is expected to reach. Not every key of the table: a
 * `cx:spPr` needs a 2016 chart and `a:blipFill` a picture-filled surface, neither of which is
 * here. Listing what it does reach keeps the deck from quietly shrinking.
 */
const REACHED = [
	'p:presentation',
	'p:sp',
	'p:spPr',
	'p:grpSpPr',
	'p:pic',
	'p:blipFill',
	'a:ln',
	'a:rPr',
	'a:defRPr',
	'a:endParaRPr',
	'a:tcPr',
	'c:spPr',
]

describe('emitted child order follows the declared schema sequence', () => {
	test('every covered element in a deck exercising the emitters', async () => {
		const zip = await richDeck()
		const failures = []
		const seen = new Set()
		let checked = 0
		for (const name of Object.keys(zip.files)) {
			if (!name.endsWith('.xml')) continue
			const xml = await zip.file(name).async('string')
			for (const element of coveredElements(xml)) {
				checked++
				seen.add(element.name)
				const problem = outOfOrder(element.children, element.sequence)
				if (problem) failures.push(`${name} <${element.name}>: ${problem} (children: ${element.children.join(', ')})`)
			}
		}
		// A silent zero would make this test pass on a deck that emitted none of these, and a
		// count alone would not notice one whole element kind dropping out of the deck.
		assert(checked > 200, `expected the deck to exercise the covered types; only ${checked} element(s) matched`)
		const missing = REACHED.filter((name) => !seen.has(name))
		assert(missing.length === 0, `the deck no longer emits: ${missing.join(', ')}`)
		assert(failures.length === 0, `children out of schema order:\n  ${failures.join('\n  ')}`)
	})
})
