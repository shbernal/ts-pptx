// `Shape.hyperlink` — the click link on a whole shape (`p:cNvPr/a:hlinkClick`).
//
// DrawingML puts the same element in two places: on a run's `a:rPr`, linking a span of text, and
// on a shape's `p:cNvPr`, linking the shape. Only the run's had a reader, so a linked shape read
// back as an ordinary one and anything re-authoring from the read model dropped the link with
// nothing to say so — although `addShape`, `addText` and `addImage` all take `hyperlink`.
//
// The ground truth for the slide-jump form is `slide-jump-link.pptx`, whose slide 2 carries a
// shape PowerPoint itself linked to another slide. The URL and action forms are authored here,
// because what they write is the same element with different attributes and the write path is the
// thing under test for those.

import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import TsPptx from '../../dist/node.js'
import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

/** Every shape on every slide, flattened. */
const allShapes = (presentation) => presentation.slides.flatMap((slide) => slide.shapes)

/** A deck with one linked shape of each form, read back. */
async function authorLinked() {
	const pres = new TsPptx()
	const slide = pres.addSlide()
	slide.addShape('rect', {
		x: 1,
		y: 1,
		w: 2,
		h: 1,
		objectName: 'urlShape',
		hyperlink: { url: 'https://example.invalid/a', tooltip: 'go' },
	})
	slide.addShape('rect', { x: 1, y: 3, w: 2, h: 1, objectName: 'jumpShape', hyperlink: { slide: 2 } })
	slide.addShape('rect', { x: 4, y: 1, w: 2, h: 1, objectName: 'actionShape', hyperlink: { action: 'nextslide' } })
	slide.addShape('rect', { x: 4, y: 3, w: 2, h: 1, objectName: 'plainShape' })
	pres.addSlide()
	const buf = await pres.toBytes()
	return { presentation: await Presentation.load(buf), buf }
}

describe('Shape.hyperlink', () => {
	test("reads a slide jump PowerPoint authored, and resolves it to the target's part", async () => {
		const presentation = await openFixture('slide-jump-link')
		const linked = allShapes(presentation).find((shape) => shape.hyperlink !== null)
		assert(linked, 'the fixture has a shape carrying a link')
		const link = linked.hyperlink
		assertEqual(link.action, 'ppaction://hlinksldjump', 'it is a slide jump')
		assert(link.relId, 'it is backed by a relationship')
		assert(
			/^\/ppt\/slides\/slide\d+\.xml$/.test(link.targetPartName ?? ''),
			`it resolves to a slide part; got ${link.targetPartName}`
		)
		assertEqual(link.url, null, 'and not to an external url')
	})

	test('reads the url, jump and action forms, and null for an unlinked shape', async () => {
		const { presentation } = await authorLinked()
		const byName = Object.fromEntries(allShapes(presentation).map((shape) => [shape.name, shape.hyperlink]))

		assertEqual(byName.urlShape.url, 'https://example.invalid/a', 'a url link reports its url')
		assertEqual(byName.urlShape.tooltip, 'go', 'and its tooltip')
		assertEqual(byName.urlShape.targetPartName, null, 'and no internal target')

		assertEqual(byName.jumpShape.action, 'ppaction://hlinksldjump', 'a slide jump reports its action')
		assertEqual(byName.jumpShape.targetPartName, '/ppt/slides/slide2.xml', 'and the slide it points at')

		// An action-only link carries no `@r:id`, so there is nothing to resolve either way.
		assertEqual(byName.actionShape.action, 'ppaction://hlinkshowjump?jump=nextslide', 'a show jump reports its action')
		assertEqual(byName.actionShape.relId, null, 'and is backed by no relationship')

		assertEqual(byName.plainShape, null, 'an unlinked shape reports null')
	})

	test('a run link and a shape link stay independent', async () => {
		// The two are different elements on different parents. Reading one through the other's
		// accessor is the mistake a single shared reader could invite.
		const pres = new TsPptx()
		pres.addSlide().addText([{ text: 'inner', options: { hyperlink: { url: 'https://example.invalid/run' } } }], {
			x: 1,
			y: 1,
			w: 4,
			h: 1,
			objectName: 'runLinked',
		})
		const presentation = await Presentation.load(await pres.toBytes())
		const shape = allShapes(presentation).find((candidate) => candidate.name === 'runLinked')
		assertEqual(shape.hyperlink, null, 'the shape itself carries no link')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.hyperlink.url, 'https://example.invalid/run', "but its run's link reads")
	})
})

describe('pptxToScript keeps a shape hyperlink', () => {
	test('prints url, slide and action, with no fidelity note', async () => {
		const { presentation } = await authorLinked()
		const ir = readModelToIr(presentation)
		const calls = ir.slides.flatMap((slide) => slide.calls)
		/** The call's options bag — always its last argument, whatever the method. */
		const optionsOf = (name) => {
			const call = calls.find((candidate) => candidate.sourceName === name)
			assert(call, `${name} emits a call`)
			return /** @type {any} */ (call.args[call.args.length - 1])
		}
		assertEqual(optionsOf('urlShape').hyperlink.url, 'https://example.invalid/a', 'a url link')
		assertEqual(optionsOf('urlShape').hyperlink.tooltip, 'go', 'with its tooltip')
		assertEqual(optionsOf('jumpShape').hyperlink.slide, 2, 'a slide jump prints the slide number')
		assertEqual(optionsOf('actionShape').hyperlink.action, 'nextslide', 'a show jump prints the action')
		assertEqual(optionsOf('plainShape').hyperlink, undefined, 'an unlinked shape prints no hyperlink')
		assertEqual(
			ir.fidelity.filter((note) => note.construct === 'shape.hyperlink').length,
			0,
			`all three forms are writable, so nothing is noted; got ${JSON.stringify(ir.fidelity.map((n) => n.construct))}`
		)
	})
})
