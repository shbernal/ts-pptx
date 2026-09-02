// One `p:sp`, two views, one answer.
//
// A master's or layout's placeholder shape is reachable two ways: through `placeholders`,
// which is the identity-and-geometry view, and through `shapes`, which is the full `AnyShape`
// paint surface. They read the same element, so every member they share must agree.
//
// They did not. `Placeholder` built its text frame with **no** inheritance context, so the
// `Run.resolved*` family — size, face, colour, bold — resolved through `SlideMaster.shapes`
// and came back `null` through `SlideMaster.placeholders`. It also re-derived identity from a
// hard-coded `p:nvSpPr` and geometry through a helper no other caller used, so two more
// answers were being computed twice with nothing keeping them in step.

import { describe, test } from 'vitest'
import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

/** The `resolved*` family plus identity and geometry, from whichever view is passed. */
function readable(view) {
	const run = view.textFrame?.paragraphs?.[0]?.runs?.[0] ?? null
	return {
		id: view.id,
		name: view.name,
		left: view.left,
		top: view.top,
		width: view.width,
		height: view.height,
		text: view.textFrame?.text ?? null,
		resolvedSizePt: run?.resolvedSizePt ?? null,
		resolvedFontFace: run?.resolvedFontFace ?? null,
		resolvedColor: run?.resolvedColor?.effectiveHex ?? null,
		resolvedBold: run?.resolvedBold ?? null,
	}
}

describe('a placeholder and the same shape read as an AutoShape', () => {
	test('agree on identity, geometry and every resolved run property', async () => {
		const presentation = await openFixture('autofit-cjk-wrap')
		const master = presentation.slides[0].master
		assert(master, 'the fixture has a master')
		assert(master.placeholders.length > 0, 'the master has placeholders')

		for (const placeholder of master.placeholders) {
			const shape = master.shapes.find((s) => s.id === placeholder.id)
			assert(shape, `the placeholder ${placeholder.type} is also in shapes`)
			assertEqual(
				JSON.stringify(readable(placeholder)),
				JSON.stringify(readable(shape)),
				`the two views of ${placeholder.type} disagree`
			)
		}
	})

	test('and a master placeholder run actually resolves its inherited size, face and colour', async () => {
		// The agreement above would also hold if both views reported `null` for everything, so
		// this is the half that makes it mean something.
		const presentation = await openFixture('autofit-cjk-wrap')
		const master = presentation.slides[0].master
		const resolved = master.placeholders
			.map((ph) => readable(ph))
			.filter((r) => r.resolvedSizePt !== null || r.resolvedFontFace !== null)
		assert(
			resolved.length > 0,
			'at least one master placeholder resolves an inherited run property; got ' +
				JSON.stringify(master.placeholders.map((ph) => readable(ph)))
		)
		resolved.forEach((r) => {
			assert(typeof r.resolvedSizePt === 'number', `resolvedSizePt is a size; got ${r.resolvedSizePt}`)
			assert(typeof r.resolvedFontFace === 'string', `resolvedFontFace is a face; got ${r.resolvedFontFace}`)
		})
	})
})
