// Read-model coverage for the footer-trio placeholder IDENTITY bug: a slide
// `dt`/`ftr`/`sldNum` placeholder with no own `a:xfrm` must resolve to the
// SAME-TYPE source placeholder's geometry, never to a different member of the
// trio.
//
// PowerPoint gives the trio different `idx` on the layout (dt=10/ftr=11/
// sldNum=12) than on the master (dt=2/ftr=3/sldNum=4), so a slide placeholder's
// `idx` never matches a master placeholder's. The buggy resolver then fell back
// to the master text-style *category* (`phCategory`), which collapses the whole
// trio into `'other'` -- so all three resolved to whichever `'other'`
// placeholder came first in document order (the date box here), producing a
// wrong position/size on read AND a wrong baked `a:xfrm` on
// `importSlide(..., { theme: 'preserve' })`. The fix matches these singleton
// types by exact `type`.
//
// Fixture: placeholder-footer-trio.pptx (real desktop PowerPoint). One slide on
// a layout whose dt/ftr/sldNum placeholders carry NO own `a:xfrm`, so all three
// resolve through the layout to the master. The master's three placeholders sit
// at three deliberately distinct boxes; the EMU values below were read straight
// out of ppt/slideMasters/slideMaster1.xml (independent of the reader under
// test), not produced by running the getter.

import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'

import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

// Read from ppt/slideMasters/slideMaster1.xml inside placeholder-footer-trio.pptx.
// Three visibly distinct boxes -- a mismatch is unambiguous.
const MASTER = {
	dt: { left: 508000, top: 6095999, width: 2540000, height: 508000 },
	ftr: { left: 3810000, top: 6349999, width: 4572000, height: 381000 },
	sldNum: { left: 9906000, top: 5841999, width: 1778000, height: 635000 },
}

// The slide placeholder shape of a given ph `type`.
function phShape(slide, type) {
	return slide.shapes.find((s) => s.placeholder?.type === type)
}

describe('read: footer-trio placeholder geometry inheritance', () => {
	for (const type of ['dt', 'ftr', 'sldNum']) {
		test(`a slide ${type} placeholder resolves to its OWN-TYPE master box, not another member of the trio`, async () => {
			const slide = (await openFixture('placeholder-footer-trio')).slides[0]
			const shape = phShape(slide, type)
			assert(shape, `expected a slide ${type} placeholder`)

			// No own transform -- proving resolvedFrame is inherited.
			assertEqual(shape.left, null, `${type} has no own left`)
			assertEqual(shape.top, null, `${type} has no own top`)

			const frame = shape.resolvedFrame
			assert(frame, `expected a resolved frame for ${type}`)
			assertEqual(frame.source, 'master', `${type} resolves through the layout (no own xfrm) to the master`)
			assertEqual(frame.left, MASTER[type].left, `${type} resolved left`)
			assertEqual(frame.top, MASTER[type].top, `${type} resolved top`)
			assertEqual(frame.width, MASTER[type].width, `${type} resolved width`)
			assertEqual(frame.height, MASTER[type].height, `${type} resolved height`)
		})
	}

	test('the three boxes are distinct, so a trio mix-up would be caught (guards the fixture)', async () => {
		const boxes = [MASTER.dt, MASTER.ftr, MASTER.sldNum].map((b) => `${b.left},${b.top},${b.width},${b.height}`)
		assertEqual(new Set(boxes).size, 3, 'expected three distinct master boxes')
	})
})
