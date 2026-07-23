// Read-model coverage for placeholder EFFECTIVE GEOMETRY resolution
// (src/read/api/theme-context.ts `resolveInheritedFrame`, surfaced by
// `Shape.resolvedFrame` in src/read/api/shapes.ts). This is the F1 read-model
// follow-on from FIDELITY-BACKLOG.md: a slide placeholder's `left`/`top`/`width`/
// `height` getters read `null` when the shape carries no own `a:xfrm` (T2.2's
// scope stopped at that raw signal); `resolvedFrame` resolves the *effective*
// position/size by walking slide -> layout -> master, matching the placeholder's
// `type`/`idx`.
//
// The write API always inlines an explicit `a:xfrm` onto every placeholder it
// authors (src/gen/slide/object.ts), so there is no authored trigger for the
// inherited path -- only PowerPoint itself produces a placeholder with no own
// transform when the user never moves it. `placeholder-inherit.pptx` (already
// committed for the sibling text-inheritance suite, placeholder-inherit.test.js)
// is exactly that deck: one slide, a title + body placeholder, neither carrying
// its own `a:xfrm`, and -- confirmed by inspecting the fixture's own
// slideLayout12.xml -- the layout's placeholders don't either, so both
// placeholders resolve all the way through to slideMaster1.xml. The master
// geometry below was read directly out of that XML (independent of the reader
// code under test), not derived by running the getter.

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

// Read from ppt/slideMasters/slideMaster1.xml inside placeholder-inherit.pptx.
const MASTER_TITLE = { left: 838200, top: 365125, width: 10515600, height: 1325563 }
const MASTER_BODY = { left: 838200, top: 1825625, width: 10515600, height: 4351338 }

describe('read: placeholder effective geometry', () => {
	test('a title placeholder with no own a:xfrm resolves through layout to the master geometry', async () => {
		const slide = (await open('placeholder-inherit')).slides[0]
		const title = slide.shapes.find((s) => /^Title/.test(s.name ?? ''))
		assert(title, 'expected a title placeholder shape')

		// The shape itself carries no own transform -- proving resolvedFrame below is inherited.
		assertEqual(title.left, null, 'title has no own left')
		assertEqual(title.top, null, 'title has no own top')
		assertEqual(title.width, null, 'title has no own width')
		assertEqual(title.height, null, 'title has no own height')

		const frame = title.resolvedFrame
		assert(frame, 'expected a resolved frame')
		assertEqual(frame.source, 'master', 'title resolves through the layout (which also has no own xfrm) to the master')
		assertEqual(frame.left, MASTER_TITLE.left, 'title resolved left')
		assertEqual(frame.top, MASTER_TITLE.top, 'title resolved top')
		assertEqual(frame.width, MASTER_TITLE.width, 'title resolved width')
		assertEqual(frame.height, MASTER_TITLE.height, 'title resolved height')
	})

	test('a body placeholder with no own a:xfrm resolves through layout to the master geometry', async () => {
		const slide = (await open('placeholder-inherit')).slides[0]
		const body = slide.shapes.find((s) => !/^Title/.test(s.name ?? '') && s.textFrame)
		assert(body, 'expected a body placeholder shape')

		assertEqual(body.left, null, 'body has no own left')
		assertEqual(body.top, null, 'body has no own top')

		const frame = body.resolvedFrame
		assert(frame, 'expected a resolved frame')
		assertEqual(frame.source, 'master', 'body resolves through the layout (which also has no own xfrm) to the master')
		assertEqual(frame.left, MASTER_BODY.left, 'body resolved left')
		assertEqual(frame.top, MASTER_BODY.top, 'body resolved top')
		assertEqual(frame.width, MASTER_BODY.width, 'body resolved width')
		assertEqual(frame.height, MASTER_BODY.height, 'body resolved height')
	})

	test('a shape with its own a:xfrm reports source "own" and does not consult the chain (negative control)', async () => {
		const presentation = await open('rotation-flip')
		const shapes = presentation.slides[0].shapes
		const rotated = shapes.find((s) => s.name === 'rotated-45')
		assert(rotated, 'expected the rotated-45 rect')
		assert(rotated.left !== null, 'rotated-45 has its own xfrm')

		const frame = rotated.resolvedFrame
		assert(frame, 'expected a resolved frame')
		assertEqual(frame.source, 'own', 'a shape with its own xfrm reports source "own"')
		assertEqual(frame.left, rotated.left, "resolvedFrame.left matches the shape's own left")
		assertEqual(frame.top, rotated.top, "resolvedFrame.top matches the shape's own top")
		assertEqual(frame.width, rotated.width, "resolvedFrame.width matches the shape's own width")
		assertEqual(frame.height, rotated.height, "resolvedFrame.height matches the shape's own height")
	})
})
