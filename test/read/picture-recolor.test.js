// Write→read fidelity for the three picture recolour modes the writer gained in
// T1.1 (grayscale / biLevel / clrChange), closing the read/write parity gap: the
// read model (src/read/api/shapes.ts `Picture.recolor`) already decoded five
// recolour effects, but the writer authored only `a:alphaModFix` (transparency)
// and `a:duotone`. Each test authors an image carrying one new mode with the
// write API, reads it back through the deep model, and asserts the `Recolor`
// union round-trips — proving the very bytes the writer emits decode faithfully.
//
// (The read side of these modes is already covered off-fixture in
// style-accessors.test.js via hand-authored `a:clrChange`/`a:grayscl`/`a:biLevel`
// XML; this file proves the *writer* now produces those same bytes.)

import { describe, test } from 'vitest'
import { authorRead, firstShape, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

// A 1×1 PNG — the smallest valid raster the writer will embed; the recolour
// effect lives on the blip, not the pixels, so its content is irrelevant.
const PNG_1x1 =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP8z8DwHwAFAAH/Re1ZlAAAAABJRU5ErkJggg=='

/** Author a single image carrying `imgOpts`, and read back its `recolor`. */
async function recolorOf(imgOpts) {
	const { presentation, buf } = await authorRead((pres) => {
		pres.addSlide().addImage({ data: PNG_1x1, x: 1, y: 1, w: 2, h: 2, ...imgOpts })
	})
	const picture = firstShape(presentation, (s) => s.shapeType === 'picture')
	assert(picture, 'the authored image is read back as a picture')
	return { recolor: picture.recolor, buf }
}

describe('Picture recolour — write→read fidelity (T1.1)', () => {
	test('grayscale: true authors a:grayscl and reads back kind "grayscale"', async () => {
		const { recolor } = await recolorOf({ grayscale: true })
		assert(recolor, 'the authored grayscale recolour reads back')
		assertEqual(recolor.kind, 'grayscale', 'grayscale: true → a:grayscl → kind "grayscale"')
	})

	test('biLevel authors a:biLevel@thresh and round-trips the 0–1 threshold', async () => {
		const { recolor } = await recolorOf({ biLevel: { threshold: 0.5 } })
		assert(recolor, 'the authored biLevel recolour reads back')
		assertEqual(recolor.kind, 'biLevel', 'biLevel → a:biLevel → kind "biLevel"')
		assertEqual(recolor.threshold, 0.5, 'threshold 0.5 → thresh 50000 → 0.5')
	})

	test('biLevel threshold 0 is authored (not dropped) and reads back 0', async () => {
		const { recolor } = await recolorOf({ biLevel: { threshold: 0 } })
		assert(recolor, 'a zero threshold still authors a biLevel effect')
		assertEqual(recolor.kind, 'biLevel', 'kind "biLevel"')
		assertEqual(recolor.threshold, 0, 'threshold 0 → thresh 0 → 0')
	})

	test('clrChange authors a:clrChange and round-trips two explicit hex colours', async () => {
		const { recolor } = await recolorOf({ clrChange: { from: '000000', to: 'FF0000' } })
		assert(recolor, 'the authored clrChange recolour reads back')
		assertEqual(recolor.kind, 'clrChange', 'clrChange → a:clrChange → kind "clrChange"')
		assertEqual(recolor.from.color, '000000', 'clrFrom is the explicit source hex')
		assertEqual(recolor.from.schemeColor, null, 'a hex clrFrom carries no scheme token')
		assertEqual(recolor.to.color, 'FF0000', 'clrTo is the explicit replacement hex')
	})

	test('clrChange preserves a scheme-colour target as a theme token', async () => {
		const { recolor } = await recolorOf({ clrChange: { from: 'FF0000', to: 'accent1' } })
		assert(recolor, 'the authored clrChange reads back')
		assertEqual(recolor.kind, 'clrChange', 'kind "clrChange"')
		assertEqual(recolor.from.color, 'FF0000', 'the hex source round-trips')
		assertEqual(recolor.to.schemeColor, 'accent1', 'a scheme clrTo is left as a token for the resolver')
		assertEqual(recolor.to.color, null, 'a scheme clrTo carries no explicit hex')
	})

	test.skipIf(!validatorInstalled)('all three authored recolour decks are schema-valid', async () => {
		const { buf: grayBuf } = await recolorOf({ grayscale: true })
		assertEqual((await schemaErrors(grayBuf)).length, 0, 'grayscale deck validates')
		const { buf: biLevelBuf } = await recolorOf({ biLevel: { threshold: 0.5 } })
		assertEqual((await schemaErrors(biLevelBuf)).length, 0, 'biLevel deck validates')
		const { buf: clrChangeBuf } = await recolorOf({ clrChange: { from: '000000', to: 'FF0000' } })
		assertEqual((await schemaErrors(clrChangeBuf)).length, 0, 'clrChange deck validates')
	})
})
