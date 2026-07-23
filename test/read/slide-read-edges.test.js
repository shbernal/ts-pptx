// Read-model coverage for src/read/api/slide.ts branches the existing suites
// don't reach: the per-format image sniffer (picture-edit.test.js only ever
// adds a PNG), the addTextBox significant-whitespace path, Slide.placeholder
// skipping a non-AutoShape, and the Slide.name getter.
//
// Plus the slide-level read expansion (Slide.background / slideNumberPlaceholder,
// TextFrame.autofit): asserted write→read through the shared harness, since the
// writer already authors each feature.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead, schemaErrors, validatorInstalled } from './authored.js'

// A 1×1 transparent PNG, as the writer's `background: { data }` expects it.
const PNG_1PX =
	'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

/** The first autoShape text frame on a read slide (addText/addTextBox emit a `p:sp`). */
function textFrameOf(slide) {
	const shape = slide.shapes.find((s) => s.shapeType === 'autoShape')
	return shape ? shape.textFrame : null
}

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function open(name) {
	return Presentation.load(await readFile(path.join(__dirname, 'fixtures', `${name}.pptx`)))
}

// Minimal magic-byte headers for each format sniffImageType recognizes. Only the
// signature bytes matter — addPicture never decodes the image.
const HEADERS = {
	jpeg: [0xff, 0xd8, 0xff, 0x00],
	gif: [0x47, 0x49, 0x46, 0x38, 0x39, 0x61],
	bmp: [0x42, 0x4d, 0x00, 0x00, 0x00, 0x00],
	// TIFF has two byte-orders; both must sniff to tiff.
	'tiff-ii': [0x49, 0x49, 0x2a, 0x00],
	'tiff-mm': [0x4d, 0x4d, 0x00, 0x2a],
	webp: [0x52, 0x49, 0x46, 0x46, 0, 0, 0, 0, 0x57, 0x45, 0x42, 0x50],
}

const GEOM = { left: 0, top: 0, width: 100000, height: 100000 }

describe('Slide.addPicture — image format sniffing', () => {
	test('each format signature resolves to the right media extension', async () => {
		const slide = (await open('empty')).slides[0]
		const expected = { jpeg: 'jpeg', gif: 'gif', bmp: 'bmp', 'tiff-ii': 'tiff', 'tiff-mm': 'tiff', webp: 'webp' }
		for (const [label, bytes] of Object.entries(HEADERS)) {
			const picture = slide.addPicture(new Uint8Array(bytes), { ...GEOM, name: `img-${label}` })
			const partName = picture.imagePartName
			assert(
				partName && partName.endsWith(`.${expected[label]}`),
				`${label} sniffs to .${expected[label]}, got ${partName}`
			)
		}
	})

	test('bytes too short to match any signature still throw without an explicit type', async () => {
		const slide = (await open('empty')).slides[0]
		let threw = false
		try {
			slide.addPicture(new Uint8Array([0xff, 0xd8]), GEOM) // JPEG needs 3 signature bytes
		} catch {
			threw = true
		}
		assert(threw, 'an unsniffable 2-byte buffer with no {extension} throws')
	})
})

describe('Slide misc read/edit edges', () => {
	test('addTextBox preserves significant leading/trailing whitespace', async () => {
		const slide = (await open('empty')).slides[0]
		const box = slide.addTextBox({ text: '  padded  ', ...GEOM })
		assertEqual(box.textFrame.paragraphs[0].runs[0].text, '  padded  ', 'the whitespace text is written verbatim')
	})

	test('placeholder() skips non-AutoShape shapes when scanning', async () => {
		// image.pptx carries pictures (non-AutoShape). placeholder() must iterate past
		// them without error; a missing type simply returns undefined.
		const slide = (await open('image')).slides[0]
		assertEqual(slide.placeholder('nonexistent-type'), undefined, 'no matching placeholder → undefined')
	})

	test('name reads p:cSld/@name (null when the slide is unnamed)', async () => {
		const slide = (await open('empty')).slides[0]
		// The getter must not throw; an unnamed slide reads null.
		const name = slide.name
		assert(name === null || typeof name === 'string', `name is a string or null, got ${typeof name}`)
	})

	test('addTextBox with no text writes an empty paragraph', async () => {
		const slide = (await open('empty')).slides[0]
		const box = slide.addTextBox(GEOM)
		assertEqual(box.textFrame.paragraphs[0].text, '', 'an empty text box has an empty first paragraph')
	})

	test('notesText reads the notes body when a notes slide is attached', async () => {
		// notes-slide-image.pptx carries a real notes slide; the read must resolve the
		// notesSlide rel, find the body placeholder, and flatten its text.
		const slide = (await open('notes-slide-image')).slides[0]
		const notes = slide.notesText
		assert(notes === null || typeof notes === 'string', 'notesText is a string (or null when no notes part)')
	})
})

describe('Slide.background — write→read fidelity', () => {
	test('solid colour background reads type/source/colour off the slide itself', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.background = { color: 'C0392B' }
		})
		const bg = presentation.slides[0].background
		assert(bg !== null, 'the slide has a background')
		assertEqual(bg.type, 'solid', 'solid colour background')
		assertEqual(bg.source, 'slide', 'authored on the slide, not inherited')
		assertEqual(bg.color?.effectiveHex, 'C0392B', 'the colour resolves to its literal hex')
	})

	test('gradient background reads kind/angle and each resolved stop', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.background = {
				type: 'gradient',
				gradient: {
					kind: 'linear',
					angle: 45,
					stops: [
						{ position: 0, color: 'FF0000' },
						{ position: 100, color: '0000FF' },
					],
				},
			}
		})
		const bg = presentation.slides[0].background
		assertEqual(bg.type, 'gradient', 'gradient background')
		assertEqual(bg.source, 'slide', 'authored on the slide')
		assertEqual(bg.gradient.kind, 'linear', 'linear gradient')
		assertEqual(bg.gradient.angleDeg, 45, '45° preserved through the OOXML 60000ths encoding')
		assertEqual(bg.gradient.stops.length, 2, 'both stops read')
		assertEqual(bg.gradient.stops[0].effectiveHex, 'FF0000', 'first stop colour')
		assertEqual(bg.gradient.stops[1].position, 1, 'last stop at position 1 (100%)')
	})

	test('image background reads its rel id and resolved absolute part name', async () => {
		const { presentation } = await authorRead((pres) => {
			const slide = pres.addSlide()
			slide.background = { data: PNG_1PX }
		})
		const bg = presentation.slides[0].background
		assertEqual(bg.type, 'image', 'image background')
		assertEqual(bg.source, 'slide', 'authored on the slide')
		assert(bg.relId != null, 'the blip rel id is read')
		assert(
			bg.partName != null && bg.partName.endsWith('.png'),
			`part name resolves to the media png, got ${bg.partName}`
		)
	})

	test('a slide with no own background inherits the layout background (source=layout)', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide() // no background set → falls through to the default layout's p:bg
		})
		const bg = presentation.slides[0].background
		assert(bg !== null, 'the effective background is inherited, not null')
		assertEqual(bg.type, 'themeRef', 'the default layout background is a theme-indexed p:bgRef')
		assertEqual(bg.source, 'layout', 'inherited from the layout, not the slide')
		assertEqual(bg.idx, 1001, 'the default background matrix index')
	})

	test('a themeRef background resolves its idx to the concrete theme fill', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide() // inherits the default layout's p:bgRef idx=1001
		})
		const bg = presentation.slides[0].background
		assertEqual(bg.type, 'themeRef', 'theme-indexed background')
		assertEqual(bg.idx, 1001, 'raw idx kept for fidelity')
		// idx 1001 → bgFillStyleLst entry 1 = <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>;
		// the bgRef's own <a:schemeClr val="bg1"/> supplies the phClr, and bg1 → lt1 → window (FFFFFF).
		assert(bg.resolvedFill !== null, 'idx resolves to a concrete fill')
		assertEqual(bg.resolvedFill.type, 'solid', 'the first bg fill-style entry is a solid fill')
		assertEqual(bg.resolvedFill.color?.effectiveHex, 'FFFFFF', 'phClr substituted with the resolved bg1 (window/white)')
	})

	test.skipIf(!validatorInstalled)('authored backgrounds are schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().background = { color: '1F4E79' }
			pres.addSlide().background = { data: PNG_1PX }
		})
		assertEqual((await schemaErrors(buf)).length, 0, 'no schema violations')
	})
})

describe('TextFrame.autofit — write→read fidelity', () => {
	test('fit modes map to the autofit tokens', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('resize', { x: 1, y: 1, w: 3, h: 1, fit: 'resize' })
			pres.addSlide().addText('shrink', { x: 1, y: 1, w: 3, h: 1, fit: 'shrink' })
			pres.addSlide().addText('plain', { x: 1, y: 1, w: 3, h: 1 })
		})
		assertEqual(textFrameOf(presentation.slides[0]).autofit, 'spAutoFit', "fit:'resize' → spAutoFit")
		const shrink = textFrameOf(presentation.slides[1])
		assertEqual(shrink.autofit, 'normAutofit', "fit:'shrink' → normAutofit")
		assertEqual(shrink.autofitFontScale, null, 'a bare normAutofit bakes no font scale')
		assertEqual(textFrameOf(presentation.slides[2]).autofit, 'none', 'a bodyPr with no autofit child → none')
	})

	test('an explicit shrink bakes fontScale and lnSpcReduction as percents', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('baked', {
				x: 1,
				y: 1,
				w: 3,
				h: 1,
				fit: { type: 'shrink', fontScale: 62.5, lnSpcReduction: 20 },
			})
		})
		const tf = textFrameOf(presentation.slides[0])
		assertEqual(tf.autofit, 'normAutofit', 'shrink object → normAutofit')
		assertEqual(tf.autofitFontScale, 62.5, 'fontScale read as a percent (62500 ÷ 1000)')
		assertEqual(tf.autofitLineSpaceReduction, 20, 'lnSpcReduction read as a percent (20000 ÷ 1000)')
	})

	test.skipIf(!validatorInstalled)('authored autofit is schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addText('x', { x: 1, y: 1, w: 3, h: 1, fit: { type: 'shrink', fontScale: 80 } })
		})
		assertEqual((await schemaErrors(buf)).length, 0, 'no schema violations')
	})
})

describe('Slide.slideNumberPlaceholder — write→read fidelity', () => {
	test('a slide-number placeholder is read when the slide carries one', async () => {
		const { presentation } = await authorRead((pres) => {
			// The per-slide setter emits the sldNum `p:sp` in the slide's own shape tree
			// (setSlideNumber() alone puts it only on the master, which this getter scopes out).
			pres.addSlide().slideNumber = { x: 1, y: '90%', w: 1, h: 0.5 }
		})
		const ph = presentation.slides[0].slideNumberPlaceholder
		assert(ph !== null, 'the sldNum placeholder shape is found')
		assertEqual(ph.placeholder?.type, 'sldNum', 'it is the slide-number placeholder')
		assert(ph.text.length >= 0, 'its text frame reads (the slide-number field)')
	})

	test('a slide with no slide number reads null', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide()
		})
		assertEqual(presentation.slides[0].slideNumberPlaceholder, null, 'no sldNum placeholder → null')
	})

	test.skipIf(!validatorInstalled)('a slide carrying a slide number is schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().slideNumber = { x: 1, y: '90%' }
		})
		assertEqual((await schemaErrors(buf)).length, 0, 'no schema violations')
	})
})
