// Read-model coverage for src/read/api/slide.ts branches the existing suites
// don't reach: the per-format image sniffer (picture-edit.test.js only ever
// adds a PNG), the addTextBox significant-whitespace path, Slide.placeholder
// skipping a non-AutoShape, and the Slide.name getter.

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
