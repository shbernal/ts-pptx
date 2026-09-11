import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterAll, describe, test } from 'vitest'
import { PNG_1X1, assert, build } from '../../helpers.js'

// An inline source carries a placeholder path that the media pass must not try to load. That pass
// used to recognise the placeholder by the word in it, so a real file whose name merely contained
// the word was skipped too, and the image it named never reached the deck.

const dir = mkdtempSync(join(tmpdir(), 'ts-pptx-media-path-'))
afterAll(() => rmSync(dir, { recursive: true, force: true }))

describe('a media path that mentions the inline placeholder', () => {
	test('is loaded like any other path', async () => {
		const bytes = Buffer.from(PNG_1X1.split('base64,')[1] ?? '', 'base64')
		const path = join(dir, 'preencoded-logo.png')
		writeFileSync(path, bytes)

		const { zip } = await build((pres) => pres.addSlide().addImage({ path, x: 1, y: 1, w: 1, h: 1 }))

		const parts = zip.file(/^ppt\/media\/.+\.png$/)
		assert(parts.length === 1, `expected one image part, got ${parts.length}`)
		const written = await parts[0].async('nodebuffer')
		assert(Buffer.compare(written, bytes) === 0, 'the image part must carry the bytes of the file it names')
	})
})
