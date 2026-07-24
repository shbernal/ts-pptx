// The Node runtime adapter's http/fetch branches (src/runtime/node.ts): loading a
// font or an image from an `http(s)://` source. The sibling node-runtime.test.mjs
// covers only the filesystem paths and error, leaving the fetch branches — the
// weakest spot in the tree (25% branch) — untouched because they need the network.
//
// `globalThis.fetch` is stubbed with a controlled response so both the success return
// and the `!response.ok` throw run deterministically, offline, on every platform.
// A real committed font (Silkscreen, OFL) feeds the font-success path so
// opentype.js actually parses the "downloaded" bytes.
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, test, expect, afterEach } from 'vitest'
import JSZip from 'jszip'
import TsPptx from '../../dist/node.js'

const FONT_PATH = fileURLToPath(new URL('../read/fixtures/fonts/Silkscreen-Regular.ttf', import.meta.url))
const FONT_BYTES = readFileSync(FONT_PATH)

// A 1x1 transparent PNG — loadMedia never decodes it (explicit w/h is given), it only
// base64-encodes the "downloaded" bytes, so any well-formed PNG payload works.
const PNG_1x1 = Buffer.from(
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
	'base64'
)

const realFetch = globalThis.fetch
afterEach(() => {
	globalThis.fetch = realFetch
})

/** Stub globalThis.fetch with a fixed response. `body` is a Buffer/Uint8Array; `ok` toggles the error branch. */
function stubFetch({ ok = true, body = Buffer.alloc(0) } = {}) {
	const calls = []
	globalThis.fetch = async (url) => {
		calls.push(url)
		return {
			ok,
			async arrayBuffer() {
				return body.buffer.slice(body.byteOffset, body.byteOffset + body.byteLength)
			},
		}
	}
	return calls
}

describe('node runtime: font loading over http', () => {
	test('registerFontMetrics fetches and parses a font from an http(s) source', async () => {
		const calls = stubFetch({ ok: true, body: FONT_BYTES })
		const pptx = new TsPptx()
		// Resolves (no throw) — the fetched bytes are a real, parseable font.
		await pptx.registerFontMetrics('SilkHttp', 'https://example.com/Silkscreen-Regular.ttf')
		expect(calls).toEqual(['https://example.com/Silkscreen-Regular.ttf'])
	})

	test('a non-ok font response rejects and names the fetch failure', async () => {
		stubFetch({ ok: false })
		const pptx = new TsPptx()
		await expect(pptx.registerFontMetrics('MissingHttp', 'http://example.com/nope.ttf')).rejects.toThrow(
			/Unable to load font \(fetch\)/
		)
	})
})

describe('node runtime: image loading over http', () => {
	test('an http(s) image path is fetched and embedded as base64 media', async () => {
		const calls = stubFetch({ ok: true, body: PNG_1x1 })
		const pptx = new TsPptx()
		pptx.addSlide().addImage({ path: 'https://example.com/pixel.png', x: 1, y: 1, w: 1, h: 1 })
		const buf = /** @type {Uint8Array} */ (await pptx.stream())
		expect(calls).toEqual(['https://example.com/pixel.png'])

		// The fetched image lands in the package as a media part.
		const zip = await JSZip.loadAsync(buf)
		const media = Object.keys(zip.files).filter((f) => f.startsWith('ppt/media/'))
		expect(media.length).toBeGreaterThan(0)
	})

	test('a non-ok image response rejects and names the fetch failure', async () => {
		stubFetch({ ok: false })
		const pptx = new TsPptx()
		pptx.addSlide().addImage({ path: 'http://example.com/missing.png', x: 1, y: 1, w: 1, h: 1 })
		// The export wraps loadMedia's "Unable to load image (fetch)" as the cause of a
		// "Failed to load media …" error; reaching either proves the non-ok fetch branch ran.
		await expect(pptx.stream()).rejects.toThrow(/Failed to load media .*missing\.png/)
	})
})
