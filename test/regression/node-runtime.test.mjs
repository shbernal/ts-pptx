// The Node runtime adapter (src/runtime/node.ts) reached through the public API:
// writeFile() to disk and the font-read error path. The http/fetch branches are not
// exercised (they need the network); this covers the filesystem paths deterministically.
import { readFileSync, rmSync, existsSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, test, expect, afterEach } from 'vitest'
import TsPptx from '../../dist/node.js'

const written = []
afterEach(() => {
	for (const f of written.splice(0)) if (existsSync(f)) rmSync(f, { force: true })
})

function tmpName(name) {
	// Math.random is unavailable in some harnesses; a monotonically unique-enough name
	// from the current test count is plenty for an isolated temp file.
	const p = join(tmpdir(), `ts-pptx-node-runtime-${written.length}-${name}`)
	written.push(p)
	return p
}

describe('node runtime: writeFile', () => {
	test('writes a non-empty .pptx to disk and returns the path', async () => {
		const pptx = new TsPptx()
		pptx.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1 })
		const target = tmpName('out.pptx')
		const returned = await pptx.writeFile({ fileName: target })
		expect(returned).toBe(target)
		expect(existsSync(target)).toBe(true)
		const bytes = readFileSync(target)
		expect(bytes.length).toBeGreaterThan(0)
		// PPTX is a ZIP → starts with the local-file-header magic "PK\x03\x04".
		expect(bytes[0]).toBe(0x50)
		expect(bytes[1]).toBe(0x4b)
	})

	test('appends .pptx when the fileName lacks the extension', async () => {
		const pptx = new TsPptx()
		pptx.addSlide().addText('hi', { x: 1, y: 1, w: 2, h: 1 })
		const base = tmpName('noext')
		written.push(`${base}.pptx`) // ensure cleanup of the extension-appended file
		const returned = await pptx.writeFile({ fileName: base })
		expect(returned).toBe(`${base}.pptx`)
		expect(existsSync(`${base}.pptx`)).toBe(true)
	})
})

describe('node runtime: font loading errors', () => {
	test('registerFontMetrics rejects and names an unreadable font path', async () => {
		const pptx = new TsPptx()
		const bad = join(tmpdir(), 'definitely-missing-font.ttf')
		await expect(pptx.registerFontMetrics('Missing', bad)).rejects.toThrow(/Unable to read font file/)
	})
})
