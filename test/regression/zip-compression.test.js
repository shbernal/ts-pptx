import PptxGenJS from '../../dist/node.js'
import { defineRegressionSuite, assert } from '../helpers.js'

// Exports previously defaulted to STORE (and the typed-output `write()` branch
// ignored `compression` entirely), producing packages several times larger than
// the same deck re-saved by PowerPoint, which always DEFLATEs (upstream #1268).
// DEFLATE is now the default; `compression: false` opts back out.

// Compression method of each ZIP local file header (offset 8 after PK\x03\x04):
// 0 = STORE, 8 = DEFLATE.
function localHeaderMethods(buf) {
	const bytes = new Uint8Array(buf)
	const methods = []
	for (let i = 0; i + 10 < bytes.length; i++) {
		if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x03 && bytes[i + 3] === 0x04) {
			methods.push(bytes[i + 8] | (bytes[i + 9] << 8))
		}
	}
	if (methods.length === 0) throw new Error('no ZIP local file headers found')
	return methods
}

// Map of each ZIP entry name -> compression method, parsed from local file
// headers (filename length at offset 26, extra length at 28, name at 30).
function localHeaderEntries(buf) {
	const bytes = new Uint8Array(buf)
	const dec = new TextDecoder()
	const entries = []
	for (let i = 0; i + 30 < bytes.length; i++) {
		if (bytes[i] === 0x50 && bytes[i + 1] === 0x4b && bytes[i + 2] === 0x03 && bytes[i + 3] === 0x04) {
			const method = bytes[i + 8] | (bytes[i + 9] << 8)
			const nameLen = bytes[i + 26] | (bytes[i + 27] << 8)
			const name = dec.decode(bytes.subarray(i + 30, i + 30 + nameLen))
			entries.push({ name, method })
		}
	}
	if (entries.length === 0) throw new Error('no ZIP local file headers found')
	return entries
}

// 1x1 PNG (already entropy-coded; DEFLATE buys nothing).
const PNG_1x1 =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

async function buildPresWithImage() {
	const pres = new PptxGenJS()
	pres.addSlide().addImage({ data: PNG_1x1, x: 1, y: 1, w: 1, h: 1 })
	return pres
}

async function buildPres() {
	const pres = new PptxGenJS()
	pres.addSlide().addText('compression default check', { x: 1, y: 1, w: 6, h: 1 })
	return pres
}

defineRegressionSuite('ZIP package compression default (#1268)', [
	{
		name: 'stream() defaults to DEFLATE entries',
		fn: async () => {
			const pres = await buildPres()
			const buf = await pres.stream()
			const methods = localHeaderMethods(buf)
			assert(
				methods.some((m) => m === 8),
				'expected DEFLATE (8) entries by default; got methods: ' + methods.join(',')
			)
			assert(
				!methods.includes(0) || methods.filter((m) => m === 0).length < methods.length,
				'expected non-STORE entries by default; got methods: ' + methods.join(',')
			)
		},
	},
	{
		name: 'compression:false opts out (all entries STORE)',
		fn: async () => {
			const pres = await buildPres()
			const buf = await pres.stream({ compression: false })
			const methods = localHeaderMethods(buf)
			assert(
				methods.every((m) => m === 0),
				'expected all STORE (0) entries with compression:false; got: ' + methods.join(',')
			)
		},
	},
	{
		name: 'write() with a typed output honors compression (previously ignored)',
		fn: async () => {
			const pres = await buildPres()
			const buf = await pres.write({ outputType: 'nodebuffer' })
			const methods = localHeaderMethods(buf)
			assert(
				methods.some((m) => m === 8),
				'expected DEFLATE entries from write({outputType}); got: ' + methods.join(',')
			)
		},
	},
	{
		// #1006: DEFLATE-ing already-compressed media wastes CPU on large decks.
		// Media parts are STORE-d per-entry while XML stays DEFLATE.
		name: 'already-compressed media is STORE while XML stays DEFLATE',
		fn: async () => {
			const pres = await buildPresWithImage()
			const buf = await pres.stream()
			const entries = localHeaderEntries(buf)
			const media = entries.filter((e) => e.name.startsWith('ppt/media/'))
			assert(media.length > 0, 'expected at least one ppt/media/* entry')
			assert(
				media.every((e) => e.method === 0),
				'expected media entries STORE (0); got: ' + media.map((e) => `${e.name}=${e.method}`).join(', ')
			)
			const slideXml = entries.filter((e) => e.name.startsWith('ppt/slides/slide') && e.name.endsWith('.xml'))
			assert(
				slideXml.some((e) => e.method === 8),
				'expected slide XML DEFLATE (8); got: ' + slideXml.map((e) => `${e.name}=${e.method}`).join(', ')
			)
		},
	},
	{
		name: 'compression:false still stores media (no per-entry override regression)',
		fn: async () => {
			const pres = await buildPresWithImage()
			const buf = await pres.stream({ compression: false })
			const entries = localHeaderEntries(buf)
			assert(
				entries.every((e) => e.method === 0),
				'expected all STORE with compression:false; got: ' + entries.map((e) => `${e.name}=${e.method}`).join(', ')
			)
		},
	},
	{
		name: 'default export is smaller than the STORE export of the same deck',
		fn: async () => {
			const presA = await buildPres()
			const presB = await buildPres()
			const deflated = await presA.stream()
			const stored = await presB.stream({ compression: false })
			assert(
				deflated.length < stored.length,
				`expected DEFLATE output (${deflated.length}B) smaller than STORE output (${stored.length}B)`
			)
		},
	},
])
