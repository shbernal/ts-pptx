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
