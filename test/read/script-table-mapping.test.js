// Unit tests on the table mapper (`src/script/from-read/table.ts`) for values both IRs of a round
// trip would carry wrong in the same way. The corpus round trip cannot see those: the source deck
// and the rebuilt deck convert to the same wrong value, so the diff is clean while the rebuilt
// deck differs from the source. Each case rewrites a built deck's XML into the shape under test,
// converts it, and asserts the IR value itself.

import { describe, test } from 'vitest'
import JSZip from 'jszip'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import { assert, assertEqual, captureDiagnostics } from '../helpers.js'
import { authorRead } from './authored.js'

/** Apply `rewrite` to every slide part of `buf`, reload, and convert. */
async function irWithSlideXml(buf, rewrite) {
	const zip = await JSZip.loadAsync(buf)
	for (const name of Object.keys(zip.files)) {
		if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue
		zip.file(name, rewrite(await zip.file(name).async('string')))
	}
	const reopened = await Presentation.load(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
	return readModelToIr(reopened)
}

/** The first `addTable` call on slide 1. */
function tableCall(ir) {
	const call = ir.slides[0].calls.find((c) => c.method === 'addTable')
	assert(call, 'the IR carries an addTable call')
	return call
}

describe('table mapper: cell margins', () => {
	test('a cell that states one margin side keeps the schema default on the other three', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'A' }]], { x: 0.5, y: 0.5, w: 4 })
		})
		let rewritten = 0
		const ir = await irWithSlideXml(buf, (xml) =>
			xml.replace(/marL="\d+" marR="\d+" marT="\d+" marB="\d+"/g, () => {
				rewritten++
				return 'marL="0"'
			})
		)
		assertEqual(rewritten, 1, 'the one cell now states only marL')
		const margin = tableCall(ir).args[0][0][0].options.margin
		// `[top, right, bottom, left]` in inches: 0.05 and 0.1 are 45720 and 91440 EMU.
		assertEqual(JSON.stringify(margin), JSON.stringify([0.05, 0.1, 0.05, 0]), 'the unset sides take the defaults')
	})
})

describe('table mapper: row heights', () => {
	// An auto-height row (`a:tr/@h` of 0) among fixed rows used to become a `rowH` entry of 0. The
	// writer rejects 0 as not a height and warns, so the IR said something the replay refused.
	test('an auto-height row among fixed rows is null, and the replay takes it without a warning', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'one' }], [{ text: 'two' }], [{ text: 'three' }]], {
				x: 0.5,
				y: 0.5,
				w: 4,
				h: 1.5,
				rowH: [0.5, 0.5, 0.5],
			})
		})
		let seen = 0
		const ir = await irWithSlideXml(buf, (xml) =>
			xml.replace(/<a:tr h="457200">/g, (match) => (++seen === 2 ? '<a:tr h="0">' : match))
		)
		assertEqual(seen, 3, 'the table has three rows at 0.5in')

		const call = tableCall(ir)
		assertEqual(JSON.stringify(call.args[1].rowH), JSON.stringify([0.5, null, 0.5]), 'the auto row is null')
		const note = ir.fidelity.find((entry) => entry.construct === 'table.rowAuto')
		assert(note, 'the auto row is noted')
		assert(
			JSON.stringify(note).includes('null'),
			`the note says the row is written as null; got ${JSON.stringify(note)}`
		)

		const { result: bytes, codes } = await captureDiagnostics(() => {
			const pres = new TsPptx()
			pres.addSlide().addTable(call.args[0], call.args[1])
			return pres.toBytes()
		})
		assert(!codes.includes('table/invalid-row-height'), `the replay accepts the row; got ${JSON.stringify(codes)}`)
		const slide = await (await JSZip.loadAsync(bytes)).file('ppt/slides/slide1.xml').async('string')
		const heights = [...slide.matchAll(/<a:tr h="(\d+)">/g)].map((m) => Number(m[1]))
		// The unpinned row takes an even share of the 1.5in table: 0.5in, the same as its neighbours.
		assertEqual(JSON.stringify(heights), JSON.stringify([457200, 457200, 457200]), 'every row is 0.5in')
	})
})
