// Unit tests on the table mapper (`src/script/from-read/table.ts`) for values both IRs of a round
// trip would carry wrong in the same way. The corpus round trip cannot see those: the source deck
// and the rebuilt deck convert to the same wrong value, so the diff is clean while the rebuilt
// deck differs from the source. Each case rewrites a built deck's XML into the shape under test,
// converts it, and asserts the IR value itself.

import { describe, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'
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
