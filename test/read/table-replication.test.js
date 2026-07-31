// Read → script → write fidelity for the table constructs `src/script/from-read/table.ts`
// used to drop on the floor.
//
// Each leg is a real round trip rather than an assertion about the mapper's source: author
// a deck with the write API, read it back through the deep model, convert to the deck IR,
// then feed the IR's own `addTable` arguments straight back into the write API. The second
// deck's XML is what proves the construct survived — a mapper that emitted a plausible-
// looking option the writer ignores would still fail here.

import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { readModelToIr } from '../../dist/script.js'
import JSZip from 'jszip'
import { authorRead } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** The IR's single `addTable` call, or a failing assertion. */
function tableCall(ir) {
	const call = ir.slides.flatMap((slide) => slide.calls).find((c) => c.method === 'addTable')
	assert(call, 'the IR carries an addTable call')
	return call
}

/** Note constructs recorded anywhere in the deck. */
function constructs(ir) {
	return new Set(ir.fidelity.map((note) => note.construct))
}

/** Replay an IR `addTable` call through the write API and return the slide part. */
async function replay(call) {
	const pres = new TsPptx()
	pres.addSlide().addTable(call.args[0], call.args[1])
	const buf = /** @type {Uint8Array} */ (await pres.stream())
	const zip = await JSZip.loadAsync(buf)
	return zip.file('ppt/slides/slide1.xml').async('string')
}

/** Author `build`, read it back, and convert to the deck IR. */
async function irFor(build) {
	const { buf } = await authorRead(build)
	return readModelToIr(await Presentation.load(buf))
}

describe('table replication — vertical cell text (a:tcPr/@vert)', () => {
	test('a vert270 cell comes back vertical instead of being flattened', async () => {
		const ir = await irFor((pres) => {
			pres
				.addSlide()
				.addTable([[{ text: 'sideways', options: { textDirection: 'vert270' } }, { text: 'upright' }]], {
					x: 1,
					y: 1,
					w: 8,
				})
		})

		const call = tableCall(ir)
		assertEqual(call.args[0][0][0].options.textDirection, 'vert270', 'the mapper carries the direction')
		assert(
			call.args[0][0][1].options?.textDirection === undefined,
			'a horizontal cell carries nothing — horz is the schema default'
		)
		assert(!constructs(ir).has('table.cell.vert'), 'a writable direction raises no note')

		const xml = await replay(call)
		const tcPrs = xml.match(/<a:tcPr[^>]*>/g)
		assertEqual(tcPrs.length, 2, 'two cells')
		assert(tcPrs[0].includes('vert="vert270"'), 'the replayed cell is vertical again; got: ' + tcPrs[0])
		assert(!tcPrs[1].includes('vert='), 'the sibling stays horizontal; got: ' + tcPrs[1])
	})
})

describe('table replication — cell border dash presets', () => {
	test('each dash preset survives instead of collapsing onto sysDash', async () => {
		// The write API is the only authoring surface here, so the dashes it can emit are the
		// dashes this leg can prove; `dashType` now spans the whole ST_PresetLineDashVal set,
		// which is the point.
		const dashes = ['lgDashDot', 'sysDot', 'dot', 'lgDash']
		const ir = await irFor((pres) => {
			pres.addSlide().addTable(
				[
					dashes.map((dash) => ({
						text: dash,
						options: { border: { type: 'solid', color: '336699', width: 1, dashType: dash } },
					})),
				],
				{ x: 1, y: 1, w: 9 }
			)
		})

		const call = tableCall(ir)
		for (const [idx, dash] of dashes.entries()) {
			const border = call.args[0][0][idx].options.border
			assert(Array.isArray(border), `cell ${idx} carries a four-side border tuple`)
			for (const side of border) {
				assertEqual(side.dashType, dash, `cell ${idx} keeps ${dash} on every side`)
			}
		}

		const xml = await replay(call)
		const values = [...xml.matchAll(/<a:prstDash val="([^"]*)"\/>/g)].map((m) => m[1])
		assertEqual(values.length, dashes.length * 4, 'four sides per cell')
		for (const [idx, dash] of dashes.entries()) {
			assertEqual(values.slice(idx * 4, idx * 4 + 4).join(','), Array(4).fill(dash).join(','), `${dash} replays`)
		}
	})

	test('a solid border still emits no dashType — solid is what its absence means', async () => {
		const ir = await irFor((pres) => {
			pres.addSlide().addTable([[{ text: 'A', options: { border: { type: 'solid', color: '000000', width: 1 } } }]], {
				x: 1,
				y: 1,
				w: 4,
			})
		})

		const border = tableCall(ir).args[0][0][0].options.border
		for (const side of border) {
			assert(side.dashType === undefined, 'a solid rule carries no dashType; got: ' + JSON.stringify(side))
			assertEqual(side.type, 'solid', 'it is still reported as solid')
		}
	})
})
