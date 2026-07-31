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
			pres.addSlide().addTable([[{ text: 'sideways', options: { textDirection: 'vert270' } }, { text: 'upright' }]], {
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

	test('a diagonal replays as a diagonal, not as an edge', async () => {
		const ir = await irFor((pres) => {
			pres.addSlide().addTable(
				[
					[
						{
							text: 'X',
							options: {
								border: { type: 'solid', color: '333333', width: 1 },
								diagonal: {
									tlToBr: { type: 'solid', color: 'C00000', width: 2 },
									blToTr: { type: 'solid', color: '0000C0', width: 1, dashType: 'lgDashDot' },
								},
							},
						},
					],
				],
				{ x: 1, y: 1, w: 4 }
			)
		})

		const call = tableCall(ir)
		const diagonal = call.args[0][0][0].options.diagonal
		assert(diagonal, 'the mapper carries a diagonal object separate from the edge tuple')
		assertEqual(diagonal.tlToBr.color, 'C00000', '╲ colour')
		assertEqual(diagonal.tlToBr.width, 2, '╲ width')
		assertEqual(diagonal.blToTr.dashType, 'lgDashDot', '╱ keeps its exact dash')
		assert(!constructs(ir).has('table.cell.borders.diagonal'), 'the old drop note is gone')

		const xml = await replay(call)
		assertEqual((xml.match(/<a:lnTlToBr/g) || []).length, 1, 'the ╲ diagonal replays')
		assertEqual((xml.match(/<a:lnBlToTr/g) || []).length, 1, 'and so does the ╱')
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

describe('table replication — the table background', () => {
	test('a tableFill comes back as a tableFill, not flattened onto every cell', async () => {
		const ir = await irFor((pres) => {
			pres.addSlide().addTable(
				[
					['A', 'B'],
					['C', 'D'],
				],
				{ x: 1, y: 1, w: 8, tableFill: { color: 'F2F2F2' } }
			)
		})

		const call = tableCall(ir)
		assertEqual(call.args[1].tableFill.color, 'F2F2F2', 'the background maps to tableFill')
		assert(call.args[1].fill === undefined, 'and not to `fill`, which would stamp it onto every cell')
		for (const cell of call.args[0].flat()) {
			assert(cell.options?.fill === undefined, 'no cell picks up the background as its own fill')
		}

		const xml = await replay(call)
		const tblPr = xml.match(/<a:tblPr(?:\/>|[^>]*>[\s\S]*?<\/a:tblPr>)/)[0]
		assert(tblPr.includes('val="F2F2F2"'), 'the replayed background is on a:tblPr again; got: ' + tblPr)
	})

	test('a gradient background round-trips with its stops and angle', async () => {
		const ir = await irFor((pres) => {
			pres.addSlide().addTable([['A']], {
				x: 1,
				y: 1,
				w: 4,
				tableFill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 90,
						stops: [
							{ position: 0, color: 'FFFFFF' },
							{ position: 100, color: '1A2B3C' },
						],
					},
				},
			})
		})

		const call = tableCall(ir)
		const gradient = call.args[1].tableFill.gradient
		assertEqual(gradient.kind, 'linear', 'kind')
		assertEqual(gradient.angle, 90, 'the angle needs no conversion — both sides use OOXML degrees')
		assertEqual(gradient.stops.length, 2, 'both stops')
		assertEqual(gradient.stops[1].color, '1A2B3C', 'the end stop keeps its colour')

		const xml = await replay(call)
		assert(xml.includes('<a:lin ang="5400000"'), 'the replayed background keeps its angle')
	})
})

describe('table replication — non-solid cell fills', () => {
	test('a gradient and a pattern cell both come back as themselves', async () => {
		// Before this, `resolvedFill` reported null for any non-solid choice and the mapper had
		// no other accessor to fall back on, so a gradient cell replicated as an UNFILLED one —
		// or, with a table style in play, as whatever banding colour the style resolved to.
		const ir = await irFor((pres) => {
			pres.addSlide().addTable(
				[
					[
						{
							text: 'grad',
							options: {
								fill: {
									type: 'gradient',
									gradient: {
										kind: 'linear',
										angle: 0,
										stops: [
											{ position: 0, color: 'FFFFFF' },
											{ position: 100, color: '1A2B3C' },
										],
									},
								},
							},
						},
						{
							text: 'hatch',
							options: {
								fill: { type: 'pattern', pattern: { preset: 'diagCross', fgColor: '1A2B3C', bgColor: 'FFFFFF' } },
							},
						},
					],
				],
				{ x: 1, y: 1, w: 8 }
			)
		})

		const call = tableCall(ir)
		const [grad, hatch] = call.args[0][0]
		assertEqual(grad.options.fill.type, 'gradient', 'the gradient cell keeps its type')
		assertEqual(grad.options.fill.gradient.stops.length, 2, 'and its stops')
		assertEqual(hatch.options.fill.type, 'pattern', 'the pattern cell keeps its type')
		assertEqual(hatch.options.fill.pattern.preset, 'diagCross', 'and its preset')
		assertEqual(hatch.options.fill.pattern.fgColor, '1A2B3C', 'and its foreground')

		const xml = await replay(call)
		assert(xml.includes('<a:gradFill'), 'the gradient cell replays as a gradient')
		assert(xml.includes('<a:pattFill prst="diagCross">'), 'and the pattern cell as a pattern')
	})
})

describe('table replication — anchorCtr and cell3D', () => {
	test('both survive the round trip and land back in a:tcPr', async () => {
		const ir = await irFor((pres) => {
			pres.addSlide().addTable(
				[
					[
						{
							text: 'A',
							options: {
								anchorCtr: true,
								cell3D: {
									preset: 'artDeco',
									width: 7,
									height: 7,
									material: 'metal',
									lightRig: { rig: 'threePt', dir: 't' },
								},
							},
						},
						{ text: 'B' },
					],
				],
				{ x: 1, y: 1, w: 8 }
			)
		})

		const call = tableCall(ir)
		const [first, second] = call.args[0][0]
		assertEqual(first.options.anchorCtr, true, 'anchorCtr carries')
		assertEqual(first.options.cell3D.preset, 'artDeco', 'the bevel preset carries')
		assertEqual(first.options.cell3D.width, 7, 'and its size, in points on both sides')
		assertEqual(first.options.cell3D.lightRig.rig, 'threePt', 'and the light rig')
		// `false` is the schema default, so an unset cell must carry nothing rather than an
		// explicit false — otherwise every plain cell in a replica grows an attribute.
		assert(second.options?.anchorCtr === undefined, 'an unset cell carries no anchorCtr')
		assert(second.options?.cell3D === undefined, 'and no cell3D')

		const xml = await replay(call)
		const tags = [...xml.matchAll(/<a:tcPr[^>]*>/g)].map((m) => m[0])
		assert(tags[0].includes('anchorCtr="1"'), 'the replayed cell is anchor-centred; got: ' + tags[0])
		assert(!tags[1].includes('anchorCtr'), 'the sibling is not; got: ' + tags[1])
		assert(xml.includes('<a:bevel w="88900" h="88900" prst="artDeco"/>'), 'the bevel replays in EMU')
		assert(xml.includes('<a:lightRig rig="threePt" dir="t"/>'), 'and the light rig')
	})
})
