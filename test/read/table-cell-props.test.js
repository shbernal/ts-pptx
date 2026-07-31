// Write→read fidelity for the `a:tcPr` accessors added alongside their write options:
// `TableCell.anchorCtr`, `.cell3D`, and the two diagonals on `.borders`.
//
// Each is proven the way `table-borders.test.js` proves the edges: author the construct
// with the write API, load the bytes back through the deep read model, and assert the
// extracted values. The write path and the read path are separate code, so neither can
// mask a bug in the other.
//
// `.id` / `.headerIds` are the exception and are tested differently — see the last
// describe block for why they cannot have a PowerPoint-authored oracle.

import { describe, test } from 'vitest'
import JSZip from 'jszip'
import { Presentation } from '../../dist/read.js'
import { authorRead, firstTable, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

/** A one-cell table carrying every new `a:tcPr` construct at once. */
function decoratedTable(pres) {
	pres.addSlide().addTable(
		[
			[
				{
					text: 'A',
					options: {
						anchorCtr: true,
						valign: 'middle',
						border: { type: 'solid', color: '333333', width: 1 },
						diagonal: {
							tlToBr: { type: 'solid', color: 'C00000', width: 2 },
							blToTr: { type: 'solid', color: '0000C0', width: 1, dashType: 'lgDashDot' },
						},
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
		{ x: 1, y: 1, w: 8, colW: [4, 4] }
	)
}

describe('TableCell.anchorCtr — a:tcPr/@anchorCtr', () => {
	test('an anchorCtr cell reads true and its sibling reads false', async () => {
		const { presentation } = await authorRead(decoratedTable)
		const table = firstTable(presentation)
		assertEqual(table.cell(0, 0).anchorCtr, true, 'the authored cell reports true')
		// `false` is the schema default and is never written, so absent must read as false —
		// not null. An accessor that reported null here would make "off" and "unset" two
		// different answers to a question that has one.
		assertEqual(table.cell(0, 1).anchorCtr, false, 'an unset cell reports false, not null')
	})
})

describe('TableCell.cell3D — a:tcPr/a:cell3D', () => {
	test('the bevel, material and light rig all read back', async () => {
		const { presentation } = await authorRead(decoratedTable)
		const cell3D = firstTable(presentation).cell(0, 0).cell3D
		assert(cell3D, 'the authored cell surfaces a cell3D')
		assertEqual(cell3D.material, 'metal', 'prstMaterial')
		assertEqual(cell3D.bevel.preset, 'artDeco', 'bevel preset')
		// The write option is points and the attribute is EMU; the accessor converts back.
		assertEqual(cell3D.bevel.widthPt, 7, 'bevel width in points')
		assertEqual(cell3D.bevel.heightPt, 7, 'bevel height in points')
		assertEqual(cell3D.lightRig.rig, 'threePt', 'light rig type')
		assertEqual(cell3D.lightRig.dir, 't', 'light rig direction')
	})

	test('a cell with no cell3D reports null', async () => {
		const { presentation } = await authorRead(decoratedTable)
		assertEqual(firstTable(presentation).cell(0, 1).cell3D, null, 'no a:cell3D → null')
	})

	test('an empty cell3D reads its required bevel with every field unset', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addTable([[{ text: 'A', options: { cell3D: {} } }]], { x: 1, y: 1, w: 4 })
		})
		const cell3D = firstTable(presentation).cell(0, 0).cell3D
		assert(cell3D, 'CT_Cell3D requires a bevel, so an empty option still produces one')
		assertEqual(cell3D.bevel.preset, null, 'an unwritten preset reads null, not its default')
		assertEqual(cell3D.bevel.widthPt, null, 'and so does an unwritten width')
		assertEqual(cell3D.lightRig, null, 'no light rig')
	})
})

describe('TableCell.borders — the two diagonals', () => {
	test('an authored diagonal reads its width, colour and dash', async () => {
		const { presentation } = await authorRead(decoratedTable)
		const borders = firstTable(presentation).cell(0, 0).borders
		assert(borders, 'the cell surfaces borders')

		assert(borders.tlToBr, 'the ╲ diagonal is read')
		assertEqual(borders.tlToBr.widthPt, 2, '╲ width in points')
		assertEqual(borders.tlToBr.color, 'C00000', '╲ colour')
		assertEqual(borders.tlToBr.dash, 'solid', '╲ dash')

		assert(borders.blToTr, 'the ╱ diagonal is read')
		assertEqual(borders.blToTr.widthPt, 1, '╱ width in points')
		assertEqual(borders.blToTr.dash, 'lgDashDot', '╱ keeps its exact dash preset')
	})

	test('a cell with no diagonals still reads its four edges', async () => {
		const { presentation } = await authorRead(decoratedTable)
		const borders = firstTable(presentation).cell(0, 1).borders
		assert(borders, 'the plain cell still carries the writer default border set')
		assertEqual(borders.tlToBr, null, 'no ╲ authored')
		assertEqual(borders.blToTr, null, 'no ╱ authored')
	})

	test.skipIf(!validatorInstalled)('the decorated deck is schema-valid', async () => {
		const { buf } = await authorRead(decoratedTable)
		assertEqual((await schemaErrors(buf)).length, 0, 'decorated deck validates')
	})
})

describe('TableCell.id / .headerIds — a:tc/@id and a:tcPr/a:headers', () => {
	// These two have NO write-API counterpart and no PowerPoint-authored oracle, and both
	// facts have the same cause: PowerPoint opens a deck carrying them without complaint and
	// then strips both on the first save. `probe-table-cell-a11y-and-3d.ps1` measures that —
	// `a:cell3D` and `a:headers` were injected into the same `a:tcPr` and PowerPoint kept one
	// and discarded the other, so it is a deliberate normalization rather than a bad patch.
	//
	// So the oracle here is ECMA-376 §21.1.3.4's own worked example rather than Office
	// output, and the fixture is built by patching an authored deck exactly as the probe
	// does. The accessors exist because a deck from another producer may carry what
	// PowerPoint will not write; reading it back is the only half of this that can work.

	/** Author a 3x3 table, then patch in the spec's header association, and reload. */
	async function withHeaderAssociation() {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addTable(
				// The 3x3 table of ECMA-376 §21.1.3.4: headers A/B across, C/D down.
				[
					[{ text: '' }, { text: 'A' }, { text: 'B' }],
					[{ text: 'C' }, { text: 'x1' }, { text: 'x2' }],
					[{ text: 'D' }, { text: 'y1' }, { text: 'y2' }],
				],
				{ x: 1, y: 1, w: 9 }
			)
		})
		const zip = await JSZip.loadAsync(buf)
		let xml = await zip.file('ppt/slides/slide1.xml').async('string')

		// The nine `<a:tc …>` in document order; header cells are 1, 2, 3 and 6.
		const ids = { 1: 'HeaderA', 2: 'HeaderB', 3: 'HeaderC', 6: 'HeaderD' }
		let tcIndex = -1
		xml = xml.replace(/<a:tc[ >]/g, (match) => {
			tcIndex++
			return ids[tcIndex] ? `<a:tc id="${ids[tcIndex]}"${match.endsWith('>') ? '>' : ' '}` : match
		})
		// `a:headers` is the LAST child of the sequence, so it appends just before the close.
		const headers = { 4: ['HeaderA', 'HeaderC'], 5: ['HeaderB', 'HeaderC'] }
		let prIndex = -1
		xml = xml.replace(/<\/a:tcPr>/g, (match) => {
			prIndex++
			const list = headers[prIndex]
			if (!list) return match
			return `<a:headers>${list.map((h) => `<a:header val="${h}"/>`).join('')}</a:headers>${match}`
		})

		zip.file('ppt/slides/slide1.xml', xml)
		return Presentation.load(await zip.generateAsync({ type: 'uint8array' }))
	}

	test('a header cell reports its id and a data cell reports the headers governing it', async () => {
		const table = firstTable(await withHeaderAssociation())
		assertEqual(table.cell(0, 1).id, 'HeaderA', 'the column header carries its id')
		assertEqual(table.cell(1, 0).id, 'HeaderC', 'the row header carries its id')
		assertEqual(table.cell(0, 0).id, null, 'a cell with no id reports null')

		assertEqual(table.cell(1, 1).headerIds.join(','), 'HeaderA,HeaderC', 'x1 names its column and row headers')
		assertEqual(table.cell(1, 2).headerIds.join(','), 'HeaderB,HeaderC', 'x2 names its own column header')
		assertEqual(table.cell(2, 1).headerIds.length, 0, 'a cell with no association reports an empty list')
	})
})
