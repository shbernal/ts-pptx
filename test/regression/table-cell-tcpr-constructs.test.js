import { defineRegressionSuite, build, readEntry, captureDiagnostics, assert, assertEqual } from '../helpers.js'

// The three `a:tcPr` constructs that had no write surface: the two diagonals
// (`a:lnTlToBr` / `a:lnBlToTr`), `@anchorCtr`, and `a:cell3D`.
//
// The invariant every case here defends is CHILD ORDER. `CT_TableCellProperties` declares
// its children as a SEQUENCE -- lnL, lnR, lnT, lnB, lnTlToBr, lnBlToTr, cell3D,
// <EG_FillProperties>, headers, extLst -- and PowerPoint reports an out-of-order slide part
// as a corrupt file rather than as a bad option, so the order is asserted directly rather
// than inferred from the deck opening.
//
// `a:headers` and `a:tc/@id` are deliberately absent from the write surface: PowerPoint
// strips both on save (probe:
// `test/read/fixtures/authoring/probe-table-cell-a11y-and-3d.ps1`).

const AT = { x: 1, y: 1, w: 9 }

/** The first `<a:tcPr>` block in the part, whether paired or self-closing. */
function firstTcPr(xml) {
	const match = xml.match(/<a:tcPr(?:\/>|[^>]*>[\s\S]*?<\/a:tcPr>)/)
	assert(match, 'expected an a:tcPr in the part; got: ' + xml)
	return match[0]
}

/** The names `CT_TableCellProperties` declares, in its own sequence order. */
const TCPR_CHILDREN = ['lnL', 'lnR', 'lnT', 'lnB', 'lnTlToBr', 'lnBlToTr', 'cell3D', 'solidFill', 'blipFill', 'headers']

/**
 * The local names of `tcPr`'s **direct** element children, in document order.
 *
 * Every border element wraps its own `a:solidFill` for the stroke colour, so a flat scan
 * would report those as siblings of the cell fill and make a correct part look wrong.
 * Paired children are collapsed to a self-closing stand-in first, which drops their
 * contents and leaves exactly the top level.
 */
function childOrder(tcPr) {
	let flat = tcPr
	for (const name of TCPR_CHILDREN) {
		flat = flat.replace(new RegExp(`<a:${name}\\b[^>]*>[\\s\\S]*?</a:${name}>`, 'g'), `<a:${name}/>`)
	}
	return [...flat.matchAll(/<a:(\w+)[^>]*?\/>/g)].map((m) => m[1]).filter((name) => TCPR_CHILDREN.includes(name))
}

defineRegressionSuite('Table cell a:tcPr constructs', [
	{
		name: 'both diagonals emit after lnB, in schema order',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[
								{
									text: 'X',
									options: {
										border: { type: 'solid', color: '000000', width: 1 },
										diagonal: {
											tlToBr: { type: 'solid', color: 'C00000', width: 2 },
											blToTr: { type: 'dash', color: '0000C0', width: 1 },
										},
										fill: { color: 'EEEEEE' },
									},
								},
							],
						],
						AT
					)
				})
			)

			const tcPr = firstTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(
				childOrder(tcPr).join(','),
				'lnL,lnR,lnT,lnB,lnTlToBr,lnBlToTr,solidFill',
				'CT_TableCellProperties sequence'
			)
			assert(tcPr.includes('<a:lnTlToBr w="25400"'), 'the ╲ diagonal carries its 2pt width; got: ' + tcPr)
			assert(tcPr.includes('<a:lnBlToTr w="12700"'), 'the ╱ diagonal carries its 1pt width; got: ' + tcPr)
			// The diagonals resolve dash and colour through the same path the edges do.
			const blToTr = tcPr.match(/<a:lnBlToTr[\s\S]*?<\/a:lnBlToTr>/)[0]
			assert(blToTr.includes('val="0000C0"'), 'the ╱ diagonal keeps its colour; got: ' + blToTr)
			assert(blToTr.includes('<a:prstDash val="sysDash"/>'), 'and its dash; got: ' + blToTr)
		},
	},
	{
		name: 'one diagonal alone emits only that element',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'X', options: { diagonal: { tlToBr: { type: 'solid' } } } }]], AT)
				})
			)

			const tcPr = firstTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assert(tcPr.includes('<a:lnTlToBr'), 'the requested diagonal is emitted')
			assert(!tcPr.includes('<a:lnBlToTr'), 'the other one is not; got: ' + tcPr)
		},
	},
	{
		name: 'a merged region draws its diagonal once, on the origin only',
		fn: async () => {
			// A covered cell inherits the origin's EDGES so the region's outside is drawn, but a
			// diagonal is a single corner-to-corner stroke: repeating it per covered cell would
			// draw a sawtooth across the merge instead of one line.
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[
								{
									text: 'wide',
									options: {
										colspan: 2,
										border: { type: 'solid', color: '000000' },
										diagonal: { tlToBr: { type: 'solid', color: 'C00000' } },
									},
								},
							],
						],
						AT
					)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assertEqual((xml.match(/<a:lnTlToBr/g) || []).length, 1, 'exactly one diagonal across the merged region')
			const covered = xml.match(/<a:tc hMerge="1"[\s\S]*?<\/a:tc>/)[0]
			assert(covered.includes('<a:lnL'), 'the covered cell still inherits the edges')
			assert(!covered.includes('<a:lnTlToBr'), 'but not the diagonal; got: ' + covered)
		},
	},
	{
		name: 'anchorCtr emits only when true, and lands after anchor',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[
								{ text: 'centred', options: { anchorCtr: true, valign: 'middle' } },
								{ text: 'plain', options: { anchorCtr: false } },
								{ text: 'unset' },
							],
						],
						AT
					)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			const tags = [...xml.matchAll(/<a:tcPr[^>]*>/g)].map((m) => m[0])
			assertEqual(tags.length, 3, 'three cells')
			assert(tags[0].includes('anchorCtr="1"'), 'the true cell carries it; got: ' + tags[0])
			assert(
				tags[0].indexOf('anchor="ctr"') < tags[0].indexOf('anchorCtr="1"'),
				'and it follows anchor, matching the schema adjacency; got: ' + tags[0]
			)
			// `false` is the schema default, so writing it would be noise PowerPoint strips anyway.
			assert(!tags[1].includes('anchorCtr'), 'an explicit false emits nothing; got: ' + tags[1])
			assert(!tags[2].includes('anchorCtr'), 'an unset cell emits nothing; got: ' + tags[2])
		},
	},
	{
		name: 'cell3D emits between the borders and the fill',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[
								{
									text: 'raised',
									options: {
										border: { type: 'solid', color: '000000' },
										cell3D: {
											preset: 'artDeco',
											width: 7,
											height: 7,
											material: 'metal',
											lightRig: { rig: 'threePt', dir: 't' },
										},
										fill: { color: 'DDDDDD' },
									},
								},
							],
						],
						AT
					)
				})
			)

			assert(!codes.includes('table/invalid-cell3d'), 'valid values must not warn; got: ' + JSON.stringify(codes))
			const tcPr = firstTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(childOrder(tcPr).join(','), 'lnL,lnR,lnT,lnB,cell3D,solidFill', 'cell3D sits after the borders')
			assert(tcPr.includes('<a:cell3D prstMaterial="metal">'), 'the material lands on cell3D; got: ' + tcPr)
			// 7pt -> 88900 EMU. The public option is points; the attribute is EMU.
			assert(tcPr.includes('<a:bevel w="88900" h="88900" prst="artDeco"/>'), 'bevel in EMU; got: ' + tcPr)
			assert(tcPr.includes('<a:lightRig rig="threePt" dir="t"/>'), 'light rig; got: ' + tcPr)
		},
	},
	{
		name: 'an empty cell3D still emits the bevel CT_Cell3D requires',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'X', options: { cell3D: {} } }]], AT)
				})
			)

			const tcPr = firstTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			// `a:bevel` is minOccurs=1, so an empty `a:cell3D` would be schema-invalid.
			assert(tcPr.includes('<a:cell3D><a:bevel/></a:cell3D>'), 'bevel is required, so it is always written')
		},
	},
	{
		name: 'an invalid cell3D value is reported and dropped rather than written',
		fn: async () => {
			const { result, codes, diagnostics } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[
								{
									text: 'X',
									options: { cell3D: { preset: /** @type {any} */ ('rounded'), material: 'metal' } },
								},
							],
						],
						AT
					)
				})
			)

			assert(codes.includes('table/invalid-cell3d'), 'expected table/invalid-cell3d; got: ' + JSON.stringify(codes))
			const diagnostic = diagnostics.find((d) => d.code === 'table/invalid-cell3d')
			assertEqual(diagnostic.detail.received, 'rounded', 'the diagnostic names the offending value')

			const tcPr = firstTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assert(!tcPr.includes('prst='), 'the invalid preset never reaches the XML; got: ' + tcPr)
			assert(tcPr.includes('prstMaterial="metal"'), 'the valid sibling still lands; got: ' + tcPr)
		},
	},
	{
		name: 'a half-specified lightRig is dropped whole — both attributes are required',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[[{ text: 'X', options: { cell3D: { lightRig: /** @type {any} */ ({ rig: 'threePt' }) } } }]],
						AT
					)
				})
			)

			assert(codes.includes('table/invalid-cell3d'), 'expected table/invalid-cell3d; got: ' + JSON.stringify(codes))
			const tcPr = firstTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assert(!tcPr.includes('lightRig'), 'a rig missing its dir is not written at all; got: ' + tcPr)
			assert(tcPr.includes('<a:cell3D>'), 'the bevel still stands on its own; got: ' + tcPr)
		},
	},
	{
		name: 'none of the three changes a cell that does not ask for them',
		fn: async () => {
			const { result } = await captureDiagnostics(() => build((p) => p.addSlide().addTable([['A', 'B']], AT)))
			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			for (const marker of ['lnTlToBr', 'lnBlToTr', 'anchorCtr', 'cell3D']) {
				assert(!xml.includes(marker), `${marker} must not appear when unrequested`)
			}
		},
	},
])
