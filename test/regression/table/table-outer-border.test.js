import { defineRegressionSuite, build, readEntry, captureDiagnostics, assert, assertEqual } from '../../helpers.js'

// `TableProps.outerBorder` -> the table's perimeter only.
//
// `TableProps.border` reads like a perimeter and is not one: it is a per-cell default, so
// `[solid, none, solid, none]` gives EVERY cell a top and bottom rule. `outerBorder` is the
// perimeter -- the top edge of row 0, the bottom edge of the last row, the left edge of
// column 0 and the right edge of the last column -- and it is applied last, so it wins on
// the sides it touches and leaves the interior to `border`.
//
// The perimeter is decided by GRID position, not by authored cell. That is what makes
// merges work: PowerPoint defines a merged region's outer edges on the *covered* cells, so
// a colspan reaching the last column gets that column's rule on its `hMerge` dummy.

const AT = { x: 1, y: 1, w: 9 }

const SOLID = { type: 'solid', color: 'FF0000', width: 2 }

/** Each `<a:tc>` in the part, in document (row-major) order. */
function cells(xml) {
	return xml.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g) || []
}

/**
 * The four edges of one cell as `'solid'` / `'none'`, in the public TRBL order.
 * `genTableCellBorderXml` writes them in LRTB document order, which this undoes.
 */
function edges(cellXml) {
	const read = (name) => {
		const block = cellXml.match(new RegExp(`<a:${name}\\b[\\s\\S]*?</a:${name}>`))
		if (!block) return 'absent'
		return block[0].includes('<a:noFill/>') ? 'none' : 'solid'
	}
	return [read('lnT'), read('lnR'), read('lnB'), read('lnL')].join(',')
}

defineRegressionSuite('Table outerBorder', [
	{
		name: 'a bare BorderProps boxes the table and leaves the interior clear',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							['A', 'B', 'C'],
							['D', 'E', 'F'],
							['G', 'H', 'I'],
						],
						{ ...AT, outerBorder: SOLID }
					)
				})
			)

			const tcs = cells(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(tcs.length, 9, '3x3 grid')
			// TRBL per cell. Corners take two sides, edges one, the middle none.
			const expected = [
				'solid,none,none,solid', 'solid,none,none,none', 'solid,solid,none,none',
				'none,none,none,solid',  'none,none,none,none',  'none,solid,none,none',
				'none,none,solid,solid', 'none,none,solid,none', 'none,solid,solid,none',
			] // prettier-ignore
			assertEqual(tcs.map(edges).join(' | '), expected.join(' | '), 'only the perimeter is drawn')
		},
	},
	{
		name: 'a sparse tuple rules above and below without touching the sides',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							['A', 'B'],
							['C', 'D'],
						],
						// TRBL, with the right and left entries left out: those edges keep whatever
						// `border` (here: nothing) already put there.
						{ ...AT, outerBorder: [SOLID, undefined, SOLID, undefined] }
					)
				})
			)

			const tcs = cells(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(
				tcs.map(edges).join(' | '),
				['solid,none,none,none', 'solid,none,none,none', 'none,none,solid,none', 'none,none,solid,none'].join(' | '),
				'only the top and bottom perimeter sides are drawn'
			)
		},
	},
	{
		name: 'perimeter and per-cell interior grid compose',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							['A', 'B'],
							['C', 'D'],
						],
						// `border` draws a full grid on every cell; `outerBorder` overrides the outside.
						{ ...AT, border: { type: 'solid', color: 'CCCCCC', width: 0.5 }, outerBorder: SOLID }
					)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			const tcs = cells(xml)
			// Every side of every cell is drawn -- the perimeter only changes *which* rule.
			for (const [idx, tc] of tcs.entries()) {
				assertEqual(edges(tc), 'solid,solid,solid,solid', `cell ${idx} keeps all four rules`)
			}
			// The top-left cell's top and left take the 2pt red perimeter; its right and bottom
			// keep the 0.5pt grey grid. Widths are in EMU-ish points (`valToPts`), colours in hex.
			const first = tcs[0]
			const perimeter = [...first.matchAll(/<a:(ln[LRTB])[^>]*>[\s\S]*?<a:srgbClr val="([^"]*)"/g)].map(
				(m) => `${m[1]}=${m[2]}`
			)
			assertEqual(
				perimeter.join(','),
				'lnL=FF0000,lnR=CCCCCC,lnT=FF0000,lnB=CCCCCC',
				'the perimeter wins only on the sides it touches'
			)
		},
	},
	{
		name: "a colspan in the corner puts the last column's rule on the covered cell",
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							// One cell spanning both columns of a 2-column table: its origin sits in
							// column 0 and PowerPoint defines the region's right edge on the hMerge dummy.
							[{ text: 'wide', options: { colspan: 2 } }],
							['C', 'D'],
						],
						{ ...AT, outerBorder: SOLID }
					)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			const tcs = cells(xml)
			assertEqual(tcs.length, 4, 'the merge grid is rectangular: 2 rows x 2 columns')
			assert(tcs[1].includes('hMerge="1"'), 'the second cell of row 0 is the covered half of the span')
			assertEqual(edges(tcs[0]), 'solid,none,none,solid', 'the origin takes the top and left')
			assertEqual(edges(tcs[1]), 'solid,solid,none,none', 'the covered cell takes the top and the right')
		},
	},
	{
		name: 'a rowspan reaching the last row puts the bottom rule on the covered cell',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'tall', options: { rowspan: 2 } }, 'B'], ['D']], {
						...AT,
						outerBorder: SOLID,
					})
				})
			)

			const tcs = cells(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(tcs.length, 4, 'the merge grid is rectangular')
			assert(tcs[2].includes('vMerge="1"'), 'the first cell of row 1 is the covered half of the span')
			assertEqual(edges(tcs[0]), 'solid,none,none,solid', 'the origin takes the top and left')
			assertEqual(edges(tcs[2]), 'none,none,solid,solid', 'the covered cell takes the left and the bottom')
		},
	},
	{
		name: 'an unset outerBorder changes nothing',
		fn: async () => {
			const rows = [
				['A', 'B'],
				['C', 'D'],
			]
			const { result: withOut } = await captureDiagnostics(() =>
				build((p) => p.addSlide().addTable(rows, { ...AT, outerBorder: undefined }))
			)
			const { result: without } = await captureDiagnostics(() => build((p) => p.addSlide().addTable(rows, AT)))

			assertEqual(
				await readEntry(withOut.zip, 'ppt/slides/slide1.xml'),
				await readEntry(without.zip, 'ppt/slides/slide1.xml'),
				'the no-perimeter path is byte-identical'
			)
		},
	},
	{
		name: 'repeated writes do not accumulate the perimeter onto the cells',
		fn: async () => {
			// `arrTabRows` holds the caller's own cell objects, so a perimeter applied by mutation
			// would compound across write() calls and leak onto interior cells the second time.
			const pres = await captureDiagnostics(async () => {
				const built = await build((p) => {
					p.addSlide().addTable(
						[
							['A', 'B'],
							['C', 'D'],
						],
						{ ...AT, outerBorder: SOLID }
					)
				})
				const second = /** @type {Uint8Array} */ (await built.pres.toBytes())
				return { first: await readEntry(built.zip, 'ppt/slides/slide1.xml'), second }
			})

			const JSZip = (await import('jszip')).default
			const zip2 = await JSZip.loadAsync(pres.result.second)
			assertEqual(
				await readEntry(zip2, 'ppt/slides/slide1.xml'),
				pres.result.first,
				'a second write produces the same XML'
			)
		},
	},
])
