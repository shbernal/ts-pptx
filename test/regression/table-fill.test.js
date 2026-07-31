import { defineRegressionSuite, build, readEntry, captureDiagnostics, assert, assertEqual } from '../helpers.js'

// The two table-level fill options, and the non-solid cell fills that were emitting all
// along with nothing documenting or pinning them.
//
// `TableProps.fill` and `TableProps.tableFill` look interchangeable and are not:
//   - `fill` is STAMPED onto every cell, so each cell ends up explicitly filled.
//   - `tableFill` writes one `a:tblPr` fill that the cells sit on top of.
// They render alike on a plain table, which is exactly why the difference needs pinning:
// with `fill` there is no such thing as an unfilled cell, so a cell can never fall back to
// a table background, and a deck read back from PowerPoint carries the second shape.
//
// `CT_TableProperties` sequences EG_FillProperties before the tableStyle choice, so a
// background must precede `a:tableStyleId` or PowerPoint reports the part as corrupt.

const AT = { x: 1, y: 1, w: 9 }

/** The `<a:tblPr>` block, paired or self-closing. */
function tblPr(xml) {
	const match = xml.match(/<a:tblPr(?:\/>|[^>]*>[\s\S]*?<\/a:tblPr>)/)
	assert(match, 'expected an a:tblPr in the part; got: ' + xml)
	return match[0]
}

/**
 * Each cell's `<a:tcPr>` in document order, with its border elements stripped.
 *
 * Two other things in a cell carry an `a:solidFill` and neither is the cell's fill: every
 * border wraps one for its stroke colour, and every text run wraps one for its font colour.
 * Narrowing to `a:tcPr` drops the runs, and stripping the `a:lnX` blocks drops the strokes,
 * so what is left is exactly the cell's own fill.
 */
function cells(xml) {
	return (xml.match(/<a:tcPr(?:\/>|[^>]*>[\s\S]*?<\/a:tcPr>)/g) || []).map((tcPr) =>
		tcPr.replace(/<a:ln(?:L|R|T|B|TlToBr|BlToTr)\b[\s\S]*?<\/a:ln(?:L|R|T|B|TlToBr|BlToTr)>/g, '')
	)
}

defineRegressionSuite('Table fill', [
	{
		name: 'tableFill lands on a:tblPr and leaves every cell unfilled',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							['A', 'B'],
							['C', 'D'],
						],
						{ ...AT, tableFill: { color: 'F2F2F2' } }
					)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assert(tblPr(xml).includes('<a:srgbClr val="F2F2F2"/>'), 'the background is on a:tblPr; got: ' + tblPr(xml))
			for (const [idx, tc] of cells(xml).entries()) {
				assert(!tc.includes('<a:solidFill>'), `cell ${idx} is left unfilled so the background shows through`)
			}
		},
	},
	{
		name: 'fill is stamped onto every cell and never reaches a:tblPr',
		fn: async () => {
			// The contrast case for the above. Same visual result, different package — and the
			// difference is what makes a cell able (or unable) to fall back to a background.
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							['A', 'B'],
							['C', 'D'],
						],
						{ ...AT, fill: { color: 'F2F2F2' } }
					)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assert(!tblPr(xml).includes('srgbClr'), 'nothing reaches a:tblPr; got: ' + tblPr(xml))
			const filled = cells(xml).filter((tc) => tc.includes('<a:srgbClr val="F2F2F2"/>'))
			assertEqual(filled.length, 4, 'every cell carries its own copy of the colour')
		},
	},
	{
		name: 'a background precedes a:tableStyleId, per the CT_TableProperties sequence',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					const style = p.defineTableStyle({ name: 'Brand', firstRow: { fill: '1A2B3C', bold: true } })
					p.addSlide().addTable([['A']], { ...AT, tableFill: { color: 'F2F2F2' }, tableStyle: style, hasHeader: true })
				})
			)

			const block = tblPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			const fillAt = block.indexOf('<a:solidFill>')
			const styleAt = block.indexOf('<a:tableStyleId>')
			assert(fillAt !== -1 && styleAt !== -1, 'both children are present; got: ' + block)
			assert(fillAt < styleAt, 'the fill comes first, as the schema sequences it; got: ' + block)
		},
	},
	{
		name: 'the two compose — a background behind, per-cell overrides on top',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'own', options: { fill: { color: 'FF0000' } } }, { text: 'inherits' }]], {
						...AT,
						tableFill: { color: 'F2F2F2' },
					})
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assert(tblPr(xml).includes('val="F2F2F2"'), 'the background is on the table')
			const tcs = cells(xml)
			assert(tcs[0].includes('<a:srgbClr val="FF0000"/>'), 'the explicit cell keeps its own fill')
			assert(!tcs[1].includes('<a:solidFill>'), 'the other cell shows the background through')
		},
	},
	{
		name: 'a gradient and a pattern cell fill both emit',
		fn: async () => {
			// These have always emitted -- `genXmlColorSelection` handles the whole fill group --
			// but nothing documented or pinned them, so "works" and "happens to work" were
			// indistinguishable. Now pinned.
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(
						[
							[
								{
									text: 'grad',
									options: {
										fill: {
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
						AT
					)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			const tcs = cells(xml)
			assert(tcs[0].includes('<a:gradFill'), 'the gradient cell emits a:gradFill; got: ' + tcs[0])
			assert(tcs[0].includes('<a:lin ang="5400000"'), 'with its angle in 60000ths of a degree; got: ' + tcs[0])
			assert(tcs[1].includes('<a:pattFill prst="diagCross">'), 'the pattern cell emits a:pattFill; got: ' + tcs[1])
		},
	},
	{
		name: 'a gradient table background emits too',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([['A']], {
						...AT,
						tableFill: {
							type: 'gradient',
							gradient: {
								kind: 'linear',
								angle: 0,
								stops: [
									{ position: 0, color: 'FFFFFF' },
									{ position: 100, color: 'DDDDDD' },
								],
							},
						},
					})
				})
			)

			const block = tblPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assert(block.includes('<a:gradFill'), 'a:tblPr takes the whole fill group, not just solid; got: ' + block)
		},
	},
	{
		name: 'an unset tableFill leaves a:tblPr self-closing',
		fn: async () => {
			const { result } = await captureDiagnostics(() => build((p) => p.addSlide().addTable([['A']], AT)))
			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assert(tblPr(xml).endsWith('/>'), 'the no-background path keeps the empty-element form; got: ' + tblPr(xml))
		},
	},
])
