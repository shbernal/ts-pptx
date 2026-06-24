import { defineRegressionSuite, build, readEntry, assert } from '../helpers.js'

// Acceptance for `headerRow` inline sugar (backlog upstream-issue-1256):
// `addTable(rows, { headerRow:{…} })` styles the first row distinctly *without* first
// registering a custom table style. It is applied as direct per-cell formatting on row 0,
// so: (1) only row 0 cells get it, (2) explicit per-cell options win over it, and
// (3) it implies `hasHeader` (emits firstRow="1") unless `hasHeader` is set explicitly.

// Split a table's serialized rows so each `<a:tr>…</a:tr>` can be inspected in isolation.
function tableRows(xml) {
	return [...xml.matchAll(/<a:tr[\s>][\s\S]*?<\/a:tr>/g)].map((m) => m[0])
}

defineRegressionSuite('Table headerRow inline sugar', [
	{
		name: 'headerRow styles only row 0 (fill + bold + color) and implies firstRow="1"',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(
					[
						['A', 'B'],
						['c', 'd'],
					],
					{
						x: 0.5,
						y: 0.5,
						w: 6,
						headerRow: { fill: { color: '1A2B3C' }, color: 'FFFFFF', bold: true },
					}
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')

			// hasHeader auto-set -> firstRow="1" on <a:tblPr>
			assert(
				/<a:tblPr[^>]*firstRow="1"/.test(xml),
				'expected firstRow="1" on <a:tblPr> (hasHeader implied by headerRow)'
			)

			const rows = tableRows(xml)
			assert(rows.length === 2, `expected 2 table rows; got ${rows.length}`)
			const [header, body] = rows

			// Header row carries the header fill, bold, and color
			assert(header.includes('1A2B3C'), 'header row should carry headerRow fill color 1A2B3C')
			assert(header.includes('FFFFFF'), 'header row should carry headerRow text color FFFFFF')
			assert(/b="1"/.test(header), 'header row should be bold (b="1")')

			// Body row must NOT inherit the header styling
			assert(!body.includes('1A2B3C'), 'body row must not carry the header fill color')
			assert(!body.includes('FFFFFF'), 'body row must not carry the header text color')
			assert(!/b="1"/.test(body), 'body row must not be bold')
		},
	},
	{
		name: 'explicit per-cell options on a row-0 cell win over headerRow',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable([[{ text: 'A', options: { fill: { color: 'AA0000' } } }, 'B']], {
					x: 0.5,
					y: 0.5,
					w: 6,
					headerRow: { fill: { color: '1A2B3C' }, color: 'FFFFFF', bold: true },
				})
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			const [header] = tableRows(xml)

			// Cell A keeps its explicit fill; cell B falls back to headerRow fill
			assert(header.includes('AA0000'), 'explicit per-cell fill AA0000 must survive')
			assert(header.includes('1A2B3C'), 'cell without explicit fill still gets headerRow fill 1A2B3C')
			// Both cells still inherit headerRow color/bold (not overridden per cell)
			assert(/b="1"/.test(header), 'headerRow bold still applies where not overridden')
		},
	},
	{
		name: 'explicit hasHeader:false is respected (no firstRow="1") while styling still applies',
		fn: async () => {
			const { zip } = await build((p) => {
				const s = p.addSlide()
				s.addTable(
					[
						['A', 'B'],
						['c', 'd'],
					],
					{
						x: 0.5,
						y: 0.5,
						w: 6,
						hasHeader: false,
						headerRow: { fill: { color: '1A2B3C' } },
					}
				)
			})
			const xml = await readEntry(zip, 'ppt/slides/slide1.xml')
			assert(!/firstRow="1"/.test(xml), 'explicit hasHeader:false must not emit firstRow="1"')
			const [header] = tableRows(xml)
			assert(header.includes('1A2B3C'), 'headerRow styling still applies even with hasHeader:false')
		},
	},
])
