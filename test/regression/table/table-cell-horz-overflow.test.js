import { defineRegressionSuite, build, readEntry, captureDiagnostics, assert, assertEqual } from '../../helpers.js'

// `TableCellProps.horzOverflow` -> `a:tcPr/@horzOverflow`.
//
// The attribute is easy to misread as a text-wrap switch, and it is not one: per ECMA-376
// §20.1.10.68 it decides what happens to a single GLYPH too wide for the line -- `clip`
// (the default) cuts it at the cell edge, `overflow` lets it draw past. Cell text always
// wraps to the column width; PowerPoint has no per-cell no-wrap at all, which
// `test/read/fixtures/authoring/probe-table-cell-wrap.ps1` establishes by showing
// `<a:bodyPr wrap="none"/>` in a cell render inert and then get stripped on the next save.
//
// The read-side oracle for the same attribute is `table-cell-horzoverflow.pptx`, whose
// bytes PowerPoint itself wrote.
//
// Values are checked before emission rather than trusted: `ST_TextHorzOverflowType` admits
// only those two, and anything else would make the slide part schema-invalid -- which
// PowerPoint reports as a corrupt file, not as a mis-set option.

/** Every `<a:tcPr …>` opening tag in the part, in document order. */
function tcPrTags(xml) {
	return [...xml.matchAll(/<a:tcPr[^>]*>/g)].map((m) => m[0])
}

const AT = { x: 1, y: 1, w: 9 }

defineRegressionSuite('Table cell horzOverflow', [
	{
		name: 'overflow reaches the authored cell, and only that cell',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'A', options: { horzOverflow: 'overflow' } }, { text: 'B' }]], AT)
				})
			)

			assert(
				!codes.includes('table/invalid-horz-overflow'),
				'a valid value must not warn; got: ' + JSON.stringify(codes)
			)

			const tags = tcPrTags(await readEntry(result.zip, 'ppt/slides/slide1.xml'))
			assertEqual(tags.length, 2, 'two cells -> two a:tcPr')
			assert(tags[0].includes('horzOverflow="overflow"'), 'authored cell carries the attribute; got: ' + tags[0])
			assert(!tags[1].includes('horzOverflow'), 'the sibling cell stays untouched; got: ' + tags[1])
		},
	},
	{
		name: 'an unset horzOverflow emits nothing (the default path is byte-unchanged)',
		fn: async () => {
			const { result } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'A' }, { text: 'B' }]], AT)
				})
			)

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assert(!xml.includes('horzOverflow'), 'no attribute may appear when the option is unset')
		},
	},
	{
		name: 'an unrecognized value is reported and dropped rather than written',
		fn: async () => {
			const { result, codes, diagnostics } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'A', options: { horzOverflow: /** @type {any} */ ('ellipsis') } }]], AT)
				})
			)

			assert(
				codes.includes('table/invalid-horz-overflow'),
				'expected the table/invalid-horz-overflow code; got: ' + JSON.stringify(codes)
			)
			const diagnostic = diagnostics.find((d) => d.code === 'table/invalid-horz-overflow')
			assertEqual(diagnostic.detail.received, 'ellipsis', 'the diagnostic names the offending value')

			const xml = await readEntry(result.zip, 'ppt/slides/slide1.xml')
			assert(!xml.includes('horzOverflow'), 'the invalid value must not reach the XML')
		},
	},
])
