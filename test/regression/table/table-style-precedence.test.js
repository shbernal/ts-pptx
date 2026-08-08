import {
	TsPptx,
	defineRegressionSuite,
	build,
	readEntry,
	captureDiagnostics,
	assert,
	assertEqual,
} from '../../helpers.js'
import { TableStyle } from '../../../dist/node.js'

// What a table style can and cannot do, and which tier of the styling stack actually paints.
//
// The load-bearing fact, measured by rendering in PowerPoint desktop 16.0 rather than read off
// the schema: PowerPoint resolves `<a:tableStyleId>` against its OWN table-style gallery and
// never reads a style definition out of the package. A built-in GUID paints even when the deck
// defines nothing; a GUID PowerPoint does not recognise paints nothing even when the deck
// defines it perfectly -- rewriting one style's GUID to a novel value in a PowerPoint-authored
// deck (both the styles part and the slide, bytes otherwise identical) drops that table to the
// no-style look while its untouched neighbours keep theirs. The same holds for a definition
// placed inline in `<a:tblPr>` as `<a:tableStyle>`, and for one nominated by the part's `def=`.
//
// So there are exactly two ways to style a table, and these cases pin both:
//   1. a built-in `TableStyle` member, which emits a GUID PowerPoint knows; or
//   2. direct per-cell formatting -- `headerRow`, `columns[i]`, the table-level `border` /
//      `color` / `fill`, or a cell's own `options` -- which is what carries every brand-coloured
//      table.
//
// The per-cell defaults are part of that second mechanism, not an accident: every cell gets a
// `{type:'none'}` border on all four sides and black text as *direct* formatting, and direct
// formatting outranks a style region. That is what keeps an unstyled table free of grid lines.
//
// `defineTableStyle()` and `TableProps.styleDrivenCells` used to live here. Both were removed
// once the render evidence showed a custom style is unreachable markup in PowerPoint whatever
// it says and wherever it sits -- see CHANGELOG.md and `docs/tables.md` -> "Table styles".

const AT = { x: 1, y: 1, w: 8 }

/** The first cell's `<a:tcPr>` block. */
function firstTcPr(xml) {
	const match = xml.match(/<a:tcPr(?:\/>|[^>]*>[\s\S]*?<\/a:tcPr>)/)
	assert(match, 'expected an a:tcPr; got: ' + xml)
	return match[0]
}

/** The `<a:rPr>` of the first text run. */
function firstRPr(xml) {
	const match = xml.match(/<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/)
	assert(match, 'expected an a:rPr; got: ' + xml)
	return match[0]
}

/** Every cell's `<a:tcPr>` block, in document order. */
function allTcPr(xml) {
	const blocks = xml.match(/<a:tcPr(?:\/>|[^>]*>[\s\S]*?<\/a:tcPr>)/g) || []
	assert(blocks.length, 'expected at least one a:tcPr; got: ' + xml)
	return blocks
}

/** Build one table and return the slide part alongside `tableStyles.xml`. */
async function tableParts(tableOpts = {}, rows = [[{ text: 'H1' }]]) {
	const { zip } = await build((p) => {
		p.addSlide().addTable(rows, { ...AT, hasHeader: true, ...tableOpts })
	})
	return {
		slide: await readEntry(zip, 'ppt/slides/slide1.xml'),
		styles: await readEntry(zip, 'ppt/tableStyles.xml'),
	}
}

defineRegressionSuite('Table styling: built-in styles and the direct-formatting tiers', [
	{
		name: 'a built-in TableStyle emits its GUID as the table-s tableStyleId',
		fn: async () => {
			// The one mechanism PowerPoint honours. The GUID is all that goes on the wire -- the
			// definition lives in PowerPoint's gallery, not in the deck.
			const { slide } = await tableParts({ tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1 })
			assert(
				slide.includes(`<a:tableStyleId>${TableStyle.MEDIUM_STYLE_2_ACCENT_1}</a:tableStyleId>`),
				'the built-in GUID is referenced; got: ' + slide
			)
		},
	},
	{
		name: 'tableStyles.xml is always the bare default-id stub, whatever the deck builds',
		fn: async () => {
			// The part ships because PowerPoint expects the relationship and the content-type
			// override, but it defines nothing: a definition here is markup PowerPoint never reads.
			// If a future change starts writing `<a:tblStyle>` again, this is what says so.
			const { styles } = await tableParts({ tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1 })
			assert(!styles.includes('<a:tblStyle '), 'no style definition is emitted; got: ' + styles)
			assert(styles.includes('<a:tblStyleLst'), 'the list element is still there; got: ' + styles)
			assert(styles.includes(`def="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}"`), 'and names a default; got: ' + styles)
		},
	},
	{
		name: 'every cell is given an explicit four-side none border as direct formatting',
		fn: async () => {
			// Not an oversight: this is what stops a table inheriting grid lines, and it is the
			// bottom rung of the precedence stack every case below sits on top of.
			const { slide } = await tableParts()
			const tcPr = firstTcPr(slide)
			assertEqual((tcPr.match(/<a:ln[LRTB] w="0"/g) || []).length, 4, 'all four sides are explicitly suppressed')
		},
	},
	{
		name: 'and black text, likewise as direct formatting',
		fn: async () => {
			const { slide } = await tableParts()
			assert(firstRPr(slide).includes('<a:srgbClr val="000000"/>'), 'black is stamped; got: ' + firstRPr(slide))
		},
	},
	{
		name: 'a table-level border reaches every cell',
		fn: async () => {
			// Tier 4: the table-level option flows into the same per-cell slot the default fills.
			const { slide } = await tableParts({ border: { type: 'solid', color: 'D9D9D9', width: 0.5 } })
			const tcPr = firstTcPr(slide)
			assert(tcPr.includes('<a:srgbClr val="D9D9D9"/>'), 'the table-level border reaches the cell; got: ' + tcPr)
		},
	},
	{
		name: 'headerRow supplies the header text colour, beating the black default',
		fn: async () => {
			// Tier 2, and the documented way to colour a header now that no region can.
			const { slide } = await tableParts({ headerRow: { color: 'FFFFFF', fill: '1A2B3C', bold: true } })
			const rPr = firstRPr(slide)
			assert(rPr.includes('<a:srgbClr val="FFFFFF"/>'), 'headerRow beats the default; got: ' + rPr)
			assert(firstTcPr(slide).includes('1A2B3C'), 'and its fill lands on the cell; got: ' + firstTcPr(slide))
		},
	},
	{
		name: 'columns[i] fills its own column and leaves the others alone',
		fn: async () => {
			const { slide } = await tableParts({ hasHeader: false, columns: [undefined, { fill: 'EAF1F8' }] }, [
				[{ text: 'A' }, { text: 'B' }],
			])
			const [first, second] = allTcPr(slide)
			assert(!first.includes('EAF1F8'), 'column 0 is untouched; got: ' + first)
			assert(second.includes('EAF1F8'), 'column 1 takes its fill; got: ' + second)
		},
	},
	{
		name: 'a cell-s own options beat headerRow, which beats the table level',
		fn: async () => {
			// The full stack in one case: cell > headerRow > table > default.
			const { slide } = await tableParts({ color: '112233', headerRow: { color: 'AABBCC' } }, [
				[{ text: 'A', options: { color: '445566' } }, { text: 'B' }],
			])
			const [first, second] = allTcPr(slide).map((_block, idx) => {
				const runs = slide.match(/<a:rPr[^>]*(?:\/>|>[\s\S]*?<\/a:rPr>)/g) || []
				return runs[idx]
			})
			assert(first.includes('445566'), "the cell's own colour wins; got: " + first)
			assert(second.includes('AABBCC'), 'and headerRow carries the rest of the row; got: ' + second)
		},
	},
	{
		name: 'a hyperlink anywhere in the grid stands the black default down',
		fn: async () => {
			// Long-standing carve-out, unrelated to table styles: the default paints the whole run,
			// so the words *after* a link would come out black instead of following the link colour.
			const { slide } = await tableParts({}, [
				[{ text: 'H1' }, { text: 'docs', options: { hyperlink: { url: 'https://example.com' } } }],
			])
			assert(slide.includes('<a:hlinkClick'), 'the hyperlink is emitted; got: ' + slide)
			assert(!slide.includes('<a:srgbClr val="000000"/>'), 'and nothing is painted black; got: ' + slide)
		},
	},
	{
		name: 'outerBorder draws the perimeter over the suppressed defaults',
		fn: async () => {
			const { slide } = await tableParts(
				{ hasHeader: false, outerBorder: { type: 'solid', color: '1A2B3C', width: 1 } },
				[
					[{ text: 'A' }, { text: 'B' }],
					[{ text: 'C' }, { text: 'D' }],
				]
			)
			const topLeft = allTcPr(slide)[0]
			assertEqual((topLeft.match(/1A2B3C/g) || []).length, 2, 'the two perimeter sides are the outer colour')
			assert(/<a:lnR w="0"/.test(topLeft) && /<a:lnB w="0"/.test(topLeft), 'interior sides stay suppressed')
		},
	},
	{
		name: 'defineTableStyle is gone from the public surface',
		fn: async () => {
			// The removal is the API contract now: a consumer still calling it must fail loudly at
			// the call site rather than silently building a deck whose styling never paints.
			// TypeScript already refuses the call — hence the cast, which is the point rather than a
			// workaround — and this pins the runtime half for a plain-JS consumer.
			const pres = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (new TsPptx()))
			assertEqual(typeof pres.defineTableStyle, 'undefined', 'no defineTableStyle method')
		},
	},
	{
		name: 'styleDrivenCells is inert and silent — an unknown option, not a supported one',
		fn: async () => {
			// It was never anything but a way to defer to a custom style, so with custom styles gone
			// it has nothing to mean. Passing it must not change the bytes.
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'H1' }]], { ...AT, styleDrivenCells: true })
				})
			)
			const tcPr = allTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))[0]
			assertEqual((tcPr.match(/<a:ln[LRTB] w="0"/g) || []).length, 4, 'the defaults are untouched')
			assert(!codes.some((code) => code.startsWith('table-style/')), 'no stale code fires; got: ' + codes.join(','))
		},
	},
])
