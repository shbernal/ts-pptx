import { TsPptx, defineRegressionSuite, build, readEntry, captureDiagnostics, assert, assertEqual } from '../helpers.js'
import { TableStyle } from '../../dist/node.js'

// Which of a custom table style's region properties can actually render, and which the
// library's own per-cell defaults override.
//
// This is the most surprising thing about the table surface and it had never been pinned.
// A region (`CT_TablePartStyle`) can carry exactly five things: `fill`, `color`, `bold`,
// `italic` and `border`. ts-pptx stamps a default `border` (`{type:'none'}` on all four
// sides) and `color` (`'000000'`) onto EVERY cell as *direct* formatting, and in PowerPoint
// direct cell formatting outranks a style region -- so those two are overridden before they
// can render, while `fill`, `bold` and `italic` are not defaulted and do apply.
//
// The stamped `fontSize` and `margin` defaults are NOT part of this: a table style region
// has nowhere to put either one (`tcTxStyle` carries only b/i/font-ref/color, `tcStyle` only
// tcBdr/fill/cell3D), so there is no region value for them to override. Don't add cases for
// them here expecting a deferral -- see the `defineTableStyle` doc comment.
//
// The result is a half-working style rather than an obviously broken one, which is exactly
// why it needs a test: `defineTableStyle({ firstRow: { fill, color, bold } })` shades and
// bolds the header and silently ignores the text colour. `docs/tables.md` documents this,
// and these cases are what stop the docs and the behaviour drifting apart.

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

/** Build a table under a custom style and return the slide part alongside `tableStyles.xml`. */
async function withStyleParts(styleDef, tableOpts = {}, rows = [[{ text: 'H1' }]]) {
	const { zip } = await build((p) => {
		const guid = p.defineTableStyle({ name: 'Brand', ...styleDef })
		p.addSlide().addTable(rows, { ...AT, tableStyle: guid, hasHeader: true, ...tableOpts })
	})
	return {
		slide: await readEntry(zip, 'ppt/slides/slide1.xml'),
		styles: await readEntry(zip, 'ppt/tableStyles.xml'),
	}
}

/** Build a one-cell table under a custom style and return its slide part. */
async function withStyle(styleDef, tableOpts = {}) {
	return (await withStyleParts(styleDef, tableOpts)).slide
}

defineRegressionSuite('Table style precedence against per-cell defaults', [
	{
		name: "a style region's fill and bold reach the cell — nothing defaults them",
		fn: async () => {
			const xml = await withStyle({ firstRow: { fill: '1A2B3C', bold: true } })
			// The cell carries no fill of its own, so the style's shading is what renders.
			assert(!firstTcPr(xml).includes('<a:solidFill>'), 'no per-cell fill is stamped; got: ' + firstTcPr(xml))
			// Likewise no `b=` on the run, so the style's bold is what renders.
			assert(!firstRPr(xml).includes(' b="'), 'no per-cell bold is stamped; got: ' + firstRPr(xml))
		},
	},
	{
		name: "a style region's border is overridden by the per-cell default",
		fn: async () => {
			const xml = await withStyle({ wholeTbl: { border: { type: 'solid', color: 'D9D9D9', width: 0.5 } } })
			const tcPr = firstTcPr(xml)
			// Every side is an explicit w="0" noFill — direct formatting, which beats the region.
			assertEqual((tcPr.match(/<a:ln[LRTB] w="0"/g) || []).length, 4, 'all four sides are explicitly suppressed')
			assert(!tcPr.includes('D9D9D9'), "the style's border colour never reaches the cell; got: " + tcPr)
		},
	},
	{
		name: "a style region's text colour is overridden by the per-cell default",
		fn: async () => {
			const xml = await withStyle({ firstRow: { color: 'FFFFFF', bold: true } })
			const rPr = firstRPr(xml)
			assert(rPr.includes('<a:srgbClr val="000000"/>'), 'the black default is stamped on the run; got: ' + rPr)
			assert(!rPr.includes('FFFFFF'), "the style's white never reaches the run; got: " + rPr)
		},
	},
	{
		name: 'setting the border on the TABLE is the way to get one alongside a style',
		fn: async () => {
			// The documented workaround, and the reason the workaround exists: the table-level
			// option flows into the same per-cell slot the default would otherwise fill.
			const xml = await withStyle(
				{ firstRow: { fill: '1A2B3C' } },
				{ border: { type: 'solid', color: 'D9D9D9', width: 0.5 } }
			)
			const tcPr = firstTcPr(xml)
			assert(tcPr.includes('<a:srgbClr val="D9D9D9"/>'), 'the table-level border reaches the cell; got: ' + tcPr)
			assert(!tcPr.includes('<a:solidFill>\n'), 'and the style still supplies the fill')
		},
	},
	{
		name: 'setting the colour on headerRow is the way to get header text colour',
		fn: async () => {
			const xml = await withStyle({ firstRow: { fill: '1A2B3C' } }, { headerRow: { color: 'FFFFFF' } })
			const rPr = firstRPr(xml)
			assert(rPr.includes('<a:srgbClr val="FFFFFF"/>'), 'headerRow beats the default; got: ' + rPr)
		},
	},
	{
		name: 'defineTableStyle warns for a region border, naming the table-level alternative',
		fn: async () => {
			const { codes, messages } = await captureDiagnostics(() =>
				new TsPptx().defineTableStyle({ name: 'Brand', wholeTbl: { border: { type: 'solid' } } })
			)
			assertEqual(codes.join(','), 'table-style/region-overridden', 'one diagnostic, keyed on the region code')
			// The code is the API; this asserts only that the message points somewhere useful,
			// since a warning that names the dead end without the way out is not worth emitting.
			assert(messages[0].includes('`border` on the table'), 'names the remedy; got: ' + messages[0])
		},
	},
	{
		name: 'defineTableStyle warns for a region colour',
		fn: async () => {
			const { codes } = await captureDiagnostics(() =>
				new TsPptx().defineTableStyle({ name: 'Brand', firstRow: { color: 'FFFFFF' } })
			)
			assertEqual(codes.join(','), 'table-style/region-overridden', 'the colour case warns too')
		},
	},
	{
		name: 'defineTableStyle stays silent for the region properties that do render',
		fn: async () => {
			// fill, bold and italic are not defaulted per cell, so they reach the render and must
			// not be warned about -- a warning here would train callers to ignore the code.
			const { codes } = await captureDiagnostics(() =>
				new TsPptx().defineTableStyle({
					name: 'Brand',
					firstRow: { fill: '1A2B3C', bold: true, italic: true },
					band1H: { fill: 'EAF1F8' },
				})
			)
			assertEqual(codes.join(','), '', 'no diagnostic for fill/bold/italic; got: ' + JSON.stringify(codes))
		},
	},
	{
		name: 'defineTableStyle warns once per offending region, not once per style',
		fn: async () => {
			// Each region is a separate authoring mistake with a separate fix, and the process-wide
			// warnOnce dedupe is deliberately not used here: a second presentation in the same
			// process (a batch deck build, or this suite) must still hear about its own style.
			const { codes } = await captureDiagnostics(() =>
				new TsPptx().defineTableStyle({
					name: 'Brand',
					wholeTbl: { border: { type: 'solid' } },
					firstRow: { color: 'FFFFFF' },
					lastRow: { color: '000000', fill: 'EEEEEE' },
				})
			)
			assertEqual(codes.length, 3, 'one per offending region property; got: ' + JSON.stringify(codes))
		},
	},
	{
		name: 'a slide resolves a registered style GUID to the definition behind it',
		fn: async () => {
			const pres = new TsPptx()
			const guid = pres.defineTableStyle({ name: 'Brand', firstRow: { fill: '1A2B3C' } })
			const slide = pres.addSlide()
			// The definition, not a boolean: deciding whether a per-cell default may stand aside
			// needs to know what the style actually says, not merely that one exists.
			assertEqual(slide.getCustomTableStyle(guid)?.name, 'Brand', 'the registry answers by GUID')
			assertEqual(slide.getCustomTableStyle(guid)?.firstRow?.fill, '1A2B3C', 'and hands back the region')
		},
	},
	{
		name: 'a built-in style GUID resolves to nothing, and so does an unregistered one',
		fn: async () => {
			// The whole reason the lookup is a registry hit rather than a test on the GUID's
			// shape: `defineTableStyle()` mints the same `{XXXXXXXX-...}` form the built-ins use.
			// Office's built-ins define borders of their own, so they must keep every per-cell
			// default exactly as it is -- keying off "a tableStyle is set" would put grid lines
			// into every deck using MEDIUM_STYLE_2_ACCENT_1.
			const pres = new TsPptx()
			pres.defineTableStyle({ name: 'Brand', firstRow: { fill: '1A2B3C' } })
			const slide = pres.addSlide()
			assertEqual(slide.getCustomTableStyle(TableStyle.MEDIUM_STYLE_2_ACCENT_1), undefined, 'built-in')
			assertEqual(slide.getCustomTableStyle('{00000000-0000-0000-0000-000000000000}'), undefined, 'unknown')
		},
	},
	{
		name: 'the registry is per-presentation, so one deck cannot resolve another deck-s style',
		fn: async () => {
			const other = new TsPptx()
			const guid = other.defineTableStyle({ name: 'Elsewhere', firstRow: { fill: '1A2B3C' } })
			// `tableStyles.xml` is written per package, so a GUID from another presentation names
			// a style this one never emits. Resolving it would let a deck defer its cell defaults
			// to a style PowerPoint will not find.
			assertEqual(new TsPptx().addSlide().getCustomTableStyle(guid), undefined, 'no cross-deck leak')
		},
	},
	{
		name: "styleDrivenCells lets a custom style's border render — no per-cell border is written",
		fn: async () => {
			const { slide, styles } = await withStyleParts(
				{ wholeTbl: { border: { type: 'solid', color: 'D9D9D9', width: 0.5 } } },
				{ styleDrivenCells: true }
			)
			// The point of the flag: the cell says nothing about its borders, so PowerPoint resolves
			// them from the style. An `{type:'none'}` here would be direct formatting and win.
			assert(!/<a:ln[LRTB][ />]/.test(allTcPr(slide)[0]), 'no border is written at all; got: ' + allTcPr(slide)[0])
			// And the style itself still carries the border it always did -- the fix is in what the
			// cell stops saying, not in what `tableStyles.xml` says.
			assert(styles.includes('D9D9D9'), 'the style keeps its border; got: ' + styles)
		},
	},
	{
		name: 'a built-in style keeps the per-cell defaults, and the flag says it did nothing',
		fn: async () => {
			// The constraint the whole feature is scoped by: Office's built-ins define borders of
			// their own, so a deck on MEDIUM_STYLE_2_ACCENT_1 must look exactly as it did. A flag
			// that quietly does nothing is the same defect this option exists to fix, hence the code.
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'H1' }]], {
						...AT,
						tableStyle: TableStyle.MEDIUM_STYLE_2_ACCENT_1,
						styleDrivenCells: true,
					})
				})
			)
			assertEqual(codes.join(','), 'table/style-driven-cells-inert', 'one diagnostic, naming the inert flag')
			const tcPr = allTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))[0]
			assertEqual((tcPr.match(/<a:ln[LRTB] w="0"/g) || []).length, 4, 'all four sides are still suppressed')
		},
	},
	{
		name: 'the flag with no table style at all warns and changes nothing',
		fn: async () => {
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable([[{ text: 'H1' }]], { ...AT, styleDrivenCells: true })
				})
			)
			assertEqual(codes.join(','), 'table/style-driven-cells-inert', 'there is nothing to defer to')
			const tcPr = allTcPr(await readEntry(result.zip, 'ppt/slides/slide1.xml'))[0]
			assertEqual((tcPr.match(/<a:ln[LRTB] w="0"/g) || []).length, 4, 'the defaults are untouched')
		},
	},
	{
		name: 'the flag warns once for a table, not once per auto-paged slide',
		fn: async () => {
			// Auto-paging re-enters `addTable` for every overflow slide, so the resolved answer is
			// stashed on the options. The mistake being reported is in the authoring, and it was
			// made once.
			const rows = Array.from({ length: 60 }, (_row, idx) => [{ text: `row ${idx}` }])
			const { result, codes } = await captureDiagnostics(() =>
				build((p) => {
					p.addSlide().addTable(rows, { ...AT, autoPage: true, styleDrivenCells: true })
				})
			)
			// Fails loudly rather than passing vacuously if the table never actually paged.
			await readEntry(result.zip, 'ppt/slides/slide2.xml')
			assertEqual(codes.join(','), 'table/style-driven-cells-inert', 'one warning across the whole table')
		},
	},
	{
		name: 'under the flag, a border the caller set still beats the style',
		fn: async () => {
			// Precedence above the defaults tier is unchanged: cell `options` > table-level >
			// style > nothing. Only the bottom rung stands aside.
			const { slide } = await withStyleParts(
				{ wholeTbl: { border: { type: 'solid', color: 'D9D9D9', width: 0.5 } } },
				{ styleDrivenCells: true },
				[[{ text: 'A', options: { border: { type: 'solid', color: '112233', width: 1 } } }, { text: 'B' }]]
			)
			const [first, second] = allTcPr(slide)
			assertEqual((first.match(/112233/g) || []).length, 4, "the cell's own border is on all four sides")
			assert(!/<a:ln[LRTB][ />]/.test(second), 'the cell next to it still defers; got: ' + second)
		},
	},
	{
		name: 'under the flag, outerBorder draws the perimeter without erasing the style elsewhere',
		fn: async () => {
			// The perimeter is applied at emit time, on a cell that now legitimately arrives with no
			// borders at all. Spelling the other three sides out as `{type:'none'}` to fill the tuple
			// -- which is right for every other table -- would trade the style's borders for it.
			const { slide } = await withStyleParts(
				{ wholeTbl: { border: { type: 'solid', color: 'D9D9D9', width: 0.5 } } },
				{ styleDrivenCells: true, outerBorder: { type: 'solid', color: '1A2B3C', width: 1 } },
				[
					[{ text: 'A' }, { text: 'B' }],
					[{ text: 'C' }, { text: 'D' }],
				]
			)
			const topLeft = allTcPr(slide)[0]
			assert(/<a:lnL /.test(topLeft) && /<a:lnT /.test(topLeft), 'the two perimeter sides draw; got: ' + topLeft)
			assert(!/<a:lnR[ />]/.test(topLeft) && !/<a:lnB[ />]/.test(topLeft), 'the interior sides defer; got: ' + topLeft)
			assertEqual((topLeft.match(/1A2B3C/g) || []).length, 2, 'and both perimeter sides are the outer colour')
		},
	},
])
