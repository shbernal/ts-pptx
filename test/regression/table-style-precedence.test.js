import { TsPptx, defineRegressionSuite, build, readEntry, captureDiagnostics, assert, assertEqual } from '../helpers.js'

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

/** Build a one-cell table under a custom style and return its slide part. */
async function withStyle(styleDef, tableOpts = {}) {
	const { zip } = await build((p) => {
		const guid = p.defineTableStyle({ name: 'Brand', ...styleDef })
		p.addSlide().addTable([[{ text: 'H1' }]], { ...AT, tableStyle: guid, hasHeader: true, ...tableOpts })
	})
	return readEntry(zip, 'ppt/slides/slide1.xml')
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
])
