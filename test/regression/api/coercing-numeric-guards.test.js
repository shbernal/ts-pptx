import { defineRegressionSuite, build, readEntry, assert, assertEqual, captureDiagnostics } from '../../helpers.js'

// The companion to `numeric-conversion-guards.test.js`: that file pins what the *converters*
// refuse, this one pins what the *guards in front of them* let through. `src/` used three tests
// for "is this a usable number" and they disagreed. The global `isNaN` and `isFinite` coerce
// their argument, so `Number('') === 0` makes `isNaN('')` false and an empty string reads as a
// valid zero, and `isNaN(Infinity)` is false, so a value with no finite representation passes a
// guard whose whole job was to stop one. `.oxlintrc.jsonc` bans both under
// `no-restricted-globals` so the sweep that removed them cannot decay; these are the behaviours
// the sweep changed, which a lint rule cannot pin.

const BOX = { x: 1, y: 1, w: 4, h: 2 }
const EMU_PER_INCH = 914400
const SERIES = [{ name: 'S', labels: ['A', 'B'], values: [1, 2] }]

async function slideXml(buildFn) {
	const { zip } = await build(buildFn)
	return readEntry(zip, 'ppt/slides/slide1.xml')
}

defineRegressionSuite('Coercing numeric guards', [
	{
		// The one with the largest blast radius. `options.x` is a `Coord`, so `'50%'` and `'2in'`
		// are ordinary in-type values — and `!isNaN(Number('50%'))` is false for both, so the
		// guard discarded them and substituted the 1in default. A chart positioned as a
		// percentage of the slide silently landed an inch from the left edge.
		name: 'a chart `x`/`y` spelled as a percentage or a unit string is honoured, not replaced by the default',
		fn: async () => {
			const xml = await slideXml((p) => p.addSlide().addChart(SERIES, { type: 'bar', x: '50%', y: '2in', w: 4, h: 3 }))
			const frame = xml.slice(xml.indexOf('<p:graphicFrame>'))
			const off = /<a:off x="(-?\d+)" y="(-?\d+)"\/>/.exec(frame)
			assert(off, 'expected an <a:off> on the chart frame; got:\n' + frame.slice(0, 400))
			// 50% of the default 10in-wide layout is 5in; '2in' is 2in. Both were 1in before.
			assertEqual(Number(off[1]), 5 * EMU_PER_INCH, "chart x: '50%' resolves against the slide width")
			assertEqual(Number(off[2]), 2 * EMU_PER_INCH, "chart y: '2in' resolves as inches")
		},
	},
	{
		// The same guard's other half: a `NaN` position used to read as valid and be replaced by
		// the default without a word. It now reaches the coordinate converter, which refuses it.
		name: 'a NaN chart position is refused rather than silently defaulted',
		fn: async () => {
			let code = null
			try {
				await build((p) => p.addSlide().addChart(SERIES, { type: 'bar', x: NaN, y: 1, w: 4, h: 3 }))
			} catch (err) {
				code = err?.code ?? null
			}
			assertEqual(code, 'coord/non-finite', 'a NaN chart x is reported')
		},
	},
	{
		// `lvl` is written straight into the attribute with no converter in between, so the
		// truthiness-plus-`> 0` guard put the literal string `Infinity` in the package.
		// `ST_TextIndentLevelType` is 0-8, so the check is a range, not just a finiteness test.
		name: 'an out-of-range indentLevel is reported and omitted instead of reaching the attribute',
		fn: async () => {
			for (const indentLevel of [Infinity, 9, 2.5, -1]) {
				const { result: xml, codes } = await captureDiagnostics(() =>
					slideXml((p) => p.addSlide().addText('x', { ...BOX, indentLevel }))
				)
				assert(!/<a:pPr[^>]*\blvl=/.test(xml), `indentLevel ${indentLevel} should emit no lvl; got:\n${xml}`)
				// -1 is falsy-adjacent but not falsy, so it must be reported like the rest.
				assert(codes.includes('text/invalid-indent-level'), `indentLevel ${indentLevel} should be reported`)
			}
			// The whole valid range still emits.
			const ok = await slideXml((p) => p.addSlide().addText('x', { ...BOX, indentLevel: 8 }))
			assert(ok.includes('lvl="8"'), `indentLevel 8 should emit lvl="8"; got:\n${ok}`)
		},
	},
	{
		// A header-row count is a count. `!isNaN(Number(x))` is true for `''` (which reads as 0
		// header rows), for a fraction, and for `Infinity`.
		name: 'autoPageHeaderRows must be a whole number no larger than the table',
		fn: async () => {
			for (const autoPageHeaderRows of [Infinity, 0, 1.5, -2, 9]) {
				const { codes } = await captureDiagnostics(() =>
					build((p) =>
						p.addSlide().addTable([[{ text: 'a' }], [{ text: 'b' }]], {
							...BOX,
							autoPage: true,
							autoPageRepeatHeader: true,
							autoPageHeaderRows,
						})
					)
				)
				assert(
					codes.includes('table/invalid-header-row-count'),
					`autoPageHeaderRows ${autoPageHeaderRows} should be reported; got: ${codes.join(',')}`
				)
			}
			const { codes } = await captureDiagnostics(() =>
				build((p) =>
					p.addSlide().addTable([[{ text: 'a' }], [{ text: 'b' }]], {
						...BOX,
						autoPage: true,
						autoPageRepeatHeader: true,
						autoPageHeaderRows: 2,
					})
				)
			)
			assert(!codes.includes('table/invalid-header-row-count'), 'a count equal to the row count is fine')
		},
	},
	{
		// Chart border widths reach `a:ln/@w`, so they go through `lineWidthToEmu` now — the
		// clamp that already existed for shape strokes and that these paths were not calling. A
		// negative width used to be emitted as a negative attribute.
		name: 'a chart border width outside ST_LineWidth is clamped rather than emitted',
		fn: async () => {
			const { result: xml, codes } = await captureDiagnostics(() =>
				slideXml((p) =>
					p.addSlide().addChart(SERIES, {
						type: 'bar',
						...BOX,
						plotArea: { border: { type: 'solid', color: '0088CC', width: -5 } },
					})
				)
			)
			assert(!/w="-\d/.test(xml), `no negative line width should reach the package; got:\n${xml}`)
			// A negative is not a width, so the define pass substitutes the documented default
			// before the emitter ever sees it, and nothing has to be clamped.
			assert(!codes.includes('line/width-out-of-range'), `expected the default, not a clamp; got: ${codes}`)
		},
	},
])
