import { assert, build, defineRegressionSuite, firstXmlBlock, slideXml } from '../../helpers.js'

// A stroke is painted like a fill: `ShapeLineProps extends ShapeFillProps`, so `line` accepts
// `gradient`/`pattern`/`image` as well as a solid `color`, plus its own `cap`. The emitters
// (`drawingml/line.ts`, `slide/object.ts`) read all of those — but the define pass rebuilt the
// caller's `line` object from a fixed key list, so a key added to the type without also being
// added to that list was dropped before the emitter ever saw it. `cap` was silently ignored on
// every shape and table border; `pattern` was worse than ignored, reaching `genXmlPatternFill`
// with no pattern object and throwing "Pattern fill requires a pattern object."
//
// Both rebuilds now spread the caller's object and override only what they default. These pin
// the paths that were dropping keys — one per rebuild site (define/shape.ts, define/text.ts).
// Table borders are pinned in border-shadow-ppt-props.test.js, and the schema fixture
// "shape with pattern line" checks a `<a:pattFill>` stroke against the validator.

/** The `<a:ln>` element of the part's first shape. */
function lineBlock(xml) {
	const ln = firstXmlBlock(xml, 'a:ln')
	assert(ln, 'expected an <a:ln> in:\n' + xml)
	return ln
}

async function expectBuildError(buildFn, expectedMessage) {
	let err
	try {
		await build(buildFn)
	} catch (e) {
		err = e
	}
	assert(err, 'expected build to fail')
	const message = String(err?.message || err)
	assert(message.includes(expectedMessage), `expected error to include "${expectedMessage}"; got: ${message}`)
}

defineRegressionSuite('Shape line paint and cap', [
	{
		name: "addShape line `cap` reaches the <a:ln cap=> attribute ('round' -> rnd, 'square' -> sq)",
		fn: async () => {
			for (const [cap, expected] of [
				['round', 'rnd'],
				['square', 'sq'],
				['flat', 'flat'],
			]) {
				const xml = await slideXml((p) => {
					p.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, line: { color: '0070C0', width: 3, cap } })
				})
				assert(
					lineBlock(xml).includes(`cap="${expected}"`),
					`expected cap:'${cap}' -> cap="${expected}"; got: ${lineBlock(xml)}`
				)
			}
		},
	},
	{
		name: 'an omitted `cap` emits no cap attribute at all — the stroke inherits the default',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, line: { color: '0070C0', width: 3 } })
			})
			// `slide/object.ts` writes the attribute only when `cap` is set, so the absence of a
			// caller value must stay an absent attribute rather than an explicit cap="flat".
			assert(!/<a:ln[^>]*\bcap=/.test(lineBlock(xml)), `expected no cap attribute; got: ${lineBlock(xml)}`)
		},
	},
	{
		name: 'addShape line `type: pattern` emits <a:pattFill> as the stroke paint',
		fn: async () => {
			const xml = await slideXml((p) => {
				p.addSlide().addShape('rect', {
					x: 1,
					y: 1,
					w: 4,
					h: 1,
					line: {
						type: 'pattern',
						width: 3,
						pattern: { preset: 'diagCross', fgColor: '003366', bgColor: 'FFFFFF' },
					},
				})
			})
			const ln = lineBlock(xml)
			assert(ln.includes('<a:pattFill prst="diagCross">'), `expected a pattern stroke; got: ${ln}`)
			assert(ln.includes('<a:fgClr><a:srgbClr val="003366"/></a:fgClr>'), `expected the fg color; got: ${ln}`)
			assert(ln.includes('<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr>'), `expected the bg color; got: ${ln}`)
		},
	},
	{
		name: 'the addText `shape: line` path carries `cap` and a gradient stroke through its own rebuild',
		fn: async () => {
			// define/text.ts has a second, near-duplicate ShapeLineProps rebuild for text objects
			// whose `shape` is a line. It dropped `gradient` as well as `cap`, so a gradient
			// stroke on this path degraded to the default solid line color.
			const xml = await slideXml((p) => {
				p.addSlide().addText('x', {
					shape: 'line',
					x: 1,
					y: 1,
					w: 4,
					h: 0,
					line: {
						cap: 'round',
						width: 2,
						gradient: {
							kind: 'linear',
							angle: 0,
							stops: [
								{ position: 0, color: '003366' },
								{ position: 100, color: 'FFFFFF' },
							],
						},
					},
				})
			})
			const ln = lineBlock(xml)
			assert(ln.includes('cap="rnd"'), `expected cap="rnd" on the text-line stroke; got: ${ln}`)
			assert(ln.includes('<a:gradFill'), `expected a gradient stroke; got: ${ln}`)
		},
	},
	{
		name: 'an unrecognized `cap` from an untyped caller fails the build rather than emitting it',
		fn: async () => {
			// `createLineCap`'s exhaustiveness arm. TypeScript rules this out, but JS callers are
			// real and `cap` is written straight into an attribute — an unknown value would reach
			// the package as cap="INVALID" and PowerPoint would offer to repair the file. Note the
			// chart path differs deliberately: `define/chart.ts` scrubs an unrecognized gridLine
			// cap before emit (chart-option-validation.test.js), so only the shape path throws.
			await expectBuildError((p) => {
				p.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, line: { color: '0070C0', cap: 'INVALID' } })
			}, 'Invalid line cap')
		},
	},
])
