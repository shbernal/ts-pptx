import { ShapeType } from '../../../dist/node.js'
import { assert, assertEqual, captureDiagnostics, defineRegressionSuite, slideXml } from '../../helpers.js'

// An empty colour string had four readings depending on which path it fell down. `fill: ''`
// emitted nothing and inherited; `fill: { color: '' }` fell through to `createColorElement` and
// painted DEF_FONT_COLOR, so the object spelling of the same intent came out visible black;
// `line: { color: '' }` resolved to the line default through `||`; and a required slot such as a
// gradient stop painted black with a `color/invalid-value` diagnostic complaining that `""` is
// not a scheme colour. `fillNamesPaint` and `genXmlColorSelection`, two functions in one file,
// disagreed about `{ color: '' }` outright.
//
// One rule now covers all of them: `''` is the caller's own missing value, not a fourth spelling
// of silence next to omission, `'inherit'` and `'none'`. It is reported under
// `color/empty-string` and then resolves to whatever omitting the option resolves to. The
// assertions below are written as *equality against the omitted spelling* rather than against a
// literal fragment, because that equality is the rule; a fragment would go stale the moment a
// default moved.

const BASE = { x: 1, y: 1, w: 2, h: 1 }

/** The slide part for one shape, plus whatever the build reported. */
async function shapeSlide(opts) {
	const { result, codes } = await captureDiagnostics(() =>
		slideXml((p) => {
			p.addSlide().addShape(ShapeType.rect, { ...BASE, ...opts })
		})
	)
	return { xml: result, codes }
}

/** The slide part for one text box, plus whatever the build reported. */
async function textSlide(opts) {
	const { result, codes } = await captureDiagnostics(() =>
		slideXml((p) => {
			p.addSlide().addText('t', { ...BASE, ...opts })
		})
	)
	return { xml: result, codes }
}

/** The slide part for a one-cell table, plus whatever the build reported. */
async function tableSlide(cellOpts, tableOpts) {
	const { result, codes } = await captureDiagnostics(() =>
		slideXml((p) => {
			p.addSlide().addTable([[{ text: 'c', options: cellOpts }]], { ...BASE, ...tableOpts })
		})
	)
	return { xml: result, codes }
}

/** Assert that `empty` produced the same part as `omitted`, and said so on the way. */
function assertSameAsOmitted(empty, omitted, label) {
	assertEqual(empty.xml, omitted.xml, `${label}: an empty colour must emit what omitting the option emits`)
	assert(
		empty.codes.includes('color/empty-string'),
		`${label}: an empty colour must be reported; got ${JSON.stringify(empty.codes)}`
	)
	assertEqual(omitted.codes.length, 0, `${label}: omitting the option must report nothing`)
}

defineRegressionSuite('An empty colour string is a missing value, not a fourth spelling of silence', [
	{
		name: "a shape's fill: '' emits what omitting fill emits",
		fn: async () => assertSameAsOmitted(await shapeSlide({ fill: '' }), await shapeSlide({}), 'shape fill'),
	},
	{
		name: "a shape's fill: { color: '' } agrees with fill: '' instead of painting black",
		fn: async () => {
			const object = await shapeSlide({ fill: { color: '' } })
			const bare = await shapeSlide({ fill: '' })
			assertEqual(object.xml, bare.xml, "fill: { color: '' } must agree with fill: ''")
			assert(
				!object.xml.includes('<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'),
				`the object spelling still painted DEF_FONT_COLOR: ${object.xml}`
			)
			assert(
				object.codes.includes('color/empty-string'),
				`the object spelling must be reported; got ${JSON.stringify(object.codes)}`
			)
		},
	},
	{
		name: "a shape's line: { color: '' } takes the line default an omitted colour takes",
		fn: async () =>
			assertSameAsOmitted(
				await shapeSlide({ line: { color: '', width: 2 } }),
				await shapeSlide({ line: { width: 2 } }),
				'line.color'
			),
	},
	{
		name: "a text run's color: '' emits what omitting color emits",
		fn: async () => assertSameAsOmitted(await textSlide({ color: '' }), await textSlide({}), 'text color'),
	},
	{
		name: "a text box's fill: '' still emits the <a:noFill/> an omitted fill emits",
		fn: async () => {
			const empty = await textSlide({ fill: '' })
			assertSameAsOmitted(empty, await textSlide({}), 'text box fill')
			assert(empty.xml.includes('<a:noFill/>'), `a text box must stay unfilled: ${empty.xml}`)
		},
	},
	{
		name: "a table cell's fill: '' emits what omitting fill emits",
		fn: async () => assertSameAsOmitted(await tableSlide({ fill: '' }), await tableSlide({}), 'cell fill'),
	},
	{
		name: "a table's tableFill: '' emits what omitting tableFill emits",
		fn: async () => assertSameAsOmitted(await tableSlide({}, { tableFill: '' }), await tableSlide({}, {}), 'tableFill'),
	},
	{
		name: "a shadow's color: '' takes the shadow default an omitted colour takes",
		fn: async () =>
			assertSameAsOmitted(
				await shapeSlide({ fill: 'FF0000', shadow: { type: 'outer', color: '' } }),
				await shapeSlide({ fill: 'FF0000', shadow: { type: 'outer' } }),
				'shadow.color'
			),
	},
	{
		name: 'the diagnostic names the option that carried the empty string',
		fn: async () => {
			// The message is not API and is free to change; that it names the *option* is the
			// point, because the whole value of reporting this is telling the caller which of
			// their own keys was blank. Naming the emitter that saw it last would not.
			const { messages } = await captureDiagnostics(() =>
				slideXml((p) => {
					p.addSlide().addShape(ShapeType.rect, { ...BASE, line: { color: '' } })
				})
			)
			assert(
				messages.some((m) => m.includes('`line.color`')),
				`no message named the option; got ${JSON.stringify(messages)}`
			)
		},
	},
	{
		name: 'a required colour slot has no absent state to fall back to, so it paints and says why',
		fn: async () => {
			// A gradient stop is the counter-case: there is no "inherit" to resolve to, so `''`
			// still paints DEF_FONT_COLOR. It is reported under the same code rather than under
			// `color/invalid-value`, whose message complained that `""` is not a scheme colour.
			const { result, codes } = await captureDiagnostics(() =>
				slideXml((p) => {
					p.addSlide().addShape(ShapeType.rect, {
						...BASE,
						fill: {
							gradient: {
								kind: 'linear',
								angle: 90,
								stops: [
									{ position: 0, color: '' },
									{ position: 100, color: 'FF0000' },
								],
							},
						},
					})
				})
			)
			assert(result.includes('<a:srgbClr val="000000"/>'), `the stop must still paint: ${result}`)
			assert(
				codes.includes('color/empty-string'),
				`the stop must report the empty string; got ${JSON.stringify(codes)}`
			)
			assert(
				!codes.includes('color/invalid-value'),
				`an empty string is not a malformed colour; got ${JSON.stringify(codes)}`
			)
		},
	},
])
