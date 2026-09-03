import { assert, assertEqual, build, defineRegressionSuite, readEntry, slideXml } from '../../helpers.js'

// A bare colour is the solid-fill shorthand: `fill: 'FF0000'` says exactly what
// `fill: { color: 'FF0000' }` says, at every fill option in the API. The runtime accepted it at
// five of the six and the types accepted it at none, so the spelling worked without being
// promised — and the sixth, the chart area, did not work at all: `copyChartOptions` spread the
// string into `{0:'F',1:'F',…}`, which named neither `color` nor `type`, so `isStatedFill`
// rejected it and the area came out `<a:noFill/>`.
//
// Both halves are settled now — `FillOption` is `Color | ShapeFillProps` and the chart clone
// leaves a string alone — so these pin the promise: every site takes the shorthand, and at every
// site it emits the same bytes the object form does.
//
// A stroke is deliberately not one of them. `line` carries width and dash alongside its paint,
// and those defaults come from rebuilding the line object at definition time; the last case here
// pins that a stroke is still spelled as an object.

const BOX = { x: 1, y: 1, w: 2, h: 1 }
const RED = 'FF0000'

/** The `<a:srgbClr>` values inside `block`, in document order. */
function colors(block) {
	return [...block.matchAll(/<a:srgbClr val="([0-9A-F]{6})"/g)].map((m) => m[1])
}

/** Slide 1 of a deck built by `buildFn`. */
/**
 * Build the same deck twice — once with `fill: 'FF0000'`, once with `fill: { color: 'FF0000' }`
 * — and assert both the shorthand painted and the two agree byte for byte.
 * @param part the package part to compare, e.g. `ppt/slides/slide1.xml`
 * @param buildFn takes the fill to use and returns a deck builder
 */
async function agreesWithObjectForm(part, buildFn) {
	const { zip: shortZip } = await build(buildFn(RED))
	const { zip: longZip } = await build(buildFn({ color: RED }))
	const short = await readEntry(shortZip, part)
	const long = await readEntry(longZip, part)
	assert(
		short.indexOf(`<a:solidFill><a:srgbClr val="${RED}"/></a:solidFill>`) !== -1,
		'shorthand painted nothing in:\n' + short
	)
	assertEqual(short, long, `the shorthand and { color } disagree in ${part}`)
}

defineRegressionSuite('A bare colour is the solid-fill shorthand [fill-option]', [
	{
		name: 'a shape interior takes it',
		fn: async () => {
			await agreesWithObjectForm('ppt/slides/slide1.xml', (fill) => (pres) => {
				pres.addSlide().addShape('rect', { ...BOX, fill })
			})
		},
	},
	{
		name: 'a text box interior takes it',
		fn: async () => {
			await agreesWithObjectForm('ppt/slides/slide1.xml', (fill) => (pres) => {
				pres.addSlide().addText('hi', { ...BOX, fill })
			})
		},
	},
	{
		name: 'a table cell takes it',
		fn: async () => {
			await agreesWithObjectForm('ppt/slides/slide1.xml', (fill) => (pres) => {
				pres.addSlide().addTable([[{ text: 'a', options: { fill } }]], { ...BOX })
			})
		},
	},
	{
		name: "a table's per-cell default and its own background both take it",
		fn: async () => {
			await agreesWithObjectForm('ppt/slides/slide1.xml', (fill) => (pres) => {
				pres.addSlide().addTable([[{ text: 'a' }]], { ...BOX, fill })
			})
			await agreesWithObjectForm('ppt/slides/slide1.xml', (fill) => (pres) => {
				pres.addSlide().addTable([[{ text: 'a' }]], { ...BOX, tableFill: fill })
			})
		},
	},
	{
		name: 'a slide background takes it',
		fn: async () => {
			await agreesWithObjectForm('ppt/slides/slide1.xml', (fill) => (pres) => {
				pres.addSlide().background = fill
			})
		},
	},
	{
		name: 'a chart plot area and chart area take it, where the string used to vanish',
		fn: async () => {
			// The regression this file exists for. Before `copyChartOptions` stopped spreading the
			// string, both of these emitted `<a:noFill/>` and the requested colour appeared nowhere
			// in the part.
			await agreesWithObjectForm('ppt/charts/chart1.xml', (fill) => (pres) => {
				pres.addSlide().addChart([{ name: 'S', labels: ['a'], values: [1] }], {
					type: 'bar',
					x: 1,
					y: 1,
					w: 4,
					h: 3,
					plotArea: { fill },
				})
			})
			await agreesWithObjectForm('ppt/charts/chart1.xml', (fill) => (pres) => {
				pres.addSlide().addChart([{ name: 'S', labels: ['a'], values: [1] }], {
					type: 'bar',
					x: 1,
					y: 1,
					w: 4,
					h: 3,
					chartArea: { fill },
				})
			})
		},
	},
	{
		name: "the caller's own fill object is not mutated by the chart path",
		fn: async () => {
			// `copyChartOptions` copies `plotArea.fill` so normalization cannot reach back into the
			// caller's object. The `typeof === 'object'` guard must not have cost that.
			const fill = { color: RED }
			await build((pres) => {
				pres.addSlide().addChart([{ name: 'S', labels: ['a'], values: [1] }], {
					type: 'bar',
					x: 1,
					y: 1,
					w: 4,
					h: 3,
					plotArea: { fill },
				})
			})
			assertEqual(Object.keys(fill).join(','), 'color', "the chart path wrote back into the caller's fill")
			assertEqual(fill.color, RED, "the chart path changed the caller's colour")
		},
	},
	{
		name: 'a stroke is still spelled as an object',
		fn: async () => {
			// Not an oversight: a bare string would paint the colour and skip the line rebuild that
			// supplies `w` and `prstDash`, so it would silently lose the stroke's defaults. The type
			// refuses it; this pins what the object form emits so the asymmetry stays visible.
			const xml = await slideXml((pres) => {
				pres.addSlide().addShape('rect', { ...BOX, line: { color: RED } })
			})
			const ln = xml.slice(xml.indexOf('<a:ln'), xml.indexOf('</a:ln>'))
			assertEqual(colors(ln).join(','), RED, 'expected the stroke colour in:\n' + ln)
			assert(ln.indexOf('w="12700"') !== -1, 'expected the default 1pt width in:\n' + ln)
			assert(ln.indexOf('<a:prstDash val="solid"/>') !== -1, 'expected the default dash in:\n' + ln)
		},
	},
])
