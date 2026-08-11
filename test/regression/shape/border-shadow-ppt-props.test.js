import JSZip from 'jszip'
import TsPptx from '../../../dist/node.js'
import { defineRegressionSuite, assert } from '../../helpers.js'

async function buildSlide1(pres) {
	const buf = await pres.stream()
	const zip = await JSZip.loadAsync(buf)
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

// PowerPoint-aligned style props: BorderProps.width (points) and BorderProps.transparency
// (0-100), plus ShadowProps.transparency (0-100, the value the PPT UI actually shows) —
// the removed 0.0-1.0 `opacity` had the same role before it was dropped from the public API.
defineRegressionSuite('PPT-aligned border/shadow props [border-shadow-ppt-props]', [
	{
		name: 'table cell border `width` emits the correct line w=',
		fn: async () => {
			const pres = new TsPptx()
			pres.addSlide().addTable([[{ text: 'x' }]], {
				x: 1,
				y: 1,
				border: { type: 'solid', width: 2, color: 'FF0000' },
			})
			const xml = await buildSlide1(pres)
			// 2pt -> 25400 EMU
			assert(xml.includes('<a:lnL w="25400"'), 'expected `width:2` to emit w="25400"; got:\n' + xml)
		},
	},
	{
		name: 'table cell border `transparency` emits <a:alpha> inside the line fill',
		fn: async () => {
			const pres = new TsPptx()
			pres.addSlide().addTable([[{ text: 'x' }]], {
				x: 1,
				y: 1,
				border: { type: 'solid', width: 2, color: 'FF0000', transparency: 25 },
			})
			const xml = await buildSlide1(pres)
			const i = xml.indexOf('<a:lnL ')
			const line = xml.substring(i, xml.indexOf('</a:lnL>', i))
			// transparency 25 -> alpha 75000
			assert(line.includes('<a:alpha val="75000"/>'), 'expected transparency:25 -> alpha 75000; got:\n' + line)
		},
	},
	{
		name: 'shadow `transparency` (0-100) sets the emitted alpha',
		fn: async () => {
			const presT = new TsPptx()
			presT.addSlide().addShape('rect', {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				shadow: { type: 'outer', color: '000000', blur: 3, offset: 2, transparency: 25 },
			})
			const xmlT = await buildSlide1(presT)
			// transparency 25 -> alpha 75000
			assert(xmlT.includes('<a:alpha val="75000"/>'), 'expected transparency:25 -> alpha 75000; got:\n' + xmlT)
		},
	},
	{
		name: 'removed `opacity` input is ignored — falls back to the shape/text-shadow default alpha',
		fn: async () => {
			const presO = new TsPptx()
			presO.addSlide().addShape('rect', {
				x: 1,
				y: 1,
				w: 2,
				h: 1,
				// `opacity` is no longer a public ShadowProps field; an untyped caller still
				// passing it must be ignored, not silently honored via the internal field reuse.
				// A value distinct from the DEF_TEXT_SHADOW default (0.75) so a still-honored
				// opacity would produce a visibly different (and wrong) alpha.
				// @ts-expect-error verifying a legacy/untyped caller's `opacity` is ignored, not honored
				shadow: { type: 'outer', color: '000000', blur: 3, offset: 2, opacity: 0.2 },
			})
			const xmlO = await buildSlide1(presO)
			// addShape's shadow defaults come from DEF_TEXT_SHADOW (opacity 0.75) -> alpha 75000,
			// NOT the ignored opacity:0.2 (which would be alpha 20000 if it were still honored).
			assert(
				xmlO.includes('<a:alpha val="75000"/>'),
				'expected removed opacity:0.2 to be ignored, falling back to the 0.75 text-shadow default; got:\n' + xmlO
			)
		},
	},
	// BorderProps.cap reached the emitter (`genTableCellBorderXml` calls `createLineCap`)
	// but never a value: both places that normalize a border into a complete 4-side tuple
	// rebuilt each side field-by-field and had no `cap` key, so every table border emitted
	// cap="flat" no matter what the caller asked for. One case per normalization path.
	{
		name: 'table-level border `cap` survives normalization (single BorderProps, applied to all four sides)',
		fn: async () => {
			const pres = new TsPptx()
			pres.addSlide().addTable([[{ text: 'x' }]], {
				x: 1,
				y: 1,
				border: { type: 'solid', width: 2, color: 'FF0000', cap: 'round' },
			})
			const xml = await buildSlide1(pres)
			for (const side of ['lnL', 'lnR', 'lnT', 'lnB']) {
				assert(
					xml.includes(`<a:${side} w="25400" cap="rnd"`),
					`expected cap:'round' -> cap="rnd" on <a:${side}>; got:\n` + xml
				)
			}
		},
	},
	{
		name: 'per-side border `cap` survives the tuple normalization, per side',
		fn: async () => {
			const pres = new TsPptx()
			pres.addSlide().addTable([[{ text: 'x' }]], {
				x: 1,
				y: 1,
				// [top, right, bottom, left] — one cap per side, including an omitted one
				// so the emitter's own 'flat' default is pinned alongside the carried values.
				border: [
					{ type: 'solid', width: 2, color: '000000', cap: 'round' },
					{ type: 'solid', width: 2, color: '000000', cap: 'square' },
					{ type: 'solid', width: 2, color: '000000', cap: 'flat' },
					{ type: 'solid', width: 2, color: '000000' },
				],
			})
			const xml = await buildSlide1(pres)
			const expected = { lnT: 'rnd', lnR: 'sq', lnB: 'flat', lnL: 'flat' }
			for (const [side, cap] of Object.entries(expected)) {
				assert(
					xml.includes(`<a:${side} w="25400" cap="${cap}"`),
					`expected <a:${side}> to emit cap="${cap}"; got:\n` + xml
				)
			}
		},
	},
	{
		name: 'a `type: none` border carries its `cap` too',
		fn: async () => {
			const pres = new TsPptx()
			pres.addSlide().addTable([[{ text: 'x' }]], {
				x: 1,
				y: 1,
				border: { type: 'none', cap: 'square' },
			})
			const xml = await buildSlide1(pres)
			// The `none` arm emits w="0" + <a:noFill/>, but still writes the cap attribute.
			assert(xml.includes('<a:lnL w="0" cap="sq"'), 'expected a none-type border to keep cap="sq"; got:\n' + xml)
		},
	},
])
