// Pins the authoring guarantees a read -> write converter depends on, and exactly where
// each one stops.
//
// A converter reads geometry as EMU and colours as raw `a:schemeClr/@val` tokens, then has
// to re-express both through the write API. These tests make the resulting precision and
// vocabulary boundaries executable rather than inferred from type declarations: widening
// `Margin`/`colW` to `Coord`, or extending `SchemeColor` towards the full
// `ST_SchemeColorVal` set, should fail here and be recorded as a deliberate change.
import JSZip from 'jszip'
import TsPptx from '../../../dist/node.js'
import { defineRegressionSuite, assert, assertIncludes } from '../../helpers.js'

const EMU_PER_INCH = 914400

async function slide1Xml(pres) {
	const zip = await JSZip.loadAsync(await pres.toBytes())
	const entry = zip.file('ppt/slides/slide1.xml')
	if (!entry) throw new Error('slide1.xml missing')
	return entry.async('string')
}

function deck() {
	const pres = new TsPptx()
	pres.defineLayout({ name: 'EMU_PROBE', width: 10, height: 5.625 })
	pres.layout = 'EMU_PROBE'
	return pres
}

// Deliberately non-round EMU values: an imprecise conversion anywhere perturbs them, so
// an exact match is proof the value survived untouched.
const X = 914401
const Y = 523241
const W = 2743201
const H = 609601

defineRegressionSuite('EMU-exact geometry and scheme-colour passthrough', [
	{
		name: 'top-level x/y/w/h accept "<n>emu" and reach a:off/a:ext unrounded',
		fn: async () => {
			const pres = deck()
			pres.addSlide().addText('emu', { x: `${X}emu`, y: `${Y}emu`, w: `${W}emu`, h: `${H}emu` })
			const xml = await slide1Xml(pres)
			assertIncludes(xml, `<a:off x="${X}" y="${Y}"/>`, 'slide XML')
			assertIncludes(xml, `<a:ext cx="${W}" cy="${H}"/>`, 'slide XML')
		},
	},
	{
		name: 'custGeom path nodes — including cubic control points — are EMU-exact',
		fn: async () => {
			const pres = deck()
			pres.addSlide().addShape('custGeom', {
				x: `${X}emu`,
				y: `${Y}emu`,
				w: '1000001emu',
				h: '1000001emu',
				points: [
					{ x: '0emu', y: '0emu', moveTo: true },
					{ x: '500001emu', y: '0emu' },
					{
						x: '1000001emu',
						y: '500001emu',
						curve: { type: 'cubic', x1: '700001emu', y1: '100001emu', x2: '900001emu', y2: '300001emu' },
					},
					{ close: true },
				],
			})
			const xml = await slide1Xml(pres)
			assertIncludes(xml, '<a:path w="1000001" h="1000001">', 'slide XML')
			// The three cubicBezTo children: two control points, then the end point.
			for (const pt of ['x="700001" y="100001"', 'x="900001" y="300001"', 'x="1000001" y="500001"']) {
				assertIncludes(xml, pt, 'custGeom path')
			}
		},
	},
	{
		name: 'inches-typed geometry (rowH) is still EMU-exact at full double precision',
		fn: async () => {
			// `colW`/`rowH`/`margin` are typed `number` (inches), not `Coord`, so a converter
			// cannot hand them raw EMU. That is NOT a precision loss: EMU -> inches -> EMU is
			// the identity for every EMU value in a realistic slide, because a double carries
			// far more precision than the ~5.5e7 EMU involved. The loss only appears if the
			// printed decimal is truncated — see the toFixed test below.
			const pres = deck()
			pres.addSlide().addTable([[{ text: 'a' }, { text: 'b' }]], {
				x: 1,
				y: 1,
				w: 8,
				colW: [3, 3],
				rowH: [H / EMU_PER_INCH],
			})
			const xml = await slide1Xml(pres)
			assertIncludes(xml, `<a:tr h="${H}"`, 'table row height')
		},
	},
	{
		name: 'six decimal places is the minimum precision that prints inches EMU-exactly',
		fn: async () => {
			// Fixes the printer's rounding constant by argument plus spot-check rather than taste.
			//
			// Rounding to 6 decimals moves the inch value by at most 5e-7 in, i.e. 5e-7 * 914400
			// = 0.4572 EMU < 0.5 — so `Math.round` always lands back on the original EMU, for
			// every EMU value, unconditionally. At 5 decimals the bound is 4.572 EMU and the
			// round-trip fails for ~89% of values. The assertions below pin both the bound that
			// drives the proof and the observed behaviour at each precision.
			const worstShift = (digits) => {
				let worst = 0
				// Prime step so the sample walks all residues rather than hitting round values.
				for (let emu = 0; emu <= EMU_PER_INCH * 60; emu += 7919) {
					const printed = Number((emu / EMU_PER_INCH).toFixed(digits))
					worst = Math.max(worst, Math.abs(printed * EMU_PER_INCH - emu))
				}
				return worst
			}
			assert(worstShift(6) < 0.5, `6 decimals must shift a value by < 0.5 EMU; measured ${worstShift(6)}`)
			assert(worstShift(5) > 0.5, 'expected 5 decimals to exceed the half-EMU bound; re-derive the constant')

			let bad6 = 0
			let bad5 = 0
			for (let emu = 0; emu <= EMU_PER_INCH * 60; emu += 7919) {
				if (Math.round(Number((emu / EMU_PER_INCH).toFixed(6)) * EMU_PER_INCH) !== emu) bad6++
				if (Math.round(Number((emu / EMU_PER_INCH).toFixed(5)) * EMU_PER_INCH) !== emu) bad5++
			}
			assert(bad6 === 0, `expected 6 decimals to be EMU-exact, got ${bad6} mismatches`)
			assert(bad5 > 0, 'expected 5 decimals to lose EMU precision; if it no longer does, re-derive the constant')
		},
	},
	{
		name: 'margin rejects a raw-EMU string rather than silently mis-scaling it',
		fn: async () => {
			// `marginToEmu` -> `inch2Emu` throws on an unparseable string. Loud is correct, and a
			// converter must divide to inches for margins; this pins that it stays loud.
			//
			// The cast is deliberate: `Margin` is `number | [number, number, number, number]`, so
			// TypeScript already rejects this at compile time — the first line of defence. The
			// runtime guard is what a JS caller (or a converter building options dynamically)
			// actually hits, and that is what this test covers.
			//
			// It throws from `addText`, not from `toBytes`: the insets are resolved when the text
			// object is defined, so the throw names the call that carries the bad value.
			const emuMargin = /** @type {any} */ (['91441emu', 0.1, 0.1, 0.1])
			const pres = deck()
			let threw = null
			try {
				pres.addSlide().addText('inset', { x: 1, y: 1, w: 4, h: 1, margin: emuMargin })
				await pres.toBytes()
			} catch (err) {
				threw = err
			}
			assert(threw, 'expected an EMU-string margin to throw; it was accepted silently')
			assertIncludes(String(threw.message), 'finite number', 'margin error message')
		},
	},
	{
		name: 'the write path accepts exactly the 10 clrMap-mapped scheme tokens',
		fn: async () => {
			const mapped = ['tx1', 'tx2', 'bg1', 'bg2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
			for (const token of mapped) {
				const pres = deck()
				pres.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, fill: { color: token } })
				const xml = await slide1Xml(pres)
				assertIncludes(xml, `<a:schemeClr val="${token}"/>`, `fill for scheme token ${token}`)
			}
		},
	},
	{
		name: 'the 7 unmapped ST_SchemeColorVal tokens degrade to a hex fallback, not schemeClr',
		fn: async () => {
			// ECMA-376 ST_SchemeColorVal has 17 values; `SchemeColor` covers 10. The read path
			// reports `a:schemeClr/@val` verbatim, so the other 7 can reach a converter, and
			// `createColorElement` degrades them to DEF_FONT_COLOR with only a console warning —
			// i.e. it repaints the shape black. A converter must therefore resolve these against
			// the theme's `a:clrScheme` and emit hex, never pass the token through.
			const unmapped = ['dk1', 'lt1', 'dk2', 'lt2', 'hlink', 'folHlink', 'phClr']
			for (const token of unmapped) {
				const pres = deck()
				pres.addSlide().addShape('rect', { x: 1, y: 1, w: 2, h: 1, fill: { color: token } })
				const xml = await slide1Xml(pres)
				assert(
					!xml.includes(`<a:schemeClr val="${token}"`),
					`"${token}" reached the slide as a schemeClr; the write path is not supposed to accept it. ` +
						`If SchemeColor was widened deliberately, a converter may stop flattening this token.`
				)
				assertIncludes(xml, '<a:srgbClr val="000000"/>', `fallback fill for unmapped token ${token}`)
			}
		},
	},
	{
		name: 'a placeholder-bound shape always carries its own explicit a:xfrm',
		fn: async () => {
			// So `placeholder` and geometry are independent, not alternatives: authoring
			// `placeholder: 'title'` never makes the shape inherit layout geometry. Omitting
			// x/y/w/h does not yield inheritance either — it yields a degenerate box (cy="0").
			// A converter must therefore always print concrete geometry, and may print
			// `placeholder` on top of it purely for semantics/accessibility.
			const pres = deck()
			pres.addSlide().addText('Title', { placeholder: 'title' })
			const xml = await slide1Xml(pres)
			assertIncludes(xml, '<p:ph', 'slide XML')
			const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(xml)
			assert(ext, 'expected an explicit <a:ext> on a placeholder-bound shape; got:\n' + xml)
			assert(
				ext[2] === '0',
				`expected a placeholder authored without h to emit a degenerate cy="0" (got cy="${ext[2]}"). ` +
					`If the writer now derives placeholder geometry, a converter could stop printing absolute h.`
			)
		},
	},
	{
		name: 'only 6 of the 16 ST_PlaceholderType values are expressible; the rest emit no type',
		fn: async () => {
			// `PlaceholderType` covers title/body/pic/chart/tbl/media. `genXmlPlaceholder` drops
			// an unrecognised type silently rather than emitting an invalid attribute, so a
			// source `ctrTitle`/`subTitle`/`ftr`/… degrades to an untyped `<p:ph>`. There is also
			// no public setter for `idx` (`_placeholderIdx` is internal), so placeholder identity
			// cannot be fully reproduced. Both are semantic losses, not visual ones — geometry is
			// always explicit — but they must be declared, not discovered.
			for (const token of ['title', 'body', 'pic', 'chart', 'tbl', 'media']) {
				const pres = deck()
				pres.addSlide().addText('x', { placeholder: token, x: 1, y: 1, w: 4, h: 1 })
				const xml = await slide1Xml(pres)
				assertIncludes(xml, `type="${token}"`, `p:ph for expressible type ${token}`)
			}
			for (const token of ['ctrTitle', 'subTitle', 'dt', 'sldNum', 'ftr', 'hdr', 'obj', 'clipArt', 'dgm', 'sldImg']) {
				const pres = deck()
				pres.addSlide().addText('x', { placeholder: token, x: 1, y: 1, w: 4, h: 1 })
				const xml = await slide1Xml(pres)
				assert(
					!xml.includes(`type="${token}"`),
					`"${token}" was emitted as a p:ph type; PlaceholderType is not supposed to cover it. ` +
						`If the vocabulary was widened deliberately, a converter can reproduce more placeholder types.`
				)
			}
		},
	},
])
