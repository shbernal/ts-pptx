import { defineRegressionSuite, build, readEntry, assert, assertEqual, captureDiagnostics } from '../helpers.js'

// The byte-identity harness (`scripts/byte-identity.mjs`) is what gates a behavior-preserving
// refactor of `src/gen/`, but its corpus is only what the showcase decks emit — and no showcase
// authors a pattern fill, an image fill, a glow, or an inner shadow. Migrating `drawingml/fill.ts`
// and `drawingml/effect.ts` onto the `gen/oxml/el.ts` builder therefore had nothing gating it:
// those four emitters produce zero bytes in all 177 baseline parts, so a PASS said nothing at all
// about them.
//
// These pin the exact emitted string for each construct the baseline cannot see. Exact, not
// `includes`: the point of the builder migration is that attribute order, self-closing form, and
// child sequence are unchanged, and every one of those is invisible to a substring check.
//
// If one of these fails after an emitter change, the question is whether the *bytes* were meant
// to change. If they were, that is an OOXML change and needs a schema fixture too (AGENTS.md);
// if they were not, the builder call is wrong.

async function slideXml(buildFn) {
	const { zip } = await build(buildFn)
	return readEntry(zip, 'ppt/slides/slide1.xml')
}

/** A 1x1 transparent PNG — the smallest thing that resolves to a real media relationship. */
const PNG =
	'image/png;base64,' +
	'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

const BOX = { x: 1, y: 1, w: 2, h: 1 }

function assertContainsExactly(xml, expected, label) {
	assert(xml.includes(expected), `expected ${label}:\n  ${expected}\nin:\n${xml}`)
}

defineRegressionSuite('Fill and effect emitters the byte-identity corpus never reaches', [
	{
		name: 'pattern fill emits prst, fgClr, bgClr in that order, with the documented defaults',
		fn: async () => {
			const defaults = await slideXml((pres) => {
				pres.addSlide().addShape('rect', { ...BOX, fill: { type: 'pattern', pattern: { preset: 'diagCross' } } })
			})
			assertContainsExactly(
				defaults,
				'<a:pattFill prst="diagCross"><a:fgClr><a:srgbClr val="000000"/></a:fgClr>' +
					'<a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr></a:pattFill>',
				'pattern fill with default fg/bg'
			)

			// A scheme colour has to survive as `<a:schemeClr>`; an earlier per-site copy of this
			// emitter hardcoded `a:srgbClr` and produced schema-invalid OOXML for scheme tokens.
			const explicit = await slideXml((pres) => {
				pres.addSlide().addShape('rect', {
					...BOX,
					fill: { type: 'pattern', pattern: { preset: 'pct50', fgColor: '003366', bgColor: 'accent2' } },
				})
			})
			assertContainsExactly(
				explicit,
				'<a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="003366"/></a:fgClr>' +
					'<a:bgClr><a:schemeClr val="accent2"/></a:bgClr></a:pattFill>',
				'pattern fill with explicit hex fg and scheme bg'
			)
		},
	},
	{
		name: 'image fill emits a stretched blip; `transparency` adds the only alphaModFix',
		fn: async () => {
			const plain = await slideXml((pres) => {
				pres.addSlide().addShape('rect', { ...BOX, fill: { type: 'image', image: { data: PNG } } })
			})
			// `<a:blip>` stays paired even with nothing to nest — self-closing is decided by
			// arity, not by whether the child list happens to be empty.
			assertContainsExactly(
				plain,
				'<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId1"></a:blip>' +
					'<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
				'image fill without transparency'
			)

			const faded = await slideXml((pres) => {
				pres.addSlide().addShape('rect', { ...BOX, fill: { type: 'image', image: { data: PNG }, transparency: 40 } })
			})
			assertContainsExactly(
				faded,
				'<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId1"><a:alphaModFix amt="60000"/></a:blip>' +
					'<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>',
				'image fill at 40% transparency'
			)
		},
	},
	{
		name: 'an image fill with no resolvable media warns and falls back to <a:noFill/>',
		fn: async () => {
			const { diagnostics, result } = await captureDiagnostics(() =>
				slideXml((pres) => {
					pres.addSlide().addShape('rect', { ...BOX, fill: { type: 'image' } })
				})
			)
			assert(
				diagnostics.some((d) => d.code === 'image-fill/missing-source'),
				`expected an image-fill diagnostic; got: ${diagnostics.map((d) => d.code).join(',')}`
			)
			assert(!result.includes('<a:blipFill'), 'expected no blipFill when the media is unresolved')
		},
	},
	{
		name: 'a radial gradient centres its fillToRect, and `center` shifts it',
		fn: async () => {
			const stops = [
				{ position: 0, color: 'FF0000' },
				{ position: 60, color: 'accent1', transparency: 30 },
				{ position: 100, color: '0000FF' },
			]
			const centred = await slideXml((pres) => {
				pres.addSlide().addShape('rect', { ...BOX, fill: { type: 'gradient', gradient: { kind: 'radial', stops } } })
			})
			assertContainsExactly(
				centred,
				'<a:gradFill rotWithShape="1"><a:gsLst>' +
					'<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
					'<a:gs pos="60000"><a:schemeClr val="accent1"><a:alpha val="70000"/></a:schemeClr></a:gs>' +
					'<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>' +
					'</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>' +
					'</a:gradFill>',
				'radial gradient with a default centre and a transparent stop'
			)

			const shifted = await slideXml((pres) => {
				pres.addSlide().addShape('rect', {
					...BOX,
					fill: {
						type: 'gradient',
						gradient: { kind: 'radial', center: { x: 20, y: 80 }, stops, rotateWithShape: false },
					},
				})
			})
			assertContainsExactly(shifted, '<a:gradFill rotWithShape="0">', 'radial gradient with rotateWithShape off')
			assertContainsExactly(
				shifted,
				'<a:path path="circle"><a:fillToRect l="20000" t="80000" r="80000" b="20000"/></a:path>',
				'radial focus shifted to 20/80'
			)
		},
	},
	{
		name: '`scaled` is present on <a:lin> only when the caller set it',
		fn: async () => {
			const stops = [
				{ position: 0, color: 'FF0000' },
				{ position: 100, color: '0000FF' },
			]
			for (const [scaled, expected] of [
				[undefined, '<a:lin ang="2700000"/>'],
				[true, '<a:lin ang="2700000" scaled="1"/>'],
				[false, '<a:lin ang="2700000" scaled="0"/>'],
			]) {
				const xml = await slideXml((pres) => {
					pres.addSlide().addShape('rect', {
						...BOX,
						fill: { type: 'gradient', gradient: { kind: 'linear', angle: 45, scaled, stops } },
					})
				})
				assertContainsExactly(xml, expected, `linear gradient with scaled=${String(scaled)}`)
			}
		},
	},
	{
		name: 'an outer shadow carries the outerShdw-only attributes; an inner one carries none of them',
		fn: async () => {
			// CT_InnerShadowEffect accepts only blurRad/dist/dir. Emitting sx/sy/kx/ky/algn on it
			// is schema-invalid, and attribute order on the outer one is part of the frozen bytes.
			const inner = await slideXml((pres) => {
				pres.addSlide().addShape('rect', {
					...BOX,
					shadow: { type: 'inner', color: 'C0504D', blur: 4, offset: 2, angle: 90, transparency: 25 },
				})
			})
			assertContainsExactly(
				inner,
				'<a:innerShdw blurRad="50800" dist="25400" dir="5400000">' +
					'<a:srgbClr val="C0504D"><a:alpha val="75000"/></a:srgbClr></a:innerShdw>',
				'inner shadow'
			)

			const outer = await slideXml((pres) => {
				pres.addSlide().addShape('rect', {
					...BOX,
					shadow: { type: 'outer', color: 'accent4', blur: 6, offset: 3, angle: 270, rotateWithShape: true },
				})
			})
			assertContainsExactly(
				outer,
				'<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" rotWithShape="1" ' +
					'blurRad="76200" dist="38100" dir="16200000">' +
					'<a:schemeClr val="accent4"><a:alpha val="75000"/></a:schemeClr></a:outerShdw>',
				'outer shadow'
			)
		},
	},
	{
		name: "shadow `type: 'none'` emits no shadow and no effect list at all",
		fn: async () => {
			// `createShadowElement` returns '' for `none`, and the shape emitter drops the whole
			// `<a:effectLst>` rather than writing the empty one — an explicit "no shadow" leaves
			// the shape with nothing to inherit around, which is the point of asking for it.
			const xml = await slideXml((pres) => {
				pres.addSlide().addShape('rect', { ...BOX, shadow: { type: 'none' } })
			})
			assert(!xml.includes('Shdw'), 'expected no shadow element for type: none')
			assert(!xml.includes('<a:effectLst'), 'expected no effect list for type: none')
		},
	},
	{
		name: 'a text glow emits <a:glow rad=> wrapping the colour and its alpha',
		fn: async () => {
			const xml = await slideXml((pres) => {
				pres.addSlide().addText([{ text: 'glow', options: { glow: { size: 8, color: 'FFFF00', opacity: 0.6 } } }], BOX)
			})
			assertContainsExactly(
				xml,
				'<a:glow rad="101600"><a:srgbClr val="FFFF00"><a:alpha val="60000"/></a:srgbClr></a:glow>',
				'text glow'
			)
		},
	},
	{
		name: 'an 8-char RGBA colour becomes a 6-char value plus exactly one <a:alpha>',
		fn: async () => {
			// CT_SRgbColor allows one alpha child. The RGBA byte must not add a second when the
			// caller already supplied one. `80` is 128/255 of full opacity, which is 50196 rather
			// than a round 50000 — the byte is scaled, not reinterpreted as a percentage.
			const xml = await slideXml((pres) => {
				pres.addSlide().addShape('rect', { ...BOX, fill: { color: '1F386480' } })
			})
			assertContainsExactly(
				xml,
				'<a:solidFill><a:srgbClr val="1F3864"><a:alpha val="50196"/></a:srgbClr></a:solidFill>',
				'RGBA fill'
			)
			assertEqual((xml.match(/<a:alpha /g) || []).length, 1, 'exactly one <a:alpha> in the part')
		},
	},
])
