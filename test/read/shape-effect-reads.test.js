// Read-model coverage for the shape EFFECT / STROKE-DETAIL / PATTERN-FILL getters
// in src/read/api/shapes/base.ts that the fixture decks don't happen to carry:
// Shape.shadow (a:effectLst/a:outerShdw), Shape.innerShadow (a:innerShdw),
// Shape.glow (a:glow), Shape.reflection (a:reflection), Shape.softEdge
// (a:softEdge), Shape.patternFill (a:pattFill), Shape.lineEnds (a:ln head/tail
// arrowheads), and the a:path branch of Shape.gradientFill.
//
// These read only the shape's own spPr (colour-bearing effects resolve through
// the slide theme, so a minimal themeContext stub is enough), so hand-authored
// OOXML exercises every branch without a round-trip through this library's
// writer — the same off-fixture pattern the style-accessor suite uses.
//
// The effects the writer DOES author (inner shadow via `shadow: { type: 'inner' }`
// and the pattern fill) additionally get a write→read fidelity leg through the
// shared harness, proving those reads round-trip the very bytes the writer emits.

import { ShapeType } from '../../dist/node.js'
import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { AutoShape } from '../../dist/read.js'
import { authorRead, firstShape, schemaErrors, validatorInstalled } from './authored.js'
import { assert, assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/** Empty colour maps resolve `a:srgbClr` literally; no theme parts beyond that. */
function ctx() {
	return {
		clrMap: new Map(),
		clrScheme: new Map(),
		fmtScheme: null,
		fontScheme: null,
		layoutRoot: null,
		masterRoot: null,
	}
}

/** An `AutoShape` over a hand-authored `p:sp` body, resolving against a theme stub. */
function sp(spPrInner) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:sp>${spPrInner}</p:sp></p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'sp')[0]
	return new AutoShape(el, /** @type {any} */ ({ themeContext: () => ctx() }))
}

describe('Shape.shadow — outer drop shadow reads', () => {
	test('reads blur / offset / angle (EMU + 60000ths) and an srgb colour with alpha', () => {
		const shape = sp(
			`<p:spPr><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000">` +
				`<a:srgbClr val="808080"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst></p:spPr>`
		)
		const shadow = shape.shadow
		assert(shadow, 'an outerShdw surfaces a shadow')
		assertEqual(shadow.color, '808080', 'effectiveHex of the plain srgb shadow colour')
		assert(Math.abs(shadow.alpha - 0.4) < 1e-9, `alpha 40000 → 0.4, got ${shadow.alpha}`)
		assertEqual(shadow.blurPt, 4, 'blurRad 50800 EMU → 4pt')
		assertEqual(shadow.offsetPt, 3, 'dist 38100 EMU → 3pt')
		assertEqual(shadow.angleDeg, 45, 'dir 2700000 (60000ths) → 45°')
		assertEqual(shadow.colorToken, undefined, 'an srgb shadow carries no scheme colour token')
	})

	test('a scheme-coloured shadow surfaces its colorToken even when unresolved', () => {
		const shadow = sp(
			`<p:spPr><a:effectLst><a:outerShdw><a:schemeClr val="accent1"/></a:outerShdw></a:effectLst></p:spPr>`
		).shadow
		assert(shadow, 'the outerShdw still surfaces a shadow')
		assertEqual(shadow.colorToken, 'accent1', 'the scheme token is reported for a downstream resolver')
		assertEqual(shadow.color, null, 'with empty colour maps the scheme colour does not resolve')
	})

	test('the colour is read whatever model it uses, not only srgb and scheme', () => {
		// `a:CT_OuterShadowEffect` is a sequence whose only child is the required, single-member
		// `a:EG_ColorChoice` group — one of `a:scrgbClr | a:srgbClr | a:hslClr | a:sysClr |
		// a:schemeClr | a:prstClr`. So "the shadow's first element child" is the colour, always,
		// which is what `glow` has always done. Naming two of the six explicitly read the other
		// four as `color: null`, and `a:sysClr` is not hypothetical: `resolveColor` resolves it
		// everywhere else in the read model, through its `lastClr` snapshot.
		const sys = sp(
			`<p:spPr><a:effectLst><a:outerShdw blurRad="50800" dist="38100">` +
				`<a:sysClr val="windowText" lastClr="000000"/></a:outerShdw></a:effectLst></p:spPr>`
		).shadow
		assert(sys, 'a sysClr outerShdw surfaces a shadow')
		assertEqual(sys.color, '000000', 'the sysClr lastClr snapshot resolves like it does elsewhere')
		assertEqual(sys.offsetPt, 3, 'and the geometry is decoded alongside it')
		// `a:prstClr` is the model this library emits itself (`gen/slide/notes.ts`). `resolveColor`
		// has no preset-name table yet, so the colour still reads `null` — but the element now
		// reaches the resolver, which is where that gap belongs, rather than being dropped here.
		const prst = sp(
			`<p:spPr><a:effectLst><a:outerShdw blurRad="50800"><a:prstClr val="black"/></a:outerShdw></a:effectLst></p:spPr>`
		).shadow
		assert(prst, 'a prstClr outerShdw surfaces a shadow')
		assertEqual(prst.blurPt, 4, 'with its geometry decoded')
		assertEqual(prst.color, null, 'the preset name is not resolvable yet — see [prstclr-resolution]')
	})

	test('no effectLst / no outerShdw → null', () => {
		assertEqual(sp(`<p:spPr/>`).shadow, null, 'no a:effectLst → null')
		assertEqual(sp(`<p:spPr><a:effectLst/></p:spPr>`).shadow, null, 'an effectLst with no outerShdw → null')
	})
})

describe('Shape.lineEnds — connector arrowheads', () => {
	test('reads both head and tail ends with their type / width / length', () => {
		const ends = sp(
			`<p:spPr><a:ln w="19050"><a:headEnd type="triangle" w="med" len="lg"/><a:tailEnd type="oval"/></a:ln></p:spPr>`
		).lineEnds
		assert(ends, 'a line with arrowheads surfaces lineEnds')
		assertEqual(ends.head.type, 'triangle', 'head type')
		assertEqual(ends.head.width, 'med', 'head width')
		assertEqual(ends.head.length, 'lg', 'head length')
		assertEqual(ends.tail.type, 'oval', 'tail type')
		assertEqual(ends.tail.width, null, 'tail has no explicit width → null')
		assertEqual(ends.tail.length, null, 'tail has no explicit length → null')
	})

	test('a headEnd with no @type defaults its type to none', () => {
		const ends = sp(`<p:spPr><a:ln w="12700"><a:headEnd/></a:ln></p:spPr>`).lineEnds
		assertEqual(ends.head.type, 'none', 'a bare headEnd reads type "none"')
		assertEqual(ends.tail, null, 'no tailEnd → null tail')
	})

	test('a line with neither end, and no line at all, both read null', () => {
		assertEqual(sp(`<p:spPr><a:ln w="12700"/></p:spPr>`).lineEnds, null, 'a plain line has no ends')
		assertEqual(sp(`<p:spPr/>`).lineEnds, null, 'no a:ln → null')
	})
})

describe('Shape.gradientFill — the a:path (radial/rectangular) branch', () => {
	test('a path gradient reports kind "path", its shape, a null angle, and its stops', () => {
		const grad = sp(
			`<p:spPr><a:gradFill><a:gsLst>` +
				`<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>` +
				`<a:gs><a:srgbClr val="0000FF"/></a:gs>` +
				`</a:gsLst><a:path path="circle"/></a:gradFill></p:spPr>`
		).gradientFill
		assert(grad, 'an a:gradFill surfaces a gradientFill')
		assertEqual(grad.kind, 'path', 'an a:path gradient is kind "path"')
		assertEqual(grad.path, 'circle', 'the path shape is surfaced')
		assertEqual(grad.angleDeg, null, 'a path gradient has no linear angle')
		assertEqual(grad.stops.length, 2, 'both stops surfaced')
		assertEqual(grad.stops[0].position, 0, 'first stop at 0%')
		assertEqual(grad.stops[1].position, null, 'a stop with no @pos reports a null position')
	})

	test('a non-gradient fill has no gradientFill', () => {
		assertEqual(
			sp(`<p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>`).gradientFill,
			null,
			'a solid fill → null'
		)
	})
})

describe('Shape.innerShadow — inset shadow reads', () => {
	test('reads blur / offset / angle and an srgb colour with alpha, like an outer shadow', () => {
		const shadow = sp(
			`<p:spPr><a:effectLst><a:innerShdw blurRad="63500" dist="25400" dir="16200000">` +
				`<a:srgbClr val="404040"><a:alpha val="60000"/></a:srgbClr></a:innerShdw></a:effectLst></p:spPr>`
		).innerShadow
		assert(shadow, 'an innerShdw surfaces an inner shadow')
		assertEqual(shadow.color, '404040', 'effectiveHex of the inner shadow colour')
		assert(Math.abs(shadow.alpha - 0.6) < 1e-9, `alpha 60000 → 0.6, got ${shadow.alpha}`)
		assertEqual(shadow.blurPt, 5, 'blurRad 63500 EMU → 5pt')
		assertEqual(shadow.offsetPt, 2, 'dist 25400 EMU → 2pt')
		assertEqual(shadow.angleDeg, 270, 'dir 16200000 (60000ths) → 270°')
	})

	test('an outer shadow is not read as an inner one, and vice versa', () => {
		const outerOnly = sp(`<p:spPr><a:effectLst><a:outerShdw dist="12700"/></a:effectLst></p:spPr>`)
		assertEqual(outerOnly.innerShadow, null, 'an outerShdw does not surface as innerShadow')
		assert(outerOnly.shadow, 'the outerShdw still surfaces as shadow')
		const innerOnly = sp(`<p:spPr><a:effectLst><a:innerShdw dist="12700"/></a:effectLst></p:spPr>`)
		assertEqual(innerOnly.shadow, null, 'an innerShdw does not surface as the outer shadow')
		assert(innerOnly.innerShadow, 'the innerShdw still surfaces as innerShadow')
	})

	test('no innerShdw → null', () => {
		assertEqual(sp(`<p:spPr><a:effectLst/></p:spPr>`).innerShadow, null, 'an effectLst with no innerShdw → null')
	})
})

describe('Shape.glow — coloured halo reads', () => {
	test('reads its radius (EMU) and an srgb colour with alpha', () => {
		const glow = sp(
			`<p:spPr><a:effectLst><a:glow rad="101600"><a:srgbClr val="FFC000"><a:alpha val="75000"/></a:srgbClr>` +
				`</a:glow></a:effectLst></p:spPr>`
		).glow
		assert(glow, 'an a:glow surfaces a glow')
		assertEqual(glow.radiusPt, 8, 'rad 101600 EMU → 8pt')
		assertEqual(glow.color, 'FFC000', 'effectiveHex of the glow colour')
		assert(Math.abs(glow.alpha - 0.75) < 1e-9, `alpha 75000 → 0.75, got ${glow.alpha}`)
	})

	test('a scheme-coloured glow surfaces its colorToken even when unresolved', () => {
		const glow = sp(
			`<p:spPr><a:effectLst><a:glow rad="50800"><a:schemeClr val="accent2"/></a:glow></a:effectLst></p:spPr>`
		).glow
		assert(glow, 'the glow surfaces')
		assertEqual(glow.colorToken, 'accent2', 'the scheme token is reported for a downstream resolver')
		assertEqual(glow.color, null, 'with empty colour maps the scheme colour does not resolve')
	})

	test('no glow → null', () => {
		assertEqual(sp(`<p:spPr><a:effectLst/></p:spPr>`).glow, null, 'an effectLst with no glow → null')
	})
})

describe('Shape.reflection / Shape.softEdge — read-only effects', () => {
	test('reflection decodes distances (EMU), directions (60000ths) and alpha/pos (1000ths of a percent)', () => {
		const refl = sp(
			`<p:spPr><a:effectLst><a:reflection blurRad="6350" stA="50000" stPos="0" endA="300" endPos="55000" ` +
				`dist="25400" dir="5400000" fadeDir="5400000"/></a:effectLst></p:spPr>`
		).reflection
		assert(refl, 'an a:reflection surfaces a reflection')
		assertEqual(refl.blurPt, 0.5, 'blurRad 6350 EMU → 0.5pt')
		assertEqual(refl.offsetPt, 2, 'dist 25400 EMU → 2pt')
		assertEqual(refl.angleDeg, 90, 'dir 5400000 (60000ths) → 90°')
		assertEqual(refl.fadeAngleDeg, 90, 'fadeDir 5400000 (60000ths) → 90°')
		assertEqual(refl.startAlpha, 0.5, 'stA 50000 → 0.5')
		assertEqual(refl.startPos, 0, 'stPos 0 → 0')
		assert(Math.abs(refl.endAlpha - 0.003) < 1e-9, `endA 300 → 0.003, got ${refl.endAlpha}`)
		assertEqual(refl.endPos, 0.55, 'endPos 55000 → 0.55')
	})

	test('an absent reflection attribute is omitted, not zeroed', () => {
		const refl = sp(`<p:spPr><a:effectLst><a:reflection blurRad="6350"/></a:effectLst></p:spPr>`).reflection
		assert(refl, 'a bare reflection still surfaces')
		assertEqual(refl.blurPt, 0.5, 'the one present attribute is decoded')
		assertEqual(refl.offsetPt, undefined, 'an absent distance is omitted (undefined), not 0')
		assertEqual(refl.startAlpha, undefined, 'an absent alpha is omitted')
	})

	test('softEdge decodes its feather radius; a bare softEdge reads radius 0', () => {
		assertEqual(
			sp(`<p:spPr><a:effectLst><a:softEdge rad="38100"/></a:effectLst></p:spPr>`).softEdge.radiusPt,
			3,
			'rad 38100 EMU → 3pt'
		)
		assertEqual(
			sp(`<p:spPr><a:effectLst><a:softEdge/></a:effectLst></p:spPr>`).softEdge.radiusPt,
			0,
			'a softEdge with no @rad → 0pt'
		)
	})

	test('no reflection / no softEdge → null', () => {
		const none = sp(`<p:spPr><a:effectLst/></p:spPr>`)
		assertEqual(none.reflection, null, 'an effectLst with no reflection → null')
		assertEqual(none.softEdge, null, 'an effectLst with no softEdge → null')
	})
})

describe('Shape.patternFill — preset hatch reads', () => {
	test('reads the preset name and both colours resolved to literal hex', () => {
		const pat = sp(
			`<p:spPr><a:pattFill prst="diagCross"><a:fgClr><a:srgbClr val="C00000"/></a:fgClr>` +
				`<a:bgClr><a:srgbClr val="FFFF00"/></a:bgClr></a:pattFill></p:spPr>`
		).patternFill
		assert(pat, 'an a:pattFill surfaces a pattern fill')
		assertEqual(pat.preset, 'diagCross', 'the preset pattern name is surfaced')
		assertEqual(pat.foreground.effectiveHex, 'C00000', 'foreground resolved to literal hex')
		assertEqual(pat.background.effectiveHex, 'FFFF00', 'background resolved to literal hex')
	})

	test('a pattern with missing fg/bg colours reports null for the absent side', () => {
		const pat = sp(
			`<p:spPr><a:pattFill prst="pct50"><a:fgClr><a:srgbClr val="000000"/></a:fgClr></a:pattFill></p:spPr>`
		).patternFill
		assert(pat, 'the pattFill still surfaces')
		assertEqual(pat.foreground.effectiveHex, '000000', 'the present foreground resolves')
		assertEqual(pat.background, null, 'an absent a:bgClr → null background')
	})

	test('a solid fill is not read as a pattern fill', () => {
		assertEqual(
			sp(`<p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>`).patternFill,
			null,
			'a solid fill → null patternFill'
		)
	})
})

// The two effects the writer emits get a write→read fidelity leg: author a shape
// carrying each with the write API, read it back, and assert the bytes round-trip.
describe('Shape effects/fill — write→read fidelity', () => {
	/** The authored rect autoShape, located in the read model. */
	function rectOf(presentation) {
		const rect = firstShape(presentation, (s) => s.shapeType === 'autoShape' && s.presetGeometry === 'rect')
		assert(rect, 'the authored rect is read back')
		return rect
	}

	test('an authored inner shadow round-trips through Shape.innerShadow', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addShape(ShapeType.rect, {
				x: 1,
				y: 1,
				w: 3,
				h: 1,
				fill: { color: 'CCCCCC' },
				shadow: { type: 'inner', color: 'C0504D', blur: 4, offset: 2, angle: 90, transparency: 25 },
			})
		})
		const shadow = rectOf(presentation).innerShadow
		assert(shadow, 'the authored inner shadow reads back')
		assertEqual(shadow.color, 'C0504D', 'authored inner shadow colour')
		assertEqual(shadow.blurPt, 4, 'blur 4pt round-trips')
		assertEqual(shadow.offsetPt, 2, 'offset 2pt round-trips')
		assertEqual(shadow.angleDeg, 90, 'angle 90° round-trips')
		assert(Math.abs(shadow.alpha - 0.75) < 1e-9, `transparency 25 → alpha 0.75, got ${shadow.alpha}`)
	})

	test('an authored pattern fill round-trips through Shape.patternFill', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addShape(ShapeType.rect, {
				x: 1,
				y: 1,
				w: 3,
				h: 1,
				fill: { type: 'pattern', pattern: { preset: 'pct50', fgColor: 'C00000', bgColor: 'FFFF00' } },
			})
		})
		const pat = rectOf(presentation).patternFill
		assert(pat, 'the authored pattern fill reads back')
		assertEqual(pat.preset, 'pct50', 'authored preset round-trips')
		assertEqual(pat.foreground.effectiveHex, 'C00000', 'authored fgColor round-trips')
		assertEqual(pat.background.effectiveHex, 'FFFF00', 'authored bgColor round-trips')
	})

	test.skipIf(!validatorInstalled)('the authored inner-shadow + pattern-fill decks are schema-valid', async () => {
		const shadowBuf = (
			await authorRead((pres) => {
				pres.addSlide().addShape(ShapeType.rect, {
					x: 1,
					y: 1,
					w: 3,
					h: 1,
					shadow: { type: 'inner', color: 'C0504D', blur: 4, offset: 2, angle: 90, transparency: 25 },
				})
			})
		).buf
		assertEqual((await schemaErrors(shadowBuf)).length, 0, 'inner-shadow deck validates')
		const patternBuf = (
			await authorRead((pres) => {
				pres.addSlide().addShape(ShapeType.rect, {
					x: 1,
					y: 1,
					w: 3,
					h: 1,
					fill: { type: 'pattern', pattern: { preset: 'pct50', fgColor: 'C00000', bgColor: 'FFFF00' } },
				})
			})
		).buf
		assertEqual((await schemaErrors(patternBuf)).length, 0, 'pattern-fill deck validates')
	})
})
