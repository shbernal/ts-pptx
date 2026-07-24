// Read-model coverage for the per-shape / per-paragraph STYLE accessors:
// Shape.lineWidthPt / adjustValues / gradientStops / hidden, and
// Paragraph.align / spaceBeforePt / spaceAfterPt / marginLeftPt / indentPt /
// bullet. These power a faithful style dump of a source slide (e.g. a downstream
// `style.json` bundle), so the reads must hold against both real
// PowerPoint-authored XML and our own serializer.
//
// Strategy:
// - mixed.pptx is genuine Office output and carries paragraph formatting and
//   group geometry — assert those reads there.
// - theme-colors.pptx, gradient-fill.pptx, and preset-geometry.pptx are minimal
//   desktop PowerPoint fixtures for the style-accessor constructs that would be
//   circular if tested only through this library's writer.

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation, AutoShape, Picture } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

/**
 * Parse a standalone shape-tree XML string and wrap its first `p:<local>`
 * descendant in `Kind`. `absoluteFrame` and `recolor` read only the shape's own
 * DOM (never the owning `Slide`), so a throwaway slide stand-in is enough to
 * exercise them off-fixture. Selecting by tag lets a nested-group fixture wrap the
 * innermost `p:sp` rather than the enclosing `p:grpSp`.
 */
function shapeFromXml(Kind, local, innerXml) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}">${innerXml}</p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, local)[0]
	if (!el) throw new Error(`no <p:${local}> in the supplied XML`)
	return new Kind(el, /** stand-in slide */ {})
}

/** A `p:pic` proxy whose blip carries the given recolour child XML. */
function pictureWithBlipChild(innerXml) {
	return shapeFromXml(Picture, 'pic', `<p:pic><p:blipFill><a:blip>${innerXml}</a:blip></p:blipFill></p:pic>`)
}

const __dirname = path.dirname(fileURLToPath(import.meta.url))

function fixturePath(name) {
	return path.join(__dirname, 'fixtures', `${name}.pptx`)
}

async function open(name) {
	return Presentation.load(await readFile(fixturePath(name)))
}

/** Flatten a shape list, descending into groups. */
function allShapes(shapes) {
	return shapes.flatMap((shape) => (shape.shapeType === 'group' ? [shape, ...allShapes(shape.shapes)] : [shape]))
}

function leafShapes(shapes) {
	return allShapes(shapes).filter((shape) => shape.shapeType !== 'group')
}

function shapeNamed(slide, name) {
	const shape = allShapes(slide.shapes).find((s) => s.name === name)
	assert(shape, `expected shape named ${name}`)
	return shape
}

/** Every paragraph of every (flattened) shape on a slide. */
function allParagraphs(slide) {
	return allShapes(slide.shapes)
		.filter((shape) => shape.hasTextFrame)
		.flatMap((shape) => shape.textFrame.paragraphs)
}

describe('Shape style reads — real PowerPoint XML (mixed.pptx)', () => {
	test('lineWidthPt converts a:ln/@w (EMU) to points', async () => {
		// slide5/slide6 draw connectors/borders with <a:ln w="15875"> = 1.25pt.
		const slide = (await open('mixed')).slides[5]
		const widths = allShapes(slide.shapes)
			.map((shape) => shape.lineWidthPt)
			.filter((w) => w !== null)
		assert(widths.length > 0, 'expected at least one shape with an explicit line width')
		assert(widths.includes(1.25), `expected a 1.25pt line (15875 EMU); got ${JSON.stringify(widths)}`)
	})

	test('shapes without a hidden flag report hidden=false', async () => {
		const slide = (await open('mixed')).slides[5]
		for (const shape of allShapes(slide.shapes)) {
			assertEqual(shape.hidden, false, `${shape.name || shape.shapeType} has no @hidden, so reads false`)
		}
	})
})

describe('Paragraph style reads — real PowerPoint XML (mixed.pptx slide7)', () => {
	// slide7 (index 6) is a bulleted, multi-level outline authored in PowerPoint:
	// algn, a:spcBef/a:spcAft (spcPts), marL/indent, and buChar/buNone bullets.
	async function slide7Paragraphs() {
		const paragraphs = allParagraphs((await open('mixed')).slides[6])
		assert(paragraphs.length > 5, `expected a multi-paragraph outline, got ${paragraphs.length}`)
		return paragraphs
	}

	test('align reads the a:pPr/@algn token', async () => {
		const aligns = (await slide7Paragraphs()).map((p) => p.align)
		assert(aligns.includes('ctr'), `expected a centered paragraph; got ${JSON.stringify([...new Set(aligns)])}`)
		// Paragraphs with no @algn report null (inherited), not a default token.
		assert(aligns.includes(null), 'expected at least one paragraph with inherited (null) alignment')
	})

	test('spacing reads a:spcPts as points, and percentage spacing (a:spcPct) as null', async () => {
		const paragraphs = await slide7Paragraphs()
		// One paragraph carries <a:spcAft><a:spcPts val="600"/> = 6pt.
		const afters = paragraphs.map((p) => p.spaceAfterPt).filter((v) => v !== null)
		assert(afters.includes(6), `expected a 6pt space-after; got ${JSON.stringify(afters)}`)
		// Every a:spcBef in this slide is a percentage (a:spcPct), which has no
		// fixed point value, so spaceBeforePt is null throughout.
		const befores = paragraphs.map((p) => p.spaceBeforePt)
		assert(
			befores.every((v) => v === null),
			`percentage space-before should read null; got ${JSON.stringify([...new Set(befores)])}`
		)
	})

	test('marginLeftPt and indentPt convert a:pPr EMU attributes to points', async () => {
		// A hanging-indent line: marL="342900" indent="-342900" → 27pt / -27pt.
		const hanging = (await slide7Paragraphs()).find((p) => p.marginLeftPt === 27)
		assert(hanging, 'expected a paragraph with marL 342900 (27pt)')
		assertEqual(hanging.indentPt, -27, 'matching hanging indent (indent -342900)')
	})

	test('bullet distinguishes buChar glyphs from explicit buNone', async () => {
		const bullets = (await slide7Paragraphs()).map((p) => p.bullet)
		assert(
			bullets.some((b) => b?.startsWith('char:')),
			`expected a glyph bullet; got ${JSON.stringify([...new Set(bullets)])}`
		)
		assert(bullets.includes('none'), 'expected an explicitly un-bulleted paragraph (a:buNone)')
	})
})

describe('Picture SVG blip reads (image.pptx)', () => {
	test('svgRelId / svgPartName resolve the asvg:svgBlip extension embed', async () => {
		const presentation = await open('image')
		const pictures = presentation.slides
			.flatMap((slide) => allShapes(slide.shapes))
			.filter((s) => s.shapeType === 'picture')
		const svgPic = pictures.find((p) => p.svgRelId !== null)
		assert(svgPic, 'expected a picture carrying an SVG blip extension')
		const svgPart = svgPic.svgPartName
		assert(svgPart && svgPart.endsWith('.svg'), `svgPartName resolves to the .svg part; got ${svgPart}`)
		assert(presentation.opc.part(svgPart), `svg part ${svgPart} exists in the package`)
		// The raster fallback (imageRelId) and the vector (svgRelId) are distinct rels.
		assert(svgPic.imageRelId !== svgPic.svgRelId, 'raster fallback and SVG embed are different relationships')
	})

	test('a raster-only picture has no svgRelId', async () => {
		const presentation = await open('image')
		const pictures = presentation.slides
			.flatMap((slide) => allShapes(slide.shapes))
			.filter((s) => s.shapeType === 'picture')
		const rasterOnly = pictures.find((p) => p.imagePartName && !p.imagePartName.endsWith('.svg') && p.svgRelId === null)
		assert(rasterOnly, 'expected at least one raster-only picture with a null svgRelId')
	})

	test('mediaKind / mediaPartName classify the raster+SVG pairing and a raster-only picture', async () => {
		const presentation = await open('image')
		const pictures = presentation.slides
			.flatMap((slide) => allShapes(slide.shapes))
			.filter((s) => s.shapeType === 'picture')

		// The picture with both a raster embed and an SVG extension reads 'both',
		// and mediaPartName prefers the raster part.
		const both = pictures.find((p) => p.imageRelId !== null && p.svgRelId !== null)
		assert(both, 'expected a picture with a raster fallback and an SVG extension')
		assertEqual(both.mediaKind, 'both', "a raster+SVG picture is 'both'")
		assertEqual(both.mediaPartName, both.imagePartName, 'mediaPartName prefers the raster part when present')

		const rasterOnly = pictures.find((p) => p.imageRelId !== null && p.svgRelId === null)
		assert(rasterOnly, 'expected a raster-only picture')
		assertEqual(rasterOnly.mediaKind, 'raster', "a raster-only picture is 'raster'")
		assertEqual(rasterOnly.mediaPartName, rasterOnly.imagePartName, 'mediaPartName is the raster part')
	})
})

describe('Shape style reads — minimal real PowerPoint fixtures', () => {
	test('lineWidthPt reads an explicit 2pt theme-colour line', async () => {
		const shape = shapeNamed((await open('theme-colors')).slides[0], 'accent1-line-accent2-2pt')
		assertEqual(shape.lineWidthPt, 2, '<a:ln w="25400"> is 2pt')
		assertEqual(shape.lineSchemeColor, 'accent2', 'line is a real PowerPoint scheme colour')
		assertEqual(shape.resolvedLine.hex, 'EA6312', 'accent2 resolves through the non-default Ion theme')
	})

	test('adjustValues exposes PowerPoint-authored avLst handles', async () => {
		const slide = (await open('preset-geometry')).slides[0]

		const roundRect = shapeNamed(slide, 'roundRect-adj')
		assertEqual(roundRect.presetGeometry, 'roundRect', 'fixture shape is a roundRect')
		assertEqual(roundRect.adjustValues.adj, 'val 12000', 'roundRect writes its single adj handle')

		const chevron = shapeNamed(slide, 'chevron-adj')
		assertEqual(chevron.presetGeometry, 'chevron', 'fixture shape is a chevron')
		assertEqual(chevron.adjustValues.adj, 'val 35000', 'chevron writes its single adj handle')

		const blockArc = shapeNamed(slide, 'blockArc-adj1-adj2-adj3')
		assertEqual(blockArc.presetGeometry, 'blockArc', 'fixture shape is a blockArc')
		assertEqual(blockArc.adjustValues.adj1, 'val 15000', 'blockArc first guide is present')
		assertEqual(blockArc.adjustValues.adj2, 'val 7200000', 'blockArc angle guide is present')
		assertEqual(blockArc.adjustValues.adj3, 'val 30000000', 'blockArc second angle guide is present')

		const rect = shapeNamed(slide, 'rect-no-adjust')
		assertEqual(Object.keys(rect.adjustValues).length, 0, 'a plain rect has no adjust handles')
	})

	test('gradientStops reads PowerPoint-authored gsLst stops with position + colour split', async () => {
		const slide = (await open('gradient-fill')).slides[0]

		const linear2 = shapeNamed(slide, 'grad-linear-2')
		assertEqual(linear2.gradientStops.length, 2, 'two-stop linear gradient')
		assertEqual(linear2.gradientStops[0].position, 0, 'first stop at 0%')
		assertEqual(linear2.gradientStops[0].color, '451DC7', 'first stop is explicit srgb')
		assertEqual(linear2.gradientStops[1].position, 1, 'last stop at 100%')
		assertEqual(linear2.gradientStops[1].color, 'FFFFFF', 'last stop is explicit srgb')

		const linear3 = shapeNamed(slide, 'grad-linear-3-scheme')
		assertEqual(linear3.gradientStops.length, 3, 'three-stop gradient')
		assertEqual(linear3.gradientStops[0].schemeColor, 'accent1', 'first stop is a scheme colour')
		assertEqual(linear3.gradientStops[0].effectiveHex, 'B01513', 'scheme stop resolves through the Ion theme')
		assertEqual(linear3.gradientStops[1].position, 0.5, 'middle stop at 50%')
		assertEqual(linear3.gradientStops[1].color, '1EB4D2', 'middle stop is explicit srgb')

		const radial = shapeNamed(slide, 'grad-radial')
		assertEqual(radial.gradientStops.length, 2, 'radial/path gradient still exposes its stops')

		const solid = shapeNamed(slide, 'solid-control')
		assertEqual(solid.gradientStops, null, 'a solid-filled shape reports null gradientStops')
	})
})

describe('Shape line dash / explicit no-line reads (off-fixture)', () => {
	// lineDash and lineNoFill read only the shape's own spPr/a:ln, so hand-authored
	// OOXML exercises every branch without a round-trip through this library's writer.
	const sp = (spPr) => shapeFromXml(AutoShape, 'sp', `<p:sp>${spPr}</p:sp>`)

	test('lineDash reads a:ln/a:prstDash/@val', () => {
		const dashed = sp(
			'<p:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr>'
		)
		assertEqual(dashed.lineDash, 'dash', 'prstDash val is surfaced verbatim')
	})

	test('lineDash is null for a solid line and for no line at all', () => {
		const solid = sp('<p:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>')
		assertEqual(solid.lineDash, null, 'a line with no prstDash is solid → null')
		const noLine = sp('<p:spPr/>')
		assertEqual(noLine.lineDash, null, 'no a:ln at all → null')
	})

	test('lineNoFill distinguishes an explicit a:ln/a:noFill from an inherited line', () => {
		const explicitNone = sp('<p:spPr><a:ln><a:noFill/></a:ln></p:spPr>')
		assertEqual(explicitNone.lineNoFill, true, 'explicit <a:ln><a:noFill/> reads true')
		assertEqual(explicitNone.lineDash, null, 'a no-fill line carries no dash')

		const solid = sp('<p:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>')
		assertEqual(solid.lineNoFill, false, 'a solid-filled line is not a no-line')

		const inherited = sp('<p:spPr/>')
		assertEqual(inherited.lineNoFill, false, 'no a:ln (inherited line) is not an explicit no-line')
	})

	// lineGradient points the (fixture-validated, see gradient-fill.pptx) a:gradFill
	// reader at the a:ln container instead of spPr. Explicit srgb stops resolve
	// without a theme, so a minimal themeContext stub is enough off-fixture.
	const spGrad = (spPr) => {
		const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:sp>${spPr}</p:sp></p:spTree>`
		const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
		const el = spTree.getElementsByTagNameNS(P_NS, 'sp')[0]
		// Minimal slide fake: only themeContext is exercised by these unit reads.
		return new AutoShape(el, /** @type {any} */ ({ themeContext: () => ({}) }))
	}

	test('lineGradient reads a:ln/a:gradFill stops + linear angle', () => {
		const shape = spGrad(
			'<p:spPr><a:ln w="57150"><a:gradFill><a:gsLst>' +
				'<a:gs pos="0"><a:srgbClr val="451DC7"/></a:gs>' +
				'<a:gs pos="100000"><a:srgbClr val="BEADF3"/></a:gs>' +
				'</a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill></a:ln></p:spPr>'
		)
		const grad = shape.lineGradient
		assert(grad, 'a gradient-stroked line surfaces a lineGradient')
		assertEqual(grad.kind, 'linear', 'a:lin ⇒ linear line gradient')
		assertEqual(grad.angleDeg, 0, 'a:lin/@ang (60000ths) ÷ 60000 ⇒ degrees')
		assertEqual(grad.stops.length, 2, 'both stops surfaced')
		assertEqual(grad.stops[0].position, 0, 'first stop at 0%')
		assertEqual(grad.stops[0].effectiveHex, '451DC7', 'explicit srgb stop resolves to itself')
		assertEqual(grad.stops[1].position, 1, 'last stop at 100%')
		assertEqual(grad.stops[1].effectiveHex, 'BEADF3', 'second explicit srgb stop resolves to itself')
	})

	test('lineGradient is null for a solid line and when there is no a:ln', () => {
		const solid = spGrad('<p:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>')
		assertEqual(solid.lineGradient, null, 'a solid-stroked line has no lineGradient')
		const noLine = spGrad('<p:spPr/>')
		assertEqual(noLine.lineGradient, null, 'no a:ln ⇒ null lineGradient')
	})
})

describe('TextFrame.resolvedAnchor — real PowerPoint XML (layout-placeholder-bodypr.pptx)', () => {
	// slide1 carries two placeholders whose own <a:bodyPr/> sets no @anchor, so the
	// effective vertical anchor must come from the layout placeholder a:bodyPr:
	// - Title 1 inherits the layout title anchor="b" (and the fixture text says so).
	// - Content Placeholder 2 inherits the layout body anchor="ctr".
	// bodyProperties.anchor (own attribute only) stays null for both.
	test('a placeholder title inherits its anchor from the layout bodyPr', async () => {
		const title = shapeNamed((await open('layout-placeholder-bodypr')).slides[0], 'Title 1')
		assertEqual(title.textFrame.bodyProperties?.anchor ?? null, null, 'the slide bodyPr sets no own @anchor')
		assertEqual(title.textFrame.resolvedAnchor, 'b', 'inherits the layout title anchor="b"')
	})

	test('a placeholder body inherits a different anchor from the layout bodyPr', async () => {
		const body = shapeNamed((await open('layout-placeholder-bodypr')).slides[0], 'Content Placeholder 2')
		assertEqual(body.textFrame.bodyProperties?.anchor ?? null, null, 'the slide bodyPr sets no own @anchor')
		assertEqual(body.textFrame.resolvedAnchor, 'ctr', 'inherits the layout body anchor="ctr"')
	})
})

describe('Theme colour resolution — real PowerPoint XML (theme-colors.pptx)', () => {
	test('resolvedFill resolves a scheme fill to the theme hex, and an explicit fill to itself', async () => {
		const slide = (await open('theme-colors')).slides[0]
		const scheme = shapeNamed(slide, 'accent1-plain')
		assertEqual(scheme.resolvedFill.hex, 'B01513', 'accent1 resolves to the Ion theme accent1 hex')
		// The raw read still reports the unresolved token — resolution is opt-in.
		assertEqual(scheme.fillColor, null, 'fillColor still reports null for a scheme-coloured fill')

		const explicit = shapeNamed(slide, 'explicit-srgb-fill')
		assertEqual(explicit.resolvedFill.hex, 'FF0000', 'an explicit srgb fill resolves to itself')
	})

	test('resolvedLine resolves a scheme line colour; null when there is no solid fill', async () => {
		const lined = shapeNamed((await open('theme-colors')).slides[0], 'accent1-line-accent2-2pt')
		assertEqual(lined.resolvedLine.hex, 'EA6312', 'accent2 line resolves to the Ion theme accent2 hex')
		// A gradient-filled shape has no a:solidFill to resolve as a fill colour.
		const gradient = shapeNamed((await open('gradient-fill')).slides[0], 'grad-linear-3-scheme')
		assertEqual(gradient.resolvedFill, null, 'a gradient fill has no a:solidFill to resolve')
	})

	test('resolvedFill reports the base hex + raw transforms and the applied effectiveHex', async () => {
		const shape = shapeNamed((await open('theme-colors')).slides[0], 'accent1-lm60-lo40')
		const fill = shape.resolvedFill
		assertEqual(fill.hex, 'B01513', 'base colour stays the theme hex')
		assertEqual(fill.transforms.length, 2, 'lumMod/lumOff transform children reported')
		assertEqual(fill.transforms[0].name, 'lumMod', 'first transform is lumMod')
		assertEqual(fill.transforms[0].value, '60000', 'lumMod raw @val is preserved')
		assertEqual(fill.transforms[1].name, 'lumOff', 'second transform is lumOff')
		assertEqual(fill.transforms[1].value, '40000', 'lumOff raw @val is preserved')
		assertEqual(fill.effectiveHex, 'ED5654', 'effectiveHex applies the PowerPoint-authored transforms')
	})

	test('Run.resolvedColor resolves a scheme run colour to the theme hex', async () => {
		const shape = shapeNamed((await open('theme-colors')).slides[0], 'text-accent5-run')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.schemeColor, 'accent5', 'the raw read reports the scheme token')
		assertEqual(run.resolvedColor.hex, '54849A', 'accent5 resolves to the Ion theme accent5 hex')
	})
})

describe('Style-matrix fill/line resolution — real PowerPoint XML (multi-theme.pptx)', () => {
	// The `style-matrix-default` shape carries NO explicit spPr fill or line: its
	// colour comes only through the p:style fillRef (idx 1, accent1) and lnRef
	// (idx 2, accent1 + shade). resolvedFill/resolvedLine must walk that style
	// matrix the same way the `theme: 'preserve'` flatten path bakes it.
	test('resolvedFill walks a p:style fillRef when the shape has no explicit fill', async () => {
		const shape = shapeNamed((await open('multi-theme')).slides[0], 'style-matrix-default')
		// Nothing explicit on the shape itself — the raw reads still report null.
		assertEqual(shape.fillColor, null, 'no explicit srgb fill')
		assertEqual(shape.fillSchemeColor, null, 'no explicit scheme fill')

		const fill = shape.resolvedFill
		assert(fill, 'the style fillRef resolves to a fill colour')
		assertEqual(fill.hex, 'B01513', 'fillRef accent1 resolves through the Ion theme')
		assertEqual(fill.transforms.length, 0, 'the fillStyleLst idx-1 solid carries no transform')
		assertEqual(fill.effectiveHex, 'B01513', 'effective fill is the plain accent1 hex')
	})

	test('resolvedLine walks a p:style lnRef, carrying the ref colour transform', async () => {
		const shape = shapeNamed((await open('multi-theme')).slides[0], 'style-matrix-default')
		assertEqual(shape.lineColor, null, 'no explicit srgb line')
		assertEqual(shape.lineSchemeColor, null, 'no explicit scheme line')

		const line = shape.resolvedLine
		assert(line, 'the style lnRef resolves to a line colour')
		assertEqual(line.hex, 'B01513', 'lnRef accent1 resolves through the Ion theme')
		assertEqual(line.transforms.length, 1, 'the lnRef shade transform is carried onto the resolved colour')
		assertEqual(line.transforms[0].name, 'shade', 'the carried transform is the authored shade')
		assertEqual(line.transforms[0].value, '15000', 'with its raw @val preserved')
	})

	test('an explicit spPr fill/line still wins over the style matrix', async () => {
		// scheme-accent1-fill has explicit solidFill accent1 + an explicit accent2 line,
		// alongside a p:style fillRef/lnRef — the explicit spPr children must govern.
		const shape = shapeNamed((await open('multi-theme')).slides[0], 'scheme-accent1-fill')
		assertEqual(shape.resolvedFill.hex, 'B01513', 'explicit accent1 fill resolves to the Ion accent1 hex')
		assertEqual(shape.resolvedLine.hex, 'EA6312', 'explicit accent2 line wins over the lnRef accent1')
	})
})

describe('Placeholder-inherited run colour — real PowerPoint XML (multi-theme.pptx slide 2)', () => {
	// Slide 2 carries two placeholders authored by desktop PowerPoint:
	// - inherited-title: a run with NO own colour, inheriting through the master
	//   titleStyle lvl1 → schemeClr tx2 → clrMap tx2=lt2 → theme lt2 = EBEBEB.
	// - explicit-body: a run with an explicit srgbClr FF00FF (the negative control).
	// Run.resolvedColor must walk the placeholder/list-style chain for the first and
	// keep the explicit colour for the second.
	test('a colourless placeholder run resolves through the master text style', async () => {
		const shape = shapeNamed((await open('multi-theme')).slides[1], 'inherited-title')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.color, null, 'the run sets no explicit srgb colour')
		assertEqual(run.schemeColor, null, 'the run sets no explicit scheme colour')
		assertEqual(run.resolvedColor.hex, 'EBEBEB', 'inherits titleStyle tx2 → lt2 from the master')
		assertEqual(run.resolvedColor.effectiveHex, 'EBEBEB', 'no transforms, so effective equals base')
	})

	test('an explicit run colour still wins over the inherited placeholder colour', async () => {
		const shape = shapeNamed((await open('multi-theme')).slides[1], 'explicit-body')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.color, 'FF00FF', 'the run carries an explicit srgb colour')
		assertEqual(run.resolvedColor.hex, 'FF00FF', 'the explicit colour governs, not the inherited body colour')
	})
})

describe('Placeholder-inherited run size + typeface — real PowerPoint XML (multi-theme.pptx slide 2)', () => {
	// The same slide-2 placeholders the colour leg uses, now read for SIZE/FACE.
	// inherited-title sets no own sz/latin, so both resolve through the master
	// titleStyle lvl1 (sz=4200 → 42pt; latin +mj-lt → theme major font Century Gothic).
	// explicit-body sets an explicit colour but no sz/latin, so size/face still
	// resolve through the body chain (master bodyStyle lvl1: sz=2000 → 20pt, +mj-lt).
	test('a placeholder title run with no own size/face resolves both through the master text style', async () => {
		const shape = shapeNamed((await open('multi-theme')).slides[1], 'inherited-title')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.fontSizePt, null, 'the run sets no own @sz')
		assertEqual(run.fontName, null, 'the run sets no own a:latin')
		assertEqual(run.resolvedSizePt, 42, 'inherits titleStyle sz=4200 from the master')
		assertEqual(run.resolvedFontFace, 'Century Gothic', '+mj-lt resolves through the theme major font')
	})

	test('a colourful body placeholder run still resolves its inherited size/face', async () => {
		const shape = shapeNamed((await open('multi-theme')).slides[1], 'explicit-body')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.fontSizePt, null, 'the run sets no own @sz')
		assertEqual(run.resolvedSizePt, 20, 'inherits bodyStyle lvl1 sz=2000 from the master')
		assertEqual(run.resolvedFontFace, 'Century Gothic', 'inherits the body +mj-lt face through the theme')
	})

	test("a non-placeholder run's own size wins; its face falls back to the default text style", async () => {
		// text-accent5-run is a plain text box (no placeholder): own sz=2400 governs the
		// size, and with no own a:latin and no placeholder chain the face falls through to
		// the presentation's p:defaultTextStyle (+mn-lt → the theme minor font, here Ion's
		// Century Gothic). Before defaultTextStyle joined the chain this read as null.
		const shape = shapeNamed((await open('theme-colors')).slides[0], 'text-accent5-run')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.resolvedSizePt, 24, "the run's own sz=2400 is reported as 24pt")
		assertEqual(run.resolvedFontFace, 'Century Gothic', 'inherits +mn-lt from p:defaultTextStyle')
	})
})

describe('Placeholder-inherited run bold — real PowerPoint XML', () => {
	// Sibling of the size/face leg. The multi-theme master text styles carry an
	// explicit b="0" on their level defRPrs, so a placeholder run with no own @b
	// resolves to false — an INHERITED non-bold, deliberately distinct from a null
	// "sets none and inherits none". A plain text box has no placeholder chain, so
	// its resolvedBold falls back to the run's own value (null here).
	test('a placeholder run with no own @b resolves inherited bold from the master text style', async () => {
		const shape = shapeNamed((await open('multi-theme')).slides[1], 'inherited-title')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.bold, null, 'the run sets no own @b')
		assertEqual(run.resolvedBold, false, 'inherits titleStyle b="0" from the master (explicit non-bold, not null)')
	})

	test('a colourful body placeholder run also resolves its inherited bold', async () => {
		const shape = shapeNamed((await open('multi-theme')).slides[1], 'explicit-body')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.bold, null, 'the run sets no own @b')
		assertEqual(run.resolvedBold, false, 'inherits bodyStyle lvl1 b="0" from the master')
	})

	test('a non-placeholder run reports no inherited bold (own value governs)', async () => {
		const shape = shapeNamed((await open('theme-colors')).slides[0], 'text-accent5-run')
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.resolvedBold, null, 'no own @b and no placeholder chain to inherit from')
	})
})

describe('Picture recolour reads (recolor)', () => {
	test('reads a real PowerPoint a:duotone, preserving the prstClr/srgbClr stop split (image.pptx)', async () => {
		// image.pptx slide2 carries an icon recoloured with the duotone tint trick:
		// <a:duotone><a:prstClr val="black"/><a:srgbClr val="B6D3ED">…</a:srgbClr></a:duotone>.
		const pictures = (await open('image')).slides
			.flatMap((slide) => allShapes(slide.shapes))
			.filter((s) => s.shapeType === 'picture')
		const tinted = pictures.find((p) => p.recolor !== null)
		assert(tinted, 'expected a picture carrying a recolour effect')
		const recolor = tinted.recolor
		assertEqual(recolor.kind, 'duotone', 'a:duotone is read as a duotone recolour')
		assertEqual(recolor.stops.length, 2, 'a duotone has two colour stops')
		assertEqual(recolor.stops[0].presetColor, 'black', 'first stop is the prstClr black')
		assertEqual(recolor.stops[0].color, null, 'a prstClr stop carries no srgb colour')
		assertEqual(recolor.stops[0].schemeColor, null, 'a prstClr stop carries no scheme colour')
		assertEqual(recolor.stops[1].color, 'B6D3ED', 'second stop is the explicit srgb tint')
		assertEqual(recolor.stops[1].presetColor, null, 'an srgb stop carries no preset colour')
	})

	test('a picture with no recolour effect reads null', () => {
		assertEqual(pictureWithBlipChild('').recolor, null, 'a bare blip has no recolour')
	})

	test('clrChange reports its from/to colours, scheme tokens included', () => {
		const recolor = pictureWithBlipChild(
			'<a:clrChange><a:clrFrom><a:srgbClr val="FF0000"/></a:clrFrom><a:clrTo><a:schemeClr val="accent1"/></a:clrTo></a:clrChange>'
		).recolor
		assertEqual(recolor.kind, 'clrChange', 'a:clrChange is read as a clrChange recolour')
		assertEqual(recolor.from.color, 'FF0000', 'clrFrom is the explicit source colour')
		assertEqual(recolor.to.schemeColor, 'accent1', 'clrTo is a scheme token left for the theme resolver')
		assertEqual(recolor.to.color, null, 'a scheme clrTo carries no explicit colour')
	})

	test('grayscl / biLevel / alphaModFix map to their kinds with 0–1 fractions', () => {
		assertEqual(pictureWithBlipChild('<a:grayscl/>').recolor.kind, 'grayscale', 'a:grayscl → grayscale')

		const biLevel = pictureWithBlipChild('<a:biLevel thresh="50000"/>').recolor
		assertEqual(biLevel.kind, 'biLevel', 'a:biLevel → biLevel')
		assertEqual(biLevel.threshold, 0.5, 'thresh 50000 (thousandths of a percent) reads as 0.5')

		const amf = pictureWithBlipChild('<a:alphaModFix amt="40000"/>').recolor
		assertEqual(amf.kind, 'alphaModFix', 'a:alphaModFix → alphaModFix')
		assertEqual(amf.amount, 0.4, 'amt 40000 reads as 0.4')
		// amt is optional and defaults to 100% per the schema.
		assertEqual(pictureWithBlipChild('<a:alphaModFix/>').recolor.amount, 1, 'a missing amt defaults to 1.0')
	})

	test('the first recolour effect in document order wins', () => {
		const recolor = pictureWithBlipChild(
			'<a:grayscl/><a:duotone><a:srgbClr val="111111"/><a:srgbClr val="222222"/></a:duotone>'
		).recolor
		assertEqual(recolor.kind, 'grayscale', 'grayscl precedes the duotone, so it is the one reported')
	})
})

describe('Group-child absolute geometry (absoluteFrame)', () => {
	function assertWithin(actual, expected, tolerance, label) {
		assert(Math.abs(actual - expected) <= tolerance, `${label}: expected ${expected} ± ${tolerance}, got ${actual}`)
	}

	function normalizedDegrees(value) {
		return ((value % 360) + 360) % 360
	}

	test('a top-level shape resolves to its own geometry', async () => {
		const presentation = await (async () => {
			const pres = new TsPptx()
			pres.addSlide().addShape(pres.ShapeType.rect, { x: 1, y: 1, w: 3, h: 1, fill: { color: 'CCCCCC' } })
			return Presentation.load(await pres.stream())
		})()
		const shape = presentation.slides[0].shapes.find((s) => s.shapeType === 'autoShape' && s.presetGeometry === 'rect')
		const frame = shape.absoluteFrame
		assertEqual(frame.left, shape.left, 'an ungrouped shape: absolute left == own left')
		assertEqual(frame.top, shape.top, 'an ungrouped shape: absolute top == own top')
		assertEqual(frame.width, shape.width, 'an ungrouped shape: absolute width == own width')
		assertEqual(frame.height, shape.height, 'an ungrouped shape: absolute height == own height')
		assertEqual(frame.rotation, 0, 'an ungrouped unrotated shape has effective rotation 0')
		assertEqual(frame.flipH, false, 'an ungrouped unflipped shape has effective flipH=false')
		assertEqual(frame.flipV, false, 'an ungrouped unflipped shape has effective flipV=false')
	})

	test('a group child composes its parent group transform (real PowerPoint XML, mixed.pptx slide5)', async () => {
		// One slide5 group translates its children down by 145757 EMU (off.y 3301445
		// vs chOff.y 3155688) with ext == chExt (no scaling). A child whose own
		// a:off.y is 3155688 must therefore resolve to an absolute top of 3301445.
		const slide = (await open('mixed')).slides[4]
		const groups = slide.shapes.filter((s) => s.shapeType === 'group')
		assert(groups.length > 0, 'expected groups on slide5')
		const child = groups.flatMap((g) => g.shapes).find((s) => s.top === 3155688 && s.absoluteFrame)
		assert(child, 'expected a (non-degenerate) group child at raw top 3155688')
		const frame = child.absoluteFrame
		assertEqual(frame.top, 3301445, 'the child top shifts by the group offset (3155688 → 3301445)')
		assertEqual(frame.left, child.left, 'this group does not shift x (off.x == chOff.x)')
		assertEqual(frame.width, child.width, 'ext == chExt, so width is unscaled')
		assertEqual(frame.height, child.height, 'ext == chExt, so height is unscaled')
	})

	test('composes offset and scale through nested groups', () => {
		// outer ratio 2 (ext 8000 / chExt 4000), inner ratio 2 (2000 / 1000):
		// sp (100,100,500,500) → inner → (1200,1200,1000,1000) → outer → (12400,12400,2000,2000).
		const inner = shapeFromXml(
			AutoShape,
			'sp',
			`<p:grpSp>
				<p:grpSpPr><a:xfrm><a:off x="10000" y="10000"/><a:ext cx="8000" cy="8000"/><a:chOff x="0" y="0"/><a:chExt cx="4000" cy="4000"/></a:xfrm></p:grpSpPr>
				<p:grpSp>
					<p:grpSpPr><a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="2000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>
					<p:sp><p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm></p:spPr></p:sp>
				</p:grpSp>
			</p:grpSp>`
		)
		const frame = inner.absoluteFrame
		assertEqual(frame.left, 12400, 'left composes inner then outer offset+scale')
		assertEqual(frame.top, 12400, 'top composes inner then outer offset+scale')
		assertEqual(frame.width, 2000, 'width scales by inner×outer ratio (500 → 2000)')
		assertEqual(frame.height, 2000, 'height scales by inner×outer ratio')
	})

	test('composes scale, rotation, and flips to match PowerPoint ungroup output', async () => {
		const [grouped, ungrouped] = (await open('group-transform')).slides
		const flattenedGroups = allShapes(ungrouped.shapes).filter((shape) => shape.shapeType === 'group')
		assertEqual(flattenedGroups.length, 0, 'slide 2 is PowerPoint-ungrouped ground truth')

		const groupedChildren = leafShapes(grouped.shapes).filter((shape) => shape.name.includes(' child '))
		const ungroupedByName = new Map(leafShapes(ungrouped.shapes).map((shape) => [shape.name, shape]))

		assertEqual(
			groupedChildren.length,
			21,
			'fixture pins the original four groups plus scale/child/nested transform cases'
		)
		assert(
			groupedChildren.some((shape) => shape.name.startsWith('scale-rot child ')),
			'expected scale+rotation group children in the fixture'
		)
		assert(
			groupedChildren.some((shape) => shape.name.startsWith('childrot-in-rot child ')),
			'expected child-owned transform children in the fixture'
		)
		assert(
			groupedChildren.some((shape) => shape.name.startsWith('nested-rot-in-scale child ')),
			'expected nested rotated group children in the fixture'
		)

		for (const child of groupedChildren) {
			const expectedName = child.name.replace(/^(.+?) child /, '$1-ungrouped child ')
			const expected = ungroupedByName.get(expectedName)
			assert(expected, `expected PowerPoint-ungrouped twin "${expectedName}" for "${child.name}"`)

			const frame = child.absoluteFrame
			assert(frame, `${child.name} should have a resolvable absolute frame`)
			assertWithin(frame.left, expected.left, 2, `${child.name} absolute left`)
			assertWithin(frame.top, expected.top, 2, `${child.name} absolute top`)
			assertWithin(frame.width, expected.width, 2, `${child.name} absolute width`)
			assertWithin(frame.height, expected.height, 2, `${child.name} absolute height`)
			assertWithin(
				normalizedDegrees(frame.rotation),
				normalizedDegrees(expected.rotation),
				1e-6,
				`${child.name} effective rotation`
			)
			assertEqual(frame.flipH, expected.flipH, `${child.name} effective flipH`)
			assertEqual(frame.flipV, expected.flipV, `${child.name} effective flipV`)
		}
	})

	test('a shape with no own transform has no resolvable absolute frame', () => {
		const shape = shapeFromXml(AutoShape, 'sp', '<p:sp><p:spPr/></p:sp>')
		assertEqual(shape.absoluteFrame, null, 'no a:xfrm → null')
	})

	test('a degenerate group (zero a:chExt) yields no resolvable frame', () => {
		const inner = shapeFromXml(
			AutoShape,
			'sp',
			`<p:grpSp>
				<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
				<p:sp><p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm></p:spPr></p:sp>
			</p:grpSp>`
		)
		assertEqual(inner.absoluteFrame, null, 'dividing by a zero child extent is degenerate → null')
	})
})

describe('Per-shape rotation / flip (rotation, flipH, flipV)', () => {
	/** A bare `p:sp` whose spPr carries the given a:xfrm XML (or none). */
	function spWithXfrm(xfrmXml) {
		return shapeFromXml(AutoShape, 'sp', `<p:sp><p:spPr>${xfrmXml}</p:spPr></p:sp>`)
	}

	const EPS = 1e-6

	test('rot (60000ths of a degree) reads as degrees; flipV reads true, flipH false', () => {
		// 2259366 / 60000 ≈ 37.6561° — the benchmark "R&D" label rotation, flipped vertically.
		const shape = spWithXfrm('<a:xfrm rot="2259366" flipV="1"><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>')
		assert(Math.abs(shape.rotation - 37.6561) < 1e-3, `expected ≈37.6561°, got ${shape.rotation}`)
		assertEqual(shape.flipV, true, 'flipV="1" reads true')
		assertEqual(shape.flipH, false, 'no flipH reads false')
	})

	test('a present xfrm with no rot/flip reads rotation 0 and both flips false', () => {
		const shape = spWithXfrm('<a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>')
		assertEqual(shape.rotation, 0, 'a transform with no @rot is not rotated (0, not null)')
		assertEqual(shape.flipH, false, 'no @flipH reads false')
		assertEqual(shape.flipV, false, 'no @flipV reads false')
	})

	test('a shape with no own transform reads rotation null and flips false', () => {
		const shape = spWithXfrm('')
		assertEqual(shape.rotation, null, 'no a:xfrm → rotation null (inherits layout geometry)')
		assertEqual(shape.flipH, false, 'no transform → not flipped')
		assertEqual(shape.flipV, false, 'no transform → not flipped')
	})

	test('rot is faithful to the XML, not normalised to a signed range', () => {
		// 19216344 / 60000 = 320.2724° — a negative angle (≈ −39.73°) as PowerPoint stores it.
		const shape = spWithXfrm('<a:xfrm rot="19216344"><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>')
		assert(Math.abs(shape.rotation - 320.2724) < 1e-3, `expected ≈320.2724° (raw ÷60000), got ${shape.rotation}`)
	})

	test('rotation and flips read from genuine PowerPoint shapes (rotation-flip.pptx)', async () => {
		// De-circularised: was a write→read round-trip (addShape{rotate,flipH}→reopen);
		// now reads two desktop-PowerPoint-authored rectangles. PowerPoint stored
		// rot="2700000" (2700000 / 60000 = 45°) on rotated-45 and flipH="1" on flipped-h.
		const presentation = await open('rotation-flip')
		const shapes = presentation.slides[0].shapes
		const rotated = shapes.find((s) => s.name === 'rotated-45')
		const flipped = shapes.find((s) => s.name === 'flipped-h')
		assert(rotated, 'expected the rotated-45 rect')
		assert(flipped, 'expected the flipped-h rect')
		assert(Math.abs(rotated.rotation - 45) < EPS, `rotated-45 reads 45°, got ${rotated.rotation}`)
		assertEqual(rotated.flipH, false, 'rotated-45 is not horizontally flipped')
		assertEqual(rotated.flipV, false, 'rotated-45 is not vertically flipped')
		assertEqual(flipped.flipH, true, 'flipped-h reads flipH true')
		assertEqual(flipped.rotation, 0, 'flipped-h has no rotation (reads 0)')
		assertEqual(flipped.flipV, false, 'flipped-h is not vertically flipped')
	})
})
