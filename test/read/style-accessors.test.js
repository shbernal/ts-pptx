// Read-model coverage for the per-shape / per-paragraph STYLE accessors:
// Shape.lineWidthPt / adjustValues / gradientStops / hidden, and
// Paragraph.align / spaceBeforePt / spaceAfterPt / marginLeftPt / indentPt /
// bullet. These power a faithful style dump of a source slide (see
// downstream's bundle `style.json`), so the reads must hold against both real
// PowerPoint-authored XML and our own serializer.
//
// Strategy:
// - mixed.pptx is genuine Office output and carries line widths, alignment,
//   spacing, indents, and bullet glyphs — assert the paragraph/line reads there.
// - geometry adjusts and gradient stops are not in the vendored fixtures, so we
//   round-trip them through the write API (generate → reopen → read).

import { readFile } from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import PptxGenJS from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'

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
})

describe('Shape style reads — write→read round-trip', () => {
	async function reopen(buildFn) {
		const pres = new PptxGenJS()
		buildFn(pres)
		const buf = await pres.stream()
		return Presentation.load(buf)
	}

	test('adjustValues exposes a roundRect rectRadius as the avLst adj handle', async () => {
		const presentation = await reopen((p) => {
			const slide = p.addSlide()
			slide.addShape(p.shapes.ROUNDED_RECTANGLE, { x: 1, y: 1, w: 3, h: 1, fill: { color: 'CCCCCC' }, rectRadius: 0.1 })
		})
		const shape = presentation.slides[0].shapes.find((s) => s.presetGeometry === 'roundRect')
		assert(shape, 'expected the roundRect')
		const adj = shape.adjustValues
		assert('adj' in adj, `expected an 'adj' handle; got ${JSON.stringify(adj)}`)
		assert(adj.adj.startsWith('val '), `expected a 'val N' formula; got ${JSON.stringify(adj.adj)}`)
		// A plain rect has no adjust handles.
		const presRect = await reopen((p) => {
			p.addSlide().addShape(p.shapes.RECTANGLE, { x: 1, y: 1, w: 3, h: 1, fill: { color: 'CCCCCC' } })
		})
		const rect = presRect.slides[0].shapes.find((s) => s.presetGeometry === 'rect')
		assertEqual(Object.keys(rect.adjustValues).length, 0, 'a plain rect has no adjust handles')
	})

	test('lineWidthPt round-trips an explicit line width', async () => {
		const presentation = await reopen((p) => {
			p.addSlide().addShape(p.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 3,
				h: 1,
				fill: { color: 'CCCCCC' },
				line: { color: '111111', width: 2 },
			})
		})
		const shape = presentation.slides[0].shapes.find((s) => s.presetGeometry === 'rect')
		assertEqual(shape.lineWidthPt, 2, 'line width 2pt round-trips')
	})

	test('gradientStops reads gsLst stops with position + colour split, null when solid', async () => {
		const presentation = await reopen((p) => {
			const slide = p.addSlide()
			slide.addShape(p.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 3,
				h: 1,
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 90,
						stops: [
							{ position: 0, color: '#451DC7' },
							{ position: 100, color: 'accent1' },
						],
					},
				},
			})
			slide.addShape(p.shapes.RECTANGLE, { x: 1, y: 3, w: 3, h: 1, fill: { color: '00FF00' } })
		})
		const shapes = presentation.slides[0].shapes.filter((s) => s.presetGeometry === 'rect')
		const gradient = shapes.find((s) => s.gradientStops !== null)
		assert(gradient, 'expected a gradient-filled rect')
		const stops = gradient.gradientStops
		assertEqual(stops.length, 2, 'two gradient stops')
		const first = stops.find((s) => s.position === 0)
		const last = stops.find((s) => s.position === 1)
		assert(first && last, `stops at 0 and 1; got positions ${JSON.stringify(stops.map((s) => s.position))}`)
		assertEqual(first.color, '451DC7', 'first stop is an explicit colour')
		assertEqual(first.schemeColor, null, 'first stop has no scheme colour')
		assertEqual(last.color, null, 'second stop has no explicit colour')
		assertEqual(last.schemeColor, 'accent1', 'second stop is a scheme colour')

		const solid = shapes.find((s) => s.fillColor === '00FF00')
		assertEqual(solid.gradientStops, null, 'a solid-filled shape reports null gradientStops')
	})
})

describe('Theme colour resolution — write→read round-trip', () => {
	// A generated deck carries the default Office theme + a slide → layout → master
	// → theme chain, so a scheme token resolves deterministically: accent1=4472C4,
	// accent2=ED7D31, accent3=A5A5A5 (see THEME_CLR_SCHEME_DEFAULTS in gen-xml).
	async function reopen(buildFn) {
		const pres = new PptxGenJS()
		buildFn(pres)
		const buf = await pres.stream()
		return Presentation.load(buf)
	}

	test('resolvedFill resolves a scheme fill to the theme hex, and an explicit fill to itself', async () => {
		const presentation = await reopen((p) => {
			const slide = p.addSlide()
			slide.addShape(p.shapes.RECTANGLE, { x: 1, y: 1, w: 3, h: 1, fill: { color: 'accent1' } })
			slide.addShape(p.shapes.RECTANGLE, { x: 1, y: 3, w: 3, h: 1, fill: { color: 'FF0000' } })
		})
		const shapes = presentation.slides[0].shapes.filter((s) => s.presetGeometry === 'rect')
		const scheme = shapes.find((s) => s.fillSchemeColor === 'accent1')
		assert(scheme, 'expected the accent1-filled rect')
		assertEqual(scheme.resolvedFill.hex, '4472C4', 'accent1 resolves to the theme accent1 hex')
		// The raw read still reports the unresolved token — resolution is opt-in.
		assertEqual(scheme.fillColor, null, 'fillColor still reports null for a scheme-coloured fill')

		const explicit = shapes.find((s) => s.fillColor === 'FF0000')
		assertEqual(explicit.resolvedFill.hex, 'FF0000', 'an explicit srgb fill resolves to itself')
	})

	test('resolvedLine resolves a scheme line colour; null when there is no solid fill', async () => {
		const presentation = await reopen((p) => {
			const slide = p.addSlide()
			slide.addShape(p.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 3,
				h: 1,
				fill: { color: 'FF0000' },
				line: { color: 'accent2', width: 1 },
			})
			slide.addShape(p.shapes.RECTANGLE, {
				x: 1,
				y: 3,
				w: 3,
				h: 1,
				fill: {
					type: 'gradient',
					gradient: {
						kind: 'linear',
						angle: 0,
						stops: [
							{ position: 0, color: 'accent1' },
							{ position: 100, color: 'accent2' },
						],
					},
				},
			})
		})
		const shapes = presentation.slides[0].shapes.filter((s) => s.presetGeometry === 'rect')
		const lined = shapes.find((s) => s.lineSchemeColor === 'accent2')
		assert(lined, 'expected the accent2-lined rect')
		assertEqual(lined.resolvedLine.hex, 'ED7D31', 'accent2 line resolves to the theme accent2 hex')
		// A gradient-filled, unlined shape resolves neither a fill nor a line colour.
		const gradient = shapes.find((s) => s.gradientStops !== null)
		assertEqual(gradient.resolvedFill, null, 'a gradient fill has no a:solidFill to resolve')
		assertEqual(gradient.resolvedLine, null, 'an unlined shape resolves no line colour')
	})

	test('resolvedFill reports colour transforms without applying them', async () => {
		const presentation = await reopen((p) => {
			p.addSlide().addShape(p.shapes.RECTANGLE, {
				x: 1,
				y: 1,
				w: 3,
				h: 1,
				fill: { color: 'accent1', transparency: 25 },
			})
		})
		const shape = presentation.slides[0].shapes.find((s) => s.presetGeometry === 'rect')
		// transparency 25 → <a:alpha val="75000"/>; the base hex stays the theme colour.
		assertEqual(shape.resolvedFill.hex, '4472C4', 'base colour stays the theme hex, not premultiplied')
		assertEqual(shape.resolvedFill.transforms.length, 1, 'one transform child reported')
		assertEqual(shape.resolvedFill.transforms[0].name, 'alpha', 'the alpha transform is reported by name')
		assertEqual(shape.resolvedFill.transforms[0].value, '75000', 'with its raw @val, left for the consumer to apply')
	})

	test('Run.resolvedColor resolves a scheme run colour to the theme hex', async () => {
		const presentation = await reopen((p) => {
			p.addSlide().addText([{ text: 'hi', options: { color: 'accent3' } }], { x: 1, y: 1, w: 3, h: 1 })
		})
		const shape = presentation.slides[0].shapes.find((s) => s.hasTextFrame)
		const run = shape.textFrame.paragraphs[0].runs[0]
		assertEqual(run.schemeColor, 'accent3', 'the raw read reports the scheme token')
		assertEqual(run.resolvedColor.hex, 'A5A5A5', 'accent3 resolves to the theme accent3 hex')
	})
})
