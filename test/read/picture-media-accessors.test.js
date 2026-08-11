// Read-model coverage for the picture-media and accessibility accessors:
// Picture.mediaKind / mediaPartName / crop and Shape.description / title /
// isDecorative. These power an accessible-export or audit tool that must be able
// to tell an SVG-only icon from an empty picture, read an existing crop, and
// surface alt text + the "mark as decorative" flag.
//
// picture-media.pptx is a minimal desktop PowerPoint fixture authored for this:
//   SvgPic     — an inserted .svg (svg-only: a:blip has no r:embed, only asvg:svgBlip)
//   CroppedPic — a raster PNG with a four-edge a:srcRect crop and an alt-text descr
//   DecoRect   — a rectangle flagged decorative (adec:decorative val="1")
//   DescRect   — a rectangle with a plain alt-text descr, not decorative
// The 'both' (raster+SVG) and plain 'raster' mediaKind cases live in
// style-accessors.test.js against image.pptx.

import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { Picture } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

/** Flatten a shape list, descending into groups. */
function allShapes(shapes) {
	return shapes.flatMap((shape) => (shape.shapeType === 'group' ? [shape, ...allShapes(shape.shapes)] : [shape]))
}

function named(slide, name) {
	const shape = allShapes(slide.shapes).find((s) => s.name === name)
	assert(shape, `expected a shape named ${name}`)
	return shape
}

/** Wrap a standalone `<p:pic>…</p:pic>` string in a Picture with a stand-in slide. */
function pictureFromXml(innerXml) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}">${innerXml}</p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'pic')[0]
	// Stand-in slide: none of the accessors under test reach through to it.
	return new Picture(el, /** @type {any} */ ({}))
}

describe('Picture media accessors (picture-media.pptx)', () => {
	test('mediaKind reports svg for an SVG-only picture and raster for a plain one', async () => {
		const slide = (await openFixture('picture-media')).slides[0]
		assertEqual(named(slide, 'SvgPic').mediaKind, 'svg', 'a blip with only asvg:svgBlip is svg-only')
		assertEqual(named(slide, 'CroppedPic').mediaKind, 'raster', 'a blip with only r:embed is raster')
	})

	test('mediaPartName falls back to the SVG part when there is no raster', async () => {
		const slide = (await openFixture('picture-media')).slides[0]
		const svgPic = named(slide, 'SvgPic')
		assertEqual(svgPic.imagePartName, null, 'an SVG-only picture has no raster part')
		const part = svgPic.mediaPartName
		assert(part && part.endsWith('.svg'), `mediaPartName resolves to the .svg part; got ${part}`)
		assertEqual(part, svgPic.svgPartName, 'mediaPartName equals svgPartName when raster is absent')
	})

	test('crop reads a:srcRect as per-edge fractions', async () => {
		const slide = (await openFixture('picture-media')).slides[0]
		const crop = named(slide, 'CroppedPic').crop
		assert(crop, 'the cropped picture reports a crop')
		// srcRect l="41666" t="27778" r="20833" b="13889" (thousandths of a percent).
		assertEqual(crop.left, 41666 / 100000, 'left edge fraction')
		assertEqual(crop.top, 27778 / 100000, 'top edge fraction')
		assertEqual(crop.right, 20833 / 100000, 'right edge fraction')
		assertEqual(crop.bottom, 13889 / 100000, 'bottom edge fraction')
	})

	test('crop is null when the picture has no a:srcRect', async () => {
		const slide = (await openFixture('picture-media')).slides[0]
		assertEqual(named(slide, 'SvgPic').crop, null, 'an uncropped picture reports null')
	})
})

describe('Shape accessibility accessors (picture-media.pptx)', () => {
	test('description reads p:cNvPr/@descr and null when unset', async () => {
		const slide = (await openFixture('picture-media')).slides[0]
		assertEqual(named(slide, 'CroppedPic').description, 'A cropped stopwatch photo', 'alt text on a picture')
		assertEqual(named(slide, 'DescRect').description, 'A described rectangle', 'alt text on an autoshape')
		assertEqual(named(slide, 'DecoRect').description, null, 'a decorative shape has no description')
	})

	test('description is settable and clearable', async () => {
		const slide = (await openFixture('picture-media')).slides[0]
		const shape = named(slide, 'DescRect')
		shape.description = 'Reworded alt text'
		assertEqual(shape.description, 'Reworded alt text', 'the setter updates @descr')
		shape.description = ''
		assertEqual(shape.description, null, 'setting empty clears @descr back to null')
	})

	test('isDecorative reflects the adec:decorative extension', async () => {
		const slide = (await openFixture('picture-media')).slides[0]
		assertEqual(named(slide, 'DecoRect').isDecorative, true, 'the marked shape reads decorative')
		assertEqual(named(slide, 'DescRect').isDecorative, false, 'a described shape is not decorative')
		assertEqual(named(slide, 'CroppedPic').isDecorative, false, 'a picture with alt text is not decorative')
	})

	test('title is null when no @title is present (modern PowerPoint omits it)', async () => {
		const slide = (await openFixture('picture-media')).slides[0]
		for (const name of ['SvgPic', 'CroppedPic', 'DecoRect', 'DescRect']) {
			assertEqual(named(slide, name).title, null, `${name} has no @title`)
		}
	})
})

describe('Picture media accessors — edge cases (off-fixture)', () => {
	test('mediaKind is none for a p:pic with no embedded blip', () => {
		const pic = pictureFromXml('<p:pic><p:blipFill><a:stretch/></p:blipFill></p:pic>')
		assertEqual(pic.mediaKind, 'none', 'no raster and no SVG')
		assertEqual(pic.mediaPartName, null, 'nothing to resolve')
	})

	test('an explicit empty a:srcRect reads as a zero crop, not null', () => {
		const pic = pictureFromXml('<p:pic><p:blipFill><a:srcRect/><a:stretch/></p:blipFill></p:pic>')
		const crop = pic.crop
		assert(crop, 'presence of srcRect is meaningful — not null')
		assertEqual(crop.left, 0, 'left defaults to 0')
		assertEqual(crop.top, 0, 'top defaults to 0')
		assertEqual(crop.right, 0, 'right defaults to 0')
		assertEqual(crop.bottom, 0, 'bottom defaults to 0')
	})

	test('a partial a:srcRect defaults the missing edges to zero', () => {
		const pic = pictureFromXml('<p:pic><p:blipFill><a:srcRect l="10000"/><a:stretch/></p:blipFill></p:pic>')
		const crop = pic.crop
		assert(crop, 'crop present')
		assertEqual(crop.left, 0.1, 'the declared edge is non-zero')
		assertEqual(crop.top, 0, 'missing top defaults to 0')
		assertEqual(crop.right, 0, 'missing right defaults to 0')
		assertEqual(crop.bottom, 0, 'missing bottom defaults to 0')
	})
})
