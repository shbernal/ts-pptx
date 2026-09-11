/**
 * ts-pptx: image slide-object serialization
 *
 * Emits an `image` slide object as a `<p:pic>`: the blip and its image effects (transparency,
 * duotone, colour change, grayscale, bi-level), the crop or `sizing` source rectangle, the
 * clipping geometry, outline and shadow. An omitted dimension is backfilled from the embedded
 * bytes' natural pixel ratio, which only becomes available once `_relsMedia` is populated.
 */

import { DEF_TEXT_SHADOW } from '../../../constants-internal.js'
import { createColorElement } from '../../drawingml/color.js'
import { createShadowEffectLst } from '../../drawingml/effect.js'
import { genXmlCustGeom, genXmlPresetGeom } from '../../drawingml/geometry.js'
import { genXmlImageCrop, genXmlVectorAspectFit, ImageSizingXml } from '../../drawingml/image.js'
import { genXmlObjectLock, PICTURE_LOCK_ATTRS } from '../../drawingml/locks.js'
import { genXmlPlaceholder } from '../../drawingml/text-body.js'
import { el, raw, voidEl, type XmlChild } from '../../oxml/el.js'
import { OOXML_NS } from '../../../ooxml/namespaces.js'
import { fractionToFixedPercent, getSmartParseNumber, transparencyToAlpha } from '../../../units-internal.js'
import { warn } from '../../../diagnostics.js'
import { cNvPrHyperlink, cNvPrOpen, genXmlShapeLine, type RenderContext, xfrmEl } from './shared.js'
import { imageNaturalSize, resolveImageExtent } from './image-extent.js'

/**
 * Render an `image` slide object to its `<p:pic>` XML (sizing/crop, rounding, hyperlink, shadow).
 */
export function renderImageObject(ctx: RenderContext): string {
	const { obj: slideItemObj, shapeId, slide, placeholder: placeholderObj, locationAttrs, itemOpts } = ctx
	const { x, y } = ctx.frame
	// The frame, as two pairs that diverge: `_szAuto` backfills an omitted dimension from the
	// image's natural ratio, and `sizing` splits the two apart for good. `frameW`/`frameH` stay the
	// box the caller asked for, which is what the source rectangle is cropped against, while
	// `drawnW`/`drawnH` become the box actually drawn and emitted. Neither pair is the image's own
	// size; that is the third quantity here, and only `naturalSize()` answers it.
	// `resolveImageExtent` computes both pairs, because a parent group's bounds need the drawn box
	// too. `drawnW`/`drawnH` start as the frame and take the `sizing` box where the source rectangle
	// is emitted below.
	const { frameW, frameH, drawnW: sizedW, drawnH: sizedH } = resolveImageExtent(slideItemObj, slide, ctx.frame)
	let drawnW = frameW
	let drawnH = frameH
	const { sizing, rounding } = itemOpts
	let strSlideXml = ''
	// `itemOpts` is the caller's already-normalized `itemOpts` (see the dispatch in
	// `slideObjectToXml`). Read it rather than re-narrowing the field: this function has exactly
	// one call site, and a contract stated there beats a defensive re-assignment here.
	// The media bytes this picture points at. Not available synchronously in `addImage()`, but
	// populated by now — which is why every question that needs the image itself (its natural
	// size, whether it is a vector) is answered here rather than at definition time.
	const mediaRel = (slide._relsMedia || []).find((rel) => rel.rId === slideItemObj.imageRid)
	const naturalSize = (): { w: number; h: number } | null => imageNaturalSize(slideItemObj, slide)
	const imgOpts = itemOpts
	const imgLink = slideItemObj.hyperlink
	strSlideXml += '<p:pic>'
	strSlideXml += '  <p:nvPicPr>'
	strSlideXml +=
		cNvPrOpen(shapeId, imgOpts.objectName, imgOpts.altText || slideItemObj.image || '') + '>' + cNvPrHyperlink(imgLink)
	strSlideXml += '    </p:cNvPr>'
	// Default to locking aspect ratio (PowerPoint's own behavior); user `objectLock` overrides any flag, incl. noChangeAspect.
	strSlideXml += el(
		'p:cNvPicPr',
		null,
		raw(
			genXmlObjectLock(
				'a:picLocks',
				PICTURE_LOCK_ATTRS,
				{ noChangeAspect: true, ...imgOpts.objectLock },
				imgOpts.objectName
			)
		),
		{ openPrefix: '    ' }
	)
	strSlideXml += el('p:nvPr', null, raw(genXmlPlaceholder(placeholderObj)), { openPrefix: '    ' })
	strSlideXml += '  </p:nvPicPr>'

	// The `<a:blip>` image-effect children (CT_Blip), shared by the SVG and raster branches:
	// transparency as `alphaModFix`, then duotone recolor (shadows→shadow, highlights→highlight),
	// both before any `extLst`.
	// NOTE: the SVG branch writes ` <a:alphaModFix` with a LEADING space and the raster branch
	// writes none. That space is byte-significant, so it is passed in rather than normalized away.
	const blipEffects = (alphaPrefix: string): XmlChild[] => [
		imgOpts.transparency
			? raw(voidEl('a:alphaModFix', { amt: transparencyToAlpha(imgOpts.transparency) }, { openPrefix: alphaPrefix }))
			: null,
		imgOpts.duotone
			? raw(
					el('a:duotone', null, [
						raw(createColorElement(imgOpts.duotone.shadow)),
						raw(createColorElement(imgOpts.duotone.highlight)),
					])
				)
			: null,
		imgOpts.clrChange
			? raw(
					el('a:clrChange', null, [
						raw(el('a:clrFrom', null, raw(createColorElement(imgOpts.clrChange.from)))),
						raw(el('a:clrTo', null, raw(createColorElement(imgOpts.clrChange.to)))),
					])
				)
			: null,
		imgOpts.grayscale ? raw(voidEl('a:grayscl')) : null,
		imgOpts.biLevel
			? raw(
					voidEl('a:biLevel', {
						thresh: fractionToFixedPercent(
							imgOpts.biLevel.threshold,
							'image/bilevel-threshold-out-of-range',
							'biLevel.threshold'
						),
					})
				)
			: null,
	]

	strSlideXml += '<p:blipFill>'
	// NOTE: This works for both cases: either `path` or `data` contains the SVG
	const isVector = mediaRel?.extn === 'svg'
	if (isVector) {
		strSlideXml += el('a:blip', { 'r:embed': `rId${(slideItemObj.imageRid ?? 0) - 1}` }, [
			...blipEffects(' '),
			raw(
				el(
					'a:extLst',
					null,
					raw(
						el(
							'a:ext',
							{ uri: '{96DAC541-7B7A-43D3-8B79-37D633B846F1}' },
							raw(
								voidEl(
									'asvg:svgBlip',
									{
										'xmlns:asvg': OOXML_NS.asvg,
										'r:embed': `rId${slideItemObj.imageRid}`,
									},
									{ openPrefix: '   ' }
								)
							),
							{ openPrefix: '  ', closePrefix: '  ' }
						)
					),
					{ openPrefix: ' ', closePrefix: ' ' }
				)
			),
		])
	} else {
		strSlideXml += el('a:blip', { 'r:embed': `rId${slideItemObj.imageRid}` }, blipEffects(''))
	}
	if (itemOpts.crop) {
		// Explicit OOXML srcRect (percentage edge insets), emitted verbatim. Crops the source
		// directly, so it wins over the inch-based `sizing` crop and works for SVG/unmeasurable
		// formats; the picture's normal w/h box stays the display extent.
		if (sizing?.type)
			warn(
				'image/crop-and-sizing-conflict',
				`addImage 'crop' and 'sizing' are mutually exclusive for image "${itemOpts.objectName}"; 'sizing' was ignored.`
			)
		strSlideXml += genXmlImageCrop(itemOpts.crop, itemOpts.objectName)
	} else if (sizing?.type) {
		const boxW = sizedW
		const boxH = sizedH
		const boxX = getSmartParseNumber(sizing.x || 0, 'X', slide._presLayout)
		const boxY = getSmartParseNumber(sizing.y || 0, 'Y', slide._presLayout)

		// `cover`/`contain` crop the *source* bitmap, so the srcRect must be derived from the
		// image's natural pixel ratio — not the displayed box (options.w/h). Measure it from the
		// embedded media bytes; if unmeasurable (SVG/unknown format) fall back to display dims + warn.
		// `crop` keeps display EMU: its contract treats the displayed extent as the crop frame.
		let cropSize: { w: number; h: number } = { w: drawnW, h: drawnH }
		if (sizing.type === 'cover' || sizing.type === 'contain') {
			const natural = naturalSize()
			if (natural) {
				cropSize = natural
			} else {
				warn(
					'image/unmeasurable-natural-size',
					`sizing '${sizing.type}' could not measure natural dimensions for image "${itemOpts.objectName}"; falling back to displayed aspect ratio (crop may be inexact). Provide a raster image (PNG/JPEG/GIF/BMP/WebP) or an SVG with width/height or a viewBox to enable an aspect-correct crop.`
				)
			}
		}

		strSlideXml += ImageSizingXml[sizing.type](cropSize, { w: boxW, h: boxH, x: boxX, y: boxY })
		drawnW = boxW
		drawnH = boxH
	} else {
		// No `sizing` at all. A raster fills its box — that box was chosen for it, and PowerPoint
		// does the same. A vector does not: an SVG states its own aspect ratio in a viewBox, and a
		// glyph squashed to a box that disagrees is a defect, not a layout choice. So a measurable
		// vector letterboxes here by default, and `sizing: { type: 'stretch' }` opts back out.
		// Unmeasurable vectors (no viewBox, no width/height) fall through silently: nothing was
		// asked for, so there is nothing to warn about.
		const natural = isVector ? naturalSize() : null
		const aspectFit = natural ? genXmlVectorAspectFit(natural, { w: drawnW, h: drawnH }) : null
		strSlideXml += aspectFit ?? el('a:stretch', null, raw(voidEl('a:fillRect')), { openPrefix: '  ' })
	}
	strSlideXml += '</p:blipFill>'
	strSlideXml += '<p:spPr>'
	strSlideXml += xfrmEl('a:xfrm', { x, y, cx: drawnW, cy: drawnH }, locationAttrs, {
		openPrefix: ' ',
		childPrefix: '  ',
		closePrefix: ' ',
	})
	// Clip the picture to a geometry. `points` (freeform custGeom) takes precedence over `shape`/`rounding`;
	// otherwise `shape` wins over `rounding` (shorthand for an ellipse), falling back to a plain rectangle.
	if (itemOpts.points) {
		strSlideXml += ' ' + genXmlCustGeom(itemOpts, drawnW, drawnH, slide._presLayout)
	} else {
		strSlideXml += ' ' + genXmlPresetGeom(itemOpts.shape ?? (rounding ? 'ellipse' : 'rect'), itemOpts, drawnW, drawnH)
	}

	// BORDER: `<a:ln>` outline (must precede `<a:effectLst>` per CT_ShapeProperties order)
	if (itemOpts.line) strSlideXml += genXmlShapeLine(itemOpts.line)

	// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
	if (itemOpts.shadow && itemOpts.shadow.type !== 'none') {
		strSlideXml += createShadowEffectLst(itemOpts.shadow, DEF_TEXT_SHADOW)
	}
	strSlideXml += '</p:spPr>'
	strSlideXml += '</p:pic>'
	return strSlideXml
}
