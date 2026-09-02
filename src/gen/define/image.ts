/**
 * ts-pptx: Image Definition
 *
 * `addImageDefinition` resolves an `addImage()` source, allocates its drawing rel(s) (SVG needs
 * a second for the PNG fallback), registers the media bytes through `registerImageMediaRel` and
 * inherits any placeholder geometry. `registerImageFillMedia` does the same media-rel registration
 * for an image *fill* used by a shape or text box.
 */
import { SlideObjectType } from '../../enums.js'
import { warn } from '../../diagnostics.js'
import type { Coord, ImageProps, ObjectOptions, ShapeFillProps } from '../../types/index.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { encodeXmlAttrValue, getNewRelId, validateObjectName } from '../utils.js'
import { correctShadowOptions } from '../drawingml/effect.js'
import { svgMarkupToDataUri } from '../../media/base64.js'
import { imageExtensionForSource } from '../../media/content-type.js'
import { getImageSizeFromBase64 } from '../../media/image-size.js'
import { getSmartParseNumber } from '../../units-internal.js'
import { nextObjectNameIdx } from './object-name.js'
import { registerImageMediaRel, registerSvgImageRels } from './image-rel.js'
import { registerHyperlinkRel } from './hyperlinks.js'
import { InvalidOptionError } from '../../errors.js'
import { pickDefined } from '../../options-internal.js'

/** DPI PowerPoint assumes when sizing an inserted raster image (natural pixels / 96 == inches) */
const IMAGE_NATURAL_DPI = 96

/**
 * Register a raster image fill as a slide media relationship and stash the resolved
 * rId on the fill object so serialization can emit `<a:blipFill r:embed="rIdN">`.
 * Mirrors the non-SVG media-registration path used by `addImageDefinition()`,
 * including de-duplication of identical sources. SVG sources are not
 * supported as fills yet.
 * @param {PresSlideInternal} target - slide the owning object belongs to
 * @param {ShapeFillProps} fill - fill options carrying `image: { path | data }`
 */
export function registerImageFillMedia(target: PresSlideInternal, fill: ShapeFillProps): void {
	const strImagePath = fill.image?.path || ''
	const strImageData = fill.image?.data || ''

	if (!strImagePath && !strImageData) {
		warn('image-fill/missing-source', 'image fill requires `image.path` or `image.data`; ignoring image fill.')
		fill.type = 'none'
		return
	}
	if (strImageData && !strImageData.toLowerCase().includes('base64,')) {
		warn(
			'image-fill/missing-base64-header',
			"image fill `data` value lacks a base64 header (ex: 'image/png;base64,...'); ignoring image fill."
		)
		fill.type = 'none'
		return
	}

	const strImgExtn = imageExtensionForSource(strImagePath, strImageData)

	if (strImgExtn === 'svg') {
		warn(
			'image-fill/svg-unsupported',
			'SVG image fills are not supported; ignoring image fill. Use a raster format (PNG/JPEG/GIF/BMP/WebP).'
		)
		fill.type = 'none'
		return
	}

	const imageRelId = getNewRelId(target)
	registerImageMediaRel(target, { path: strImagePath, data: strImageData, extn: strImgExtn }, imageRelId)
	fill.type = 'image'
	fill._imgRid = imageRelId
}

/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {ImageProps} `opt` - object containing `path`/`data`, `x`, `y`, etc.
 * @param {PresSlideInternal} `target` - slide that the image should be added to (if not specified as the 2nd arg)
 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
 */
/**
 * Normalize an `addImage()` call into a slide image object and register its media relationships.
 *
 * Resolves the image source (path or base64 `data:`), allocates the drawing relationship id(s)
 * — SVG needs a second rId for its PNG rasterization fallback — appends the bytes to
 * `target._relsMedia` (deduping identical sources), and inherits any omitted x/y/w/h from a
 * matching layout picture placeholder. The resulting `SlideObject` is pushed onto the slide;
 * `gen/slide/objects/image.ts` later emits the `<p:pic>`.
 * @param target - slide (or master/group) the image is appended to
 * @param opt - the caller's `ImageProps` (path/data, geometry, sizing, hyperlink, placeholder, …)
 */
export function addImageDefinition(target: PresSlideInternal, opt: ImageProps): void {
	const newObject: SlideObject = {
		_type: SlideObjectType.image,
	}

	// Inherit geometry from a matching layout placeholder: an image targeting a
	// placeholder adopts that placeholder's position/size for any of x/y/w/h the caller omits.
	// Explicit `opt` values always win; this only fills the gaps so a picture placeholder no longer
	// collapses to the image's natural/1in fallback when no dimensions are supplied. Mirrors the
	// text-object placeholder inheritance in addTextDefinition().
	let phX: Coord | undefined
	let phY: Coord | undefined
	let phW: Coord | undefined
	let phH: Coord | undefined
	if (opt.placeholder && target._slideLayout?._slideObjects) {
		const placeHold = target._slideLayout._slideObjects.find(
			(item) => item._type === SlideObjectType.placeholder && item.options?.placeholder === opt.placeholder
		)
		if (placeHold?.options) {
			phX = placeHold.options.x
			phY = placeHold.options.y
			phW = placeHold.options.w
			phH = placeHold.options.h
		}
	}

	// FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
	const intPosX = opt.x ?? phX ?? 0
	const intPosY = opt.y ?? phY ?? 0
	const intWidth = opt.w ?? phW ?? 0
	const intHeight = opt.h ?? phH ?? 0
	const objHyperlink = opt.hyperlink || ''
	// Convenience: accept raw SVG markup via `svg` and encode it to a data URI.
	// `data`/`path` win when also supplied, matching the documented precedence.
	const strImageData = opt.data || (opt.svg && !opt.path ? svgMarkupToDataUri(opt.svg) : '')
	const strImagePath = opt.path || ''
	const imageRelId = getNewRelId(target)
	const imageNameIdx = nextObjectNameIdx(target, SlideObjectType.image)
	const objectName = opt.objectName
		? encodeXmlAttrValue(validateObjectName(opt.objectName, 'image'))
		: `Image ${imageNameIdx}`

	// REALITY-CHECK: an unusable source has nothing to degrade to — there is no image to place —
	// so these reject rather than warn. `addMedia()` already rejects the missing-source and
	// missing-header cases, as does the `hyperlink` check further down this same function.
	if (!strImagePath && !strImageData) {
		throw new InvalidOptionError('image/missing-source', "addImage(): either 'data' or 'path' is required")
	} else if (strImagePath && typeof strImagePath !== 'string') {
		throw new InvalidOptionError(
			'image/path-not-a-string',
			`addImage(): 'path' should be a string, ex: {path:'/img/sample.png'} - you sent ${String(strImagePath)}`
		)
	} else if (strImageData && typeof strImageData !== 'string') {
		throw new InvalidOptionError(
			'image/data-not-a-string',
			`addImage(): 'data' should be a string, ex: {data:'image/png;base64,NMP[...]'} - you sent ${String(strImageData)}`
		)
	} else if (strImageData && typeof strImageData === 'string' && !strImageData.toLowerCase().includes('base64,')) {
		throw new InvalidOptionError(
			'image/missing-base64-header',
			"addImage(): `data` value lacks a base64 header, ex: 'image/png;base64,NMP[...]'"
		)
	}

	// STEP 1: Set extension (the `data:` mime wins over the path when both are supplied)
	const strImgExtn = imageExtensionForSource(strImagePath, strImageData)

	// STEP 2: Set type/path
	newObject._type = SlideObjectType.image
	newObject.image = strImagePath || 'preencoded.png'

	// STEP 3: Default any missing dimension from the image's intrinsic (natural) size.
	// For base64 `data` images the bytes are already in hand, so we can read the
	// natural pixel size synchronously and avoid the legacy 1x1 fallback that
	// squished data-only images into a 1in square.
	// Path images can't be measured synchronously (bytes load async during export),
	// so the missing extent is flagged via `_szAuto` and backfilled at serialize time
	// once the media bytes are available.
	// PowerPoint inserts images at 96 DPI, so natural pixels / 96 == inches.
	//
	// A vector source is measured for its RATIO only. An SVG's user units are dependable
	// relative to each other and merely conventional in absolute terms — a 24-unit icon is
	// authored to be drawn at whatever size it is placed, not as a quarter-inch object — so an
	// SVG with one side given derives the other, and an SVG with neither keeps the 1in fallback.
	let defWidth = intWidth
	let defHeight = intHeight
	let szAuto: { w: boolean; h: boolean } | undefined
	const vectorWithNoExtent = strImgExtn === 'svg' && !intWidth && !intHeight
	if ((!intWidth || !intHeight) && !vectorWithNoExtent) {
		const natural = strImageData ? getImageSizeFromBase64(strImageData) : null
		if (natural) {
			if (!intWidth && !intHeight) {
				// Neither given: use the natural size (inches @ 96 DPI)
				defWidth = natural.w / IMAGE_NATURAL_DPI
				defHeight = natural.h / IMAGE_NATURAL_DPI
			} else if (typeof intWidth === 'number' && intWidth && !intHeight) {
				// Only width given: preserve aspect ratio for height (same unit as width)
				defHeight = intWidth * (natural.h / natural.w)
			} else if (typeof intHeight === 'number' && intHeight && !intWidth) {
				// Only height given: preserve aspect ratio for width (same unit as height)
				defWidth = intHeight * (natural.w / natural.h)
			}
		} else if (strImagePath) {
			// Path image: defer measurement to serialize time. Record which side(s) to derive
			// from the natural ratio; the 1in fallback below still applies if it stays unmeasurable.
			szAuto = { w: !intWidth, h: !intHeight }
		}
	}

	// STEP 4: Set image properties & options
	const shadow = correctShadowOptions(opt.shadow)
	// The twelve pass-through options go through `pickDefined` rather than being listed as
	// `key: opt.key`: an image that asks for no `crop` carries no `crop` key, where the literal
	// wrote one holding `undefined`. Every one of these is read by an emitter that branches on the
	// option being *there* (`if (opts.crop)`, `if (opts.duotone)`), so they were two spellings of
	// the same request — and `ObjectOptions` is spread onto a placeholder's options besides.
	const objectOptions: ObjectOptions = {
		x: intPosX || 0,
		y: intPosY || 0,
		w: defWidth || 1,
		h: defHeight || 1,
		altText: opt.altText || '',
		rounding: typeof opt.rounding === 'boolean' ? opt.rounding : false,
		...pickDefined(opt, [
			'shape',
			'points',
			'rectRadius',
			'shapeAdjust',
			'sizing',
			'crop',
			'placeholder',
			'duotone',
			'grayscale',
			'biLevel',
			'clrChange',
			'objectLock',
		]),
		rotate: opt.rotate || 0,
		flipV: opt.flipV || false,
		flipH: opt.flipH || false,
		transparency: opt.transparency || 0,
		objectName,
		...(shadow ? { shadow } : {}),
		...(szAuto ? { _szAuto: szAuto } : {}),
	}
	newObject.options = objectOptions

	// STEP 5: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	if (strImgExtn === 'svg') {
		// An SVG consumes *TWO* rels — the PNG fallback and the SVG itself — allocated and
		// pushed by `registerSvgImageRels`, which the picture-bullet definer shares.
		newObject.imageRid = registerSvgImageRels(target, {
			path: strImagePath ?? '',
			data: strImageData ?? '',
			svgSize: {
				w: getSmartParseNumber(objectOptions.w, 'X', target._presLayout),
				h: getSmartParseNumber(objectOptions.h, 'Y', target._presLayout),
			},
		}).svgRid
	} else {
		registerImageMediaRel(target, { path: strImagePath, data: strImageData, extn: strImgExtn }, imageRelId)
		newObject.imageRid = imageRelId
	}

	// STEP 6: Hyperlink support
	if (typeof objHyperlink === 'object') {
		if (!objHyperlink.url && !objHyperlink.slide)
			throw new InvalidOptionError('hyperlink/missing-target', 'addImage: `hyperlink` requires either `url` or `slide`')
		else {
			registerHyperlinkRel(target, objHyperlink)
			newObject.hyperlink = objHyperlink
		}
	}

	// STEP 7: Add object to slide
	target._slideObjects.push(newObject)
}
