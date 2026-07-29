/**
 * ts-pptx: Image Definition
 *
 * `addImageDefinition` resolves an `addImage()` source, allocates its drawing rel(s) (SVG needs
 * a second for the PNG fallback), registers the media bytes (deduped) and inherits any
 * placeholder geometry. `registerImageFillMedia` does the same media-rel registration for an
 * image *fill* used by a shape or text box.
 */
import { SlideObjectType } from '../../core-enums.js'
import { warn } from '../../log.js'
import type { Coord, ImageProps, ObjectOptions, ShapeFillProps } from '../../core-interfaces.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { encodeXmlAttrValue, getNewRelId, validateObjectName } from '../../gen-utils.js'
import { correctShadowOptions } from '../drawingml/effect.js'
import { svgMarkupToDataUri } from '../../media/base64.js'
import { imageContentType, imageExtensionForSource } from '../../media/content-type.js'
import { getImageSizeFromBase64 } from '../../media/image-size.js'
import { getSmartParseNumber } from '../../units-internal.js'
import { nextObjectNameIdx } from './object-name.js'

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
		warn('image fill requires `image.path` or `image.data`; ignoring image fill.')
		fill.type = 'none'
		return
	}
	if (strImageData && !strImageData.toLowerCase().includes('base64,')) {
		warn("Warning: image fill `data` value lacks a base64 header (ex: 'image/png;base64,...'); ignoring image fill.")
		fill.type = 'none'
		return
	}

	const strImgExtn = imageExtensionForSource(strImagePath, strImageData)

	if (strImgExtn === 'svg') {
		warn('SVG image fills are not supported; ignoring image fill. Use a raster format (PNG/JPEG/GIF/BMP/WebP).')
		fill.type = 'none'
		return
	}

	const imageRelId = getNewRelId(target)
	const mediaSlideKey =
		target._slideNum == null ? 'sm' : target._slideNum >= 1000 ? `sl-${target._slideNum}` : target._slideNum
	const imgContentType = imageContentType(strImgExtn)
	const dupeItem = target._relsMedia.find((item) => {
		if (item.isDuplicate || !item.Target || item.type !== imgContentType) return false
		return strImagePath ? item.path === strImagePath : !!strImageData && item.data === strImageData
	})

	target._relsMedia.push({
		path: strImagePath || 'preencoded.' + strImgExtn,
		type: imgContentType,
		extn: strImgExtn,
		data: strImageData || '',
		rId: imageRelId,
		isDuplicate: !!dupeItem?.Target,
		Target: dupeItem?.Target
			? dupeItem.Target
			: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${strImgExtn}`,
	})
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
	const sizing = opt.sizing
	const objHyperlink = opt.hyperlink || ''
	// Convenience: accept raw SVG markup via `svg` and encode it to a data URI.
	// `data`/`path` win when also supplied, matching the documented precedence.
	const strImageData = opt.data || (opt.svg && !opt.path ? svgMarkupToDataUri(opt.svg) : '')
	const strImagePath = opt.path || ''
	let imageRelId = getNewRelId(target)
	const imageNameIdx = nextObjectNameIdx(target, SlideObjectType.image)
	const objectName = opt.objectName
		? encodeXmlAttrValue(validateObjectName(opt.objectName, 'image'))
		: `Image ${imageNameIdx}`

	// REALITY-CHECK:
	if (!strImagePath && !strImageData) {
		console.error("ERROR: addImage() requires either 'data' or 'path' parameter!")
		return
	} else if (strImagePath && typeof strImagePath !== 'string') {
		console.error(
			`ERROR: addImage() 'path' should be a string, ex: {path:'/img/sample.png'} - you sent ${String(strImagePath)}`
		)
		return
	} else if (strImageData && typeof strImageData !== 'string') {
		console.error(
			`ERROR: addImage() 'data' should be a string, ex: {data:'image/png;base64,NMP[...]'} - you sent ${String(strImageData)}`
		)
		return
	} else if (strImageData && typeof strImageData === 'string' && !strImageData.toLowerCase().includes('base64,')) {
		console.error("ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')")
		return
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
	let defWidth = intWidth
	let defHeight = intHeight
	let szAuto: { w: boolean; h: boolean } | undefined
	if ((!intWidth || !intHeight) && strImgExtn !== 'svg') {
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
	const objectOptions: ObjectOptions = {
		x: intPosX || 0,
		y: intPosY || 0,
		w: defWidth || 1,
		h: defHeight || 1,
		altText: opt.altText || '',
		rounding: typeof opt.rounding === 'boolean' ? opt.rounding : false,
		shape: opt.shape,
		points: opt.points,
		rectRadius: opt.rectRadius,
		shapeAdjust: opt.shapeAdjust,
		sizing,
		crop: opt.crop,
		placeholder: opt.placeholder,
		rotate: opt.rotate || 0,
		flipV: opt.flipV || false,
		flipH: opt.flipH || false,
		transparency: opt.transparency || 0,
		duotone: opt.duotone,
		grayscale: opt.grayscale,
		biLevel: opt.biLevel,
		clrChange: opt.clrChange,
		objectName,
		objectLock: opt.objectLock,
		shadow: correctShadowOptions(opt.shadow),
		...(szAuto ? { _szAuto: szAuto } : {}),
	}
	newObject.options = objectOptions

	// STEP 5: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	// Use a namespaced key for media targets so slide master (sm) and slide layouts (sl-N, _slideNum >= 1000)
	// never collide with regular slide media names in large decks.
	const mediaSlideKey =
		target._slideNum == null ? 'sm' : target._slideNum >= 1000 ? `sl-${target._slideNum}` : target._slideNum
	if (strImgExtn === 'svg') {
		// SVG files consume *TWO* rId's: (a png version and the svg image)
		// <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
		// <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>
		target._relsMedia.push({
			path: strImagePath || strImageData + 'png',
			type: 'image/png',
			extn: 'png',
			data: strImageData || '',
			rId: imageRelId,
			Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.png`,
			isSvgPng: true,
			svgSize: {
				w: getSmartParseNumber(objectOptions.w, 'X', target._presLayout),
				h: getSmartParseNumber(objectOptions.h, 'Y', target._presLayout),
			},
		})
		newObject.imageRid = imageRelId
		target._relsMedia.push({
			path: strImagePath || strImageData,
			type: 'image/svg+xml',
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId + 1,
			Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${strImgExtn}`,
		})
		newObject.imageRid = imageRelId + 1
	} else {
		// PERF: Duplicate media should reuse existing `Target` value and not create an additional copy.
		// File-path images are matched by `path`; base64/`data` images have no real path
		// (all share the `preencoded.<extn>` placeholder), so they are matched by their data
		// payload instead so identical inline images are embedded once.
		const imgContentType = imageContentType(strImgExtn)
		const dupeItem = target._relsMedia.find((item) => {
			if (item.isDuplicate || !item.Target || item.type !== imgContentType) return false
			return strImagePath ? item.path === strImagePath : !!strImageData && item.data === strImageData
		})

		target._relsMedia.push({
			path: strImagePath || 'preencoded.' + strImgExtn,
			type: imgContentType,
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId,
			isDuplicate: !!dupeItem?.Target,
			Target: dupeItem?.Target
				? dupeItem.Target
				: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${strImgExtn}`,
		})
		newObject.imageRid = imageRelId
	}

	// STEP 6: Hyperlink support
	if (typeof objHyperlink === 'object') {
		if (!objHyperlink.url && !objHyperlink.slide)
			throw new Error('ERROR: `hyperlink` option requires either: `url` or `slide`')
		else {
			imageRelId++

			target._rels.push({
				type: SlideObjectType.hyperlink,
				data: objHyperlink.slide ? 'slide' : 'dummy',
				rId: imageRelId,
				// `Target` is stored RAW; every emitter escapes it. See the note on `SlideRel.Target`.
				Target: objHyperlink.url ? objHyperlink.url : String(objHyperlink.slide),
			})

			objHyperlink._rId = imageRelId
			newObject.hyperlink = objHyperlink
		}
	}

	// STEP 7: Add object to slide
	target._slideObjects.push(newObject)
}
