/**
 * PptxGenJS: Text Definition
 *
 * `addTextDefinition` cleans shape / run options (color, bullets, placeholder inheritance, body
 * properties, columns, align / valign), registers hyperlink + picture-bullet + image-fill rels,
 * and pushes a `text` / `placeholder` object. `createBulletImageRels` handles the picture-bullet
 * media rels.
 */
import {
	AlignH,
	DEF_FONT_COLOR,
	DEF_SHAPE_LINE_COLOR,
	type PLACEHOLDER_TYPE,
	ShapeType,
	SlideObjectType,
	TextAnchor,
} from '../../core-enums.js'
import { warn } from '../../log.js'
import type {
	ObjectOptions,
	PresSlideInternal,
	ShapeLineProps,
	SlideObject,
	TextProps,
	TextPropsOptions,
} from '../../core-interfaces.js'
import { correctShadowOptions, encodeXmlEntities, getNewRelId, validateObjectName } from '../../gen-utils.js'
import { imageContentType } from '../../media/content-type.js'
import { valToPts } from '../../units-internal.js'
import { nextObjectNameIdx } from './object-name.js'
import { createHyperlinkRels } from './hyperlinks.js'
import { registerImageFillMedia } from './image.js'

/**
 * Adds a text object to a slide definition.
 * @param {PresSlideInternal} target - slide object that the text should be added to
 * @param {string|TextProps[]} text text string or object
 * @param {TextPropsOptions} opts text options
 * @param {boolean} isPlaceholder whether this a placeholder object
 */
export function addTextDefinition(
	target: PresSlideInternal,
	text: TextProps[],
	opts: TextPropsOptions,
	isPlaceholder: boolean
): void {
	const textObjects = !text || text.length === 0 ? [{ text: '' }] : text
	const objectOptions: ObjectOptions = opts || {}
	const newObject: SlideObject = {
		_type: isPlaceholder ? SlideObjectType.placeholder : SlideObjectType.text,
		shape: opts.shape || ShapeType.rect,
		text: textObjects,
		options: objectOptions,
	}
	// One index for the whole text object, taken here rather than inside `cleanOpts` — that runs once
	// for the object and again for every run, so naming from inside it would burn an index per run.
	const textNameIdx = nextObjectNameIdx(target, newObject._type)

	function cleanOpts(itemOpts: ObjectOptions): TextPropsOptions {
		// STEP 1: Set some options
		{
			// A.1: Color (placeholders should inherit their colors or override them, so don't default them)
			if (!itemOpts.placeholder) {
				// A hyperlink run with no color configured anywhere inherits the theme hyperlink color
				// (a:schemeClr hlink, and folHlink once visited), which PowerPoint applies automatically
				// when the run carries no explicit fill. Defaulting it to DEF_FONT_COLOR would emit a
				// solidFill plus hlinkClr="tx", pinning the link to black and suppressing the theme
				// hyperlink/visited colors. Only non-hyperlink text falls back to DEF_FONT_COLOR.
				itemOpts.color =
					itemOpts.color ||
					objectOptions.color ||
					target.color ||
					(itemOpts.hyperlink || objectOptions.hyperlink ? undefined : DEF_FONT_COLOR)
			}

			// A.2: Placeholder should inherit their bullets or override them, so don't default them
			if (itemOpts.placeholder || isPlaceholder) {
				itemOpts.bullet = itemOpts.bullet || false
			}

			// A.3: Text targeting a placeholder need to inherit the placeholders options (eg: margin, valign, etc.)
			if (itemOpts.placeholder && target._slideLayout && target._slideLayout._slideObjects) {
				const placeHold = target._slideLayout._slideObjects.filter(
					(item) =>
						item._type === SlideObjectType.placeholder &&
						item.options &&
						item.options.placeholder &&
						item.options.placeholder === itemOpts.placeholder
				)[0]
				if (placeHold?.options) itemOpts = { ...itemOpts, ...placeHold.options }
			}

			// B:
			if (itemOpts.shape === ShapeType.line) {
				const itemLine = typeof itemOpts.line === 'object' && itemOpts.line ? itemOpts.line : {}
				// ShapeLineProps defaults
				const newLineOpts: ShapeLineProps = {
					type: itemLine.type || 'solid',
					color: itemLine.color || DEF_SHAPE_LINE_COLOR,
					transparency: itemLine.transparency || 0,
					width: itemLine.width || 1,
					dashType: itemLine.dashType || 'solid',
					beginArrowType: itemLine.beginArrowType,
					endArrowType: itemLine.endArrowType,
				}
				if (typeof itemOpts.line === 'object') itemOpts.line = newLineOpts
			}

			// C: Line opts
			itemOpts.line = itemOpts.line || {}
			itemOpts.lineSpacing = itemOpts.lineSpacing && !isNaN(itemOpts.lineSpacing) ? itemOpts.lineSpacing : undefined
			itemOpts.lineSpacingMultiple =
				itemOpts.lineSpacingMultiple && !isNaN(itemOpts.lineSpacingMultiple) ? itemOpts.lineSpacingMultiple : undefined

			// D: Transform text options to bodyProperties as thats how we build XML
			itemOpts._bodyProp = itemOpts._bodyProp || {}
			itemOpts._bodyProp.anchor = !itemOpts.placeholder ? TextAnchor.ctr : undefined // VALS: [t,ctr,b]
			// `textDirection` is the documented public option; `vert` is a legacy/extended alias kept as an
			// escape hatch for the full ST_TextVerticalType range (eaVert, mongolianVert, wordArtVertRtl).
			// Both map directly to the `<a:bodyPr vert="…">` attribute, so prefer the documented one.
			itemOpts._bodyProp.vert = itemOpts.textDirection ?? itemOpts.vert // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
			itemOpts._bodyProp.wrap = typeof itemOpts.wrap === 'boolean' ? itemOpts.wrap : true
			itemOpts._bodyProp.prstTxWarp = itemOpts.textWarp // preset text warp (`<a:prstTxWarp>`), e.g. 'textArchUp'

			// D.1: Text columns (`numCol` range is 1-16 per ECMA-376 ST_TextColumnCount)
			if (itemOpts.columns !== undefined) {
				if (
					typeof itemOpts.columns !== 'number' ||
					isNaN(itemOpts.columns) ||
					itemOpts.columns < 1 ||
					itemOpts.columns > 16
				) {
					warn('text `columns` must be a number 1-16 (ignoring value)')
				} else {
					itemOpts._bodyProp.numCol = Math.round(itemOpts.columns)
				}
			}
			if (itemOpts.columnSpacing !== undefined) {
				if (typeof itemOpts.columnSpacing !== 'number' || isNaN(itemOpts.columnSpacing) || itemOpts.columnSpacing < 0) {
					warn('text `columnSpacing` must be a number >= 0 (ignoring value)')
				} else {
					itemOpts._bodyProp.spcCol = valToPts(itemOpts.columnSpacing)
				}
			}

			// E: Normalize shorthand `underline: true` to the object form
			if (typeof itemOpts.underline === 'boolean' && itemOpts.underline === true) itemOpts.underline = { style: 'sng' }
		}

		// STEP 2: Transform `align`/`valign` to XML values, store in _bodyProp for XML gen
		{
			const align = (itemOpts.align || '').toLowerCase()
			const valign = (itemOpts.valign || '').toLowerCase()
			if (align.startsWith('c')) itemOpts._bodyProp.align = AlignH.center
			else if (align.startsWith('l')) itemOpts._bodyProp.align = AlignH.left
			else if (align.startsWith('r')) itemOpts._bodyProp.align = AlignH.right
			else if (align.startsWith('j')) itemOpts._bodyProp.align = AlignH.justify

			if (valign.startsWith('b')) itemOpts._bodyProp.anchor = TextAnchor.b
			else if (valign.startsWith('m')) itemOpts._bodyProp.anchor = TextAnchor.ctr
			else if (valign.startsWith('t')) itemOpts._bodyProp.anchor = TextAnchor.t
		}

		// STEP 3: ROBUST: Set rational values for some shadow props if needed
		correctShadowOptions(itemOpts.shadow)

		return itemOpts
	}

	// STEP 1: Create/Clean object options
	newObject.options = cleanOpts(objectOptions)

	// STEP 1a: Selection Pane identity (`objectName`). Set once here, on the shape-level object
	// only — not inside `cleanOpts`, which also runs per text run (STEP 2 below). `Slide.addText`'s
	// single-string convenience form reuses the same options object for both the shape and its lone
	// run, so encoding this inside `cleanOpts` encoded a caller-supplied name twice. A placeholder's
	// default identity is its declared name (falling back to its type, then its idx). Placeholders
	// are `placeholder`-typed objects and so take their name index from their own bucket; naming
	// them `Text N` off the text-box bucket would collide with the slide's real text boxes.
	newObject.options.objectName = newObject.options.objectName
		? encodeXmlEntities(validateObjectName(newObject.options.objectName, 'text'))
		: isPlaceholder
			? encodeXmlEntities(
					String(
						newObject.options.placeholder ||
							newObject.options._placeholderType ||
							`Placeholder ${newObject.options._placeholderIdx ?? target._slideObjects.length}`
					)
				)
			: `Text ${textNameIdx}`

	// STEP 1b: Standalone placeholder type (accessibility "Missing Slide Title")
	// `placeholder` is documented as a placeholder *type* ('title', 'body', et. al.). When it
	// resolves to a layout placeholder the layout object supplies the <p:ph> at serialize time,
	// but with a blank/default layout there is no match and no <p:ph> was emitted - so PowerPoint's
	// accessibility checker reports the slide as having no title. Record the type here so a real
	// <p:ph type="..."/> is emitted on the slide shape even without a matching layout placeholder.
	if (!isPlaceholder && newObject.options.placeholder && !newObject.options._placeholderType) {
		newObject.options._placeholderType = newObject.options.placeholder as PLACEHOLDER_TYPE
	}

	// STEP 2: Create/Clean text options
	textObjects.forEach((item) => (item.options = cleanOpts(item.options || {})))

	// STEP 3: Create hyperlinks
	createHyperlinkRels(target, textObjects)

	// STEP 4: Create picture-bullet image rels
	createBulletImageRels(target, newObject.options, textObjects)

	// STEP 5: Register an image fill (if any) as a media relationship for serialize-time blipFill
	if (
		typeof newObject.options.fill === 'object' &&
		(newObject.options.fill.type === 'image' || newObject.options.fill.image)
	) {
		registerImageFillMedia(target, newObject.options.fill)
	}

	// LAST: Add object to Slide
	target._slideObjects.push(newObject)
}

/**
 * Register slide media relationships for any picture bullets (`bullet.image`) used by a text object.
 * Picture bullets render as `<a:buBlip><a:blip r:embed="rId.."/></a:buBlip>`, so the bullet image
 * needs the same media-rel + package-part plumbing as `addImage()`. The assigned `rId` is stored on
 * the bullet options object (`_rId`) so XML generation can reference it.
 * @param {PresSlideInternal} target - slide receiving the rels
 * @param {ObjectOptions} objectOptions - shape-level text options (bullet may live here)
 * @param {TextProps[]} textObjects - per-paragraph text options (bullet may live here too)
 */
function createBulletImageRels(
	target: PresSlideInternal,
	objectOptions: ObjectOptions,
	textObjects: TextProps[]
): void {
	// Collect every bullet options object that requests a picture bullet (shape-level + per-paragraph).
	// Shape-level bullets are later shared by reference onto the first run, so the same object may appear
	// twice; the `_rId` guard below makes the registration idempotent.
	const bulletObjs: Array<{ image?: { path?: string; data?: string }; _rId?: number; _rIdSvg?: number }> = []
	const collect = (opts?: TextPropsOptions): void => {
		if (opts && typeof opts.bullet === 'object' && opts.bullet) bulletObjs.push(opts.bullet)
	}
	collect(objectOptions)
	textObjects.forEach((item) => collect(item.options))

	bulletObjs.forEach((bullet) => {
		const img = bullet.image
		if (!img || (!img.path && !img.data)) return

		// REALITY-CHECK: base64 `data` must carry a base64 header (mirror addImage())
		if (img.data && (typeof img.data !== 'string' || !img.data.toLowerCase().includes('base64,'))) {
			console.error("ERROR: bullet.image `data` value lacks a base64 header! Ex: 'image/png;base64,iVBOR[...]'")
			return
		}

		// Auto-paging clones text objects onto new slides while sharing the bullet options object by
		// reference, so `_rId` may already be set from the originating slide. Skip when this slide already
		// carries the rel; otherwise (re-)register so the new slide's .rels and media part exist.
		if (bullet._rId && target._relsMedia.some((rel) => rel.rId === bullet._rId)) return

		// Determine extension: path wins, else sniff the data: mime-type (mirror addImageDefinition())
		let strImgExtn = 'png'
		if (img.path) {
			const imagePathFile = img.path.slice(img.path.lastIndexOf('/') + 1).split('?')[0] || ''
			strImgExtn = ((imagePathFile.split('.').pop() || 'png').split('#')[0] || 'png').toLowerCase()
		}
		const imageMimeMatch = /image\/(\w+);/.exec(img.data || '')
		if (img.data && imageMimeMatch) strImgExtn = imageMimeMatch[1] ?? strImgExtn
		// `image/svg+xml` does not match the `\w+` sniff above (the `+`), so detect it explicitly (mirror addImageDefinition())
		else if (img.data?.toLowerCase().includes('image/svg+xml')) strImgExtn = 'svg'
		// Path-based SVG sniffing is already handled by the extension parse above.

		const relId = bullet._rId || getNewRelId(target)
		const mediaSlideKey =
			target._slideNum == null ? 'sm' : target._slideNum >= 1000 ? `sl-${target._slideNum}` : target._slideNum

		if (strImgExtn === 'svg') {
			// SVG bullets consume *TWO* rels, mirroring addImage(): a PNG preview (referenced by the
			// `<a:buBlip><a:blip r:embed>`) plus the SVG itself (referenced by the `asvg:svgBlip` ext).
			// The preview rel is flagged `isSvgPng` so the media pipeline generates its PNG fallback.
			target._relsMedia.push({
				path: img.path || img.data + 'png',
				type: 'image/png',
				extn: 'png',
				data: img.data || '',
				rId: relId,
				Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.png`,
				isSvgPng: true,
			})
			target._relsMedia.push({
				path: img.path || img.data || 'preencoded.svg',
				type: 'image/svg+xml',
				extn: 'svg',
				data: img.data || '',
				rId: relId + 1,
				Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.svg`,
			})
			bullet._rId = relId
			bullet._rIdSvg = relId + 1
		} else {
			target._relsMedia.push({
				path: img.path || 'preencoded.' + strImgExtn,
				type: imageContentType(strImgExtn),
				extn: strImgExtn,
				data: img.data || '',
				rId: relId,
				Target: `../media/image-${mediaSlideKey}-${target._relsMedia.length + 1}.${strImgExtn}`,
			})
			bullet._rId = relId
		}
	})
}
