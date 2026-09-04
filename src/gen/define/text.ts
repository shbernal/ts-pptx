/**
 * ts-pptx: Text Definition
 *
 * `addTextDefinition` cleans shape / run options (color, bullets, placeholder inheritance, body
 * properties, columns, align / valign), registers hyperlink + picture-bullet + image-fill rels,
 * and pushes a `text` / `placeholder` object. `createBulletImageRels` handles the picture-bullet
 * media rels.
 */
import { AlignH, type PLACEHOLDER_TYPE, ShapeType, SlideObjectType, TextAnchor } from '../../enums.js'
import { DEF_FONT_COLOR, DEF_SHAPE_LINE_COLOR } from '../../constants-internal.js'
import { warn } from '../../diagnostics.js'
import type { ShapeLineProps, TextProps, TextPropsOptions } from '../../types/index.js'
import type { ObjectOptionsInternal, PresSlideInternal, SlideObject } from '../../types/internal.js'
import { encodeXmlAttrValue, getNewRelId, nextMediaTarget } from '../utils.js'
import { registerSvgImageRels } from './image-rel.js'
import { setOrClear } from '../../options-internal.js'
import { normalizeShadowOptions } from '../drawingml/effect.js'
import { resolveFillKind, resolveLineKind } from '../drawingml/fill.js'
import { resolveTextAnchor } from '../drawingml/text-body.js'
import { imageContentType, imageExtensionForSource } from '../../media/content-type.js'
import { ptsToEmuLenient, resolveInsetsEmu } from '../../units-internal.js'
import { resolveObjectName } from './object-name.js'
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
	// Take ownership of the options before touching them. Everything below writes internal state
	// (`_bodyProp`, `objectName`, `_placeholderType`, defaulted `color`/`line`) onto whatever object
	// it is handed, and `gen/drawingml/text-run.ts` writes more of it (`_lineIdx`, paragraph props
	// inherited from the shape) at emit time. Without `own()` that state lands on the CALLER's
	// object, so a style literal reused across shapes carries one shape's settings to the next:
	//
	//   slide.addText('a', STYLE)                      // STYLE now holds a `_bodyProp`
	//   slide.addText('b', { ...STYLE, columns: 2 })   // the spread aliases that same `_bodyProp`…
	//   slide.addText('c', { ...STYLE })               // …so all three shapes emit numCol="2"
	//
	// plus a duplicate-`objectName` warning, because the name assigned to the first shape is spread
	// onto the second.
	//
	// Identity WITHIN one call is preserved on purpose, which is why this memoizes instead of
	// spreading at each use. Sharing between the shape's options and a run's is load-bearing:
	// `SlideBuilder.addText`'s string shorthand hands the same object to both, so `cleanOpts` runs
	// over it twice, and the second pass emits bytes the first cannot (see the line-defaults case in
	// `test/regression/text/text-definition.test.js`). Copying each reference separately would
	// quietly change that shape's `<a:ln>`. The caller is protected either way — the aliasing is
	// between two objects this function now owns.
	//
	// Nested option objects the caller supplies (`bullet`, `shadow`, `fill`) are deliberately still
	// shared: `bullet._rId` and the image fill's rel id are registered through those references and
	// read back at emit time, and auto-paging relies on a cloned text object reaching the same bullet.
	const owned = new Map<TextPropsOptions, ObjectOptionsInternal>()
	/**
	 * The keys the CALLER wrote on each owned copy, recorded before any default reaches it.
	 *
	 * `cleanOpts` defaults several keys (`color`, `bullet`, `line`, `_bodyProp`) on the way past,
	 * so by the time the placeholder inheritance in A.3 runs, "the bag has this key" no longer
	 * means "the caller asked for this". That distinction is the whole rule there: a placeholder
	 * supplies an option the caller left out, and a *default* is not the caller leaving it out.
	 */
	const authoredKeys = new WeakMap<ObjectOptionsInternal, ReadonlySet<string>>()
	/** Copy a caller-supplied options object once, returning the same copy for the same input. */
	const own = (source?: TextPropsOptions): ObjectOptionsInternal => {
		if (!source) return {}
		const already = owned.get(source)
		if (already) return already
		const copy: ObjectOptionsInternal = { ...source }
		owned.set(source, copy)
		authoredKeys.set(copy, new Set(Object.keys(source)))
		return copy
	}

	const textObjects = (!text || text.length === 0 ? [{ text: '' }] : text).map((item) => ({
		...item,
		options: own(item.options),
	}))
	const objectOptions: ObjectOptionsInternal = own(opts)
	const newObject: SlideObject = {
		_type: isPlaceholder ? SlideObjectType.placeholder : SlideObjectType.text,
		shape: opts.shape || ShapeType.rect,
		text: textObjects,
		options: objectOptions,
	}
	// One index for the whole text object, taken here rather than inside `cleanOpts` — that runs once
	// for the object and again for every run, so naming from inside it would burn an index per run.

	function cleanOpts(itemOpts: ObjectOptionsInternal): TextPropsOptions {
		// STEP 1: Set some options
		{
			// A.1: Color (placeholders should inherit their colors or override them, so don't default them)
			if (!itemOpts.placeholder) {
				// A hyperlink run with no color configured anywhere inherits the theme hyperlink color
				// (a:schemeClr hlink, and folHlink once visited), which PowerPoint applies automatically
				// when the run carries no explicit fill. Defaulting it to DEF_FONT_COLOR would emit a
				// solidFill plus hlinkClr="tx", pinning the link to black and suppressing the theme
				// hyperlink/visited colors. Only non-hyperlink text falls back to DEF_FONT_COLOR.
				setOrClear(
					itemOpts,
					'color',
					itemOpts.color ||
						objectOptions.color ||
						target.color ||
						(itemOpts.hyperlink || objectOptions.hyperlink ? undefined : DEF_FONT_COLOR)
				)
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
				if (placeHold?.options) {
					// A placeholder SUPPLIES an option; it never IMPOSES one. Every key the caller wrote
					// on this bag wins, and the placeholder fills in the rest -- which is what "text
					// targeting a placeholder inherits its options" means, and what makes a slide match
					// its layout without silently discarding what the caller asked for.
					//
					// `{ ...itemOpts, ...placeHold.options }` was the other way round, so the placeholder
					// won on every key it stated: `addText('x', { placeholder: 'body', valign: 'top' })`
					// took the layout's anchor and the caller's `valign` did nothing.
					//
					// It is PowerPoint's own model. A layout placeholder given a bottom anchor and a 1in
					// left inset, with the slide's placeholder then re-anchored to the top, writes
					// `<a:bodyPr lIns="914400" anchor="b"/>` on the layout and `<a:bodyPr anchor="t"/>` on
					// the slide: the slide states only what it overrides, the inset is simply absent, and
					// the stated anchor is the one that applies (`placeholder-override.pptx`).
					//
					// `authoredKeys`, not "is the key present": `bullet` is defaulted to `false` a few
					// lines above for exactly these objects, so testing presence would let that default
					// beat the layout's bullet -- which is not the same statement as letting a caller
					// beat it.
					const stated = authoredKeys.get(itemOpts) ?? new Set<string>()
					const inherited: Record<string, unknown> = {}
					for (const [key, value] of Object.entries(placeHold.options)) {
						// A key the placeholder carries as an explicit `undefined` supplies nothing, and
						// copying it would turn "the layout said nothing" into a stated `undefined` on a
						// bag that is spread further downstream.
						if (value !== undefined && !stated.has(key)) inherited[key] = value
					}
					itemOpts = { ...itemOpts, ...inherited }
				}
			}

			// B:
			if (itemOpts.shape === ShapeType.line) {
				const itemLine = typeof itemOpts.line === 'object' && itemOpts.line ? itemOpts.line : {}
				// ShapeLineProps defaults, the same block as define/shape.ts. Spread first,
				// override only what is defaulted here — see there for why listing the carried
				// keys instead silently dropped `gradient`, `pattern` and `cap`. The kind is
				// `resolveLineKind`'s answer, and only a solid stroke gets the default line
				// color; every other kind takes its paint from its own sub-object.
				const itemLineKind = resolveLineKind(itemLine)
				const newLineOpts: ShapeLineProps = {
					...itemLine,
					type: itemLineKind,
					transparency: itemLine.transparency || 0,
					width: itemLine.width || 1,
					dashType: itemLine.dashType || 'solid',
				}
				// Only the solid arm writes `color`. The spread already carried whatever colour the
				// other kinds stated, so re-writing it would turn an unstated one into a present
				// `undefined` — a distinction the next spread of this bag can see.
				if (itemLineKind === 'solid') newLineOpts.color = itemLine.color || DEF_SHAPE_LINE_COLOR
				if (typeof itemOpts.line === 'object') itemOpts.line = newLineOpts
			}

			// C: Line opts
			itemOpts.line = itemOpts.line || {}
			// `NaN` is falsy, so the truthiness test is the whole guard; an out-of-range value is
			// clamped and reported by `clamp.ts` at emit rather than dropped without a word here.
			if (!itemOpts.lineSpacing) delete itemOpts.lineSpacing
			if (!itemOpts.lineSpacingMultiple) delete itemOpts.lineSpacingMultiple

			// D: Transform text options to bodyProperties as thats how we build XML
			// Copy, never adopt: an incoming `_bodyProp` belongs to something else. It arrives here
			// from a caller reusing a literal that a previous `addText` wrote to, and from the
			// A.3 branch above, which spreads a LAYOUT placeholder's options onto the slide's.
			// Adopting it makes the two shapes share one body-property record, so the last writer
			// of `numCol`/`anchor`/`vert` wins for all of them.
			itemOpts._bodyProp = { ...itemOpts._bodyProp }
			// A placeholder inherits its anchor from the layout, so the key comes off rather than
			// being written as an `undefined` — this bag is spread onto a slide's options in A.3.
			setOrClear(itemOpts._bodyProp, 'anchor', !itemOpts.placeholder ? TextAnchor.ctr : undefined) // VALS: [t,ctr,b]
			// `textDirection` is the documented public option; `vert` is a legacy/extended alias kept as an
			// escape hatch for the full ST_TextVerticalType range (eaVert, mongolianVert, wordArtVertRtl).
			// Both map directly to the `<a:bodyPr vert="…">` attribute, so prefer the documented one.
			setOrClear(itemOpts._bodyProp, 'vert', itemOpts.textDirection ?? itemOpts.vert) // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
			itemOpts._bodyProp.wrap = typeof itemOpts.wrap === 'boolean' ? itemOpts.wrap : true
			// The four text insets. This used to be computed by the text serializer, which owns no
			// options and says so: `_bodyProp` is normalized here, and normalizing half of it in
			// one place and half in another is how the two readings drift.
			const insets = resolveInsetsEmu(itemOpts.margin)
			setOrClear(itemOpts._bodyProp, 'lIns', insets?.l)
			setOrClear(itemOpts._bodyProp, 'tIns', insets?.t)
			setOrClear(itemOpts._bodyProp, 'rIns', insets?.r)
			setOrClear(itemOpts._bodyProp, 'bIns', insets?.b)
			setOrClear(itemOpts._bodyProp, 'prstTxWarp', itemOpts.textWarp) // preset text warp (`<a:prstTxWarp>`), e.g. 'textArchUp'

			// D.1: Text columns (`numCol` range is 1-16 per ECMA-376 ST_TextColumnCount)
			if (itemOpts.columns !== undefined) {
				if (
					typeof itemOpts.columns !== 'number' ||
					!Number.isFinite(itemOpts.columns) ||
					itemOpts.columns < 1 ||
					itemOpts.columns > 16
				) {
					warn('text/invalid-columns', 'text `columns` must be a number 1-16 (ignoring value)')
				} else {
					itemOpts._bodyProp.numCol = Math.round(itemOpts.columns)
				}
			}
			if (itemOpts.columnSpacing !== undefined) {
				if (
					typeof itemOpts.columnSpacing !== 'number' ||
					!Number.isFinite(itemOpts.columnSpacing) ||
					itemOpts.columnSpacing < 0
				) {
					warn('text/invalid-column-spacing', 'text `columnSpacing` must be a number >= 0 (ignoring value)')
				} else {
					itemOpts._bodyProp.spcCol = ptsToEmuLenient(itemOpts.columnSpacing)
				}
			}

			// E: Normalize shorthand `underline: true` to the object form
			if (typeof itemOpts.underline === 'boolean' && itemOpts.underline === true) itemOpts.underline = { style: 'sng' }
		}

		// STEP 2: Transform `align`/`valign` to XML values, store in _bodyProp for XML gen
		{
			const align = (itemOpts.align || '').toLowerCase()
			if (align.startsWith('c')) itemOpts._bodyProp.align = AlignH.center
			else if (align.startsWith('l')) itemOpts._bodyProp.align = AlignH.left
			else if (align.startsWith('r')) itemOpts._bodyProp.align = AlignH.right
			else if (align.startsWith('j')) itemOpts._bodyProp.align = AlignH.justify

			const anchor = resolveTextAnchor(itemOpts.valign)
			if (anchor) itemOpts._bodyProp.anchor = anchor
		}

		// STEP 3: ROBUST: Set rational values for some shadow props if needed.
		// Assigned, not discarded: the normalizer is pure, and `_alpha` is what
		// `genXmlTextRunProperties` reads for a run's shadow transparency.
		setOrClear(itemOpts, 'shadow', normalizeShadowOptions(itemOpts.shadow))

		return itemOpts
	}

	// STEP 1: Create/Clean object options
	newObject.options = cleanOpts(objectOptions)

	// STEP 1-A: A text box with no height stated anywhere gets 0.3in of it.
	//
	// `addShapeDefinition` defaults all four axes; this path defaulted none, so `addText('hi',
	// { x, y, w })` emitted `<a:ext cy="0">` — the degenerate zero-size object the project's own
	// API rules say to refuse. There WAS a rescue for it, in the text serializer, guarded by
	// `!itemOpts.line`; step C above has written `itemOpts.line = itemOpts.line || {}` since
	// before that guard existed, so the guard was never once true and the rescue never ran.
	//
	// Two things move it here rather than fixing the guard in place. It sits beside the other
	// defaults, where the caller's own value is still distinguishable from silence — `h: 0`
	// stated explicitly is kept, exactly as `addShapeDefinition` keeps it, which the serializer
	// could not do because a stated zero and an absent height both arrive there as `cy === 0`.
	// And the exemption the guard was reaching for is a *shape* question, not a line-options
	// one: a `line` is drawn zero-height on purpose (a horizontal rule), and that is knowable
	// here from `shape` alone.
	//
	// Read after `cleanOpts`, not before: a caller who named a layout placeholder inherits that
	// placeholder's frame in step A.3, so a height can arrive without the caller stating one.
	if (newObject.options.h === undefined && newObject.shape !== ShapeType.line) newObject.options.h = 0.3

	// STEP 1a: Selection Pane identity (`objectName`). Set once here, on the shape-level object
	// only — not inside `cleanOpts`, which also runs per text run (STEP 2 below). `Slide.addText`'s
	// single-string convenience form reuses the same options object for both the shape and its lone
	// run, so encoding this inside `cleanOpts` encoded a caller-supplied name twice. A placeholder's
	// default identity is its declared name (falling back to its type, then its idx). Placeholders
	// are `placeholder`-typed objects and so take their name index from their own bucket; naming
	// them `Text N` off the text-box bucket would collide with the slide's real text boxes.
	// A placeholder's default identity is its declared name (falling back to its type, then its
	// idx), so it does not take the `Text N` default — but it still takes an index from its own
	// bucket, which is what keeps a slide's real text boxes from colliding with it.
	const placeholderName = isPlaceholder
		? encodeXmlAttrValue(
				String(
					newObject.options.placeholder ||
						newObject.options._placeholderType ||
						`Placeholder ${newObject.options._placeholderIdx ?? target._slideObjects.length}`
				)
			)
		: undefined
	newObject.options.objectName = resolveObjectName(target, newObject._type, {
		label: 'Text',
		kind: 'text',
		supplied: newObject.options.objectName,
		...(placeholderName === undefined ? {} : { fallback: placeholderName }),
	})

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
	if (typeof newObject.options.fill === 'object' && resolveFillKind(newObject.options.fill) === 'image') {
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
 * @param objectOptions - shape-level text options (bullet may live here)
 * @param {TextProps[]} textObjects - per-paragraph text options (bullet may live here too)
 */
function createBulletImageRels(
	target: PresSlideInternal,
	objectOptions: ObjectOptionsInternal,
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

		// REALITY-CHECK: base64 `data` must carry a base64 header (mirror addImage()). Unlike
		// `addImage()` this warns rather than throws: refusing the rel is not fatal, because the run
		// emitter falls back to a default glyph and the deck still opens.
		if (img.data && (typeof img.data !== 'string' || !img.data.toLowerCase().includes('base64,'))) {
			warn(
				'bullet/image-missing-base64-header',
				"bullet.image `data` value lacks a base64 header, ex: 'image/png;base64,iVBOR[...]'"
			)
			return
		}

		// Auto-paging clones text objects onto new slides while sharing the bullet options object by
		// reference, so `_rId` may already be set from the originating slide. Skip when this slide already
		// carries the rel; otherwise (re-)register so the new slide's .rels and media part exist.
		if (bullet._rId && target._relsMedia.some((rel) => rel.rId === bullet._rId)) return

		// Determine extension: the `data:` mime wins, else parse the path (mirror addImageDefinition())
		const strImgExtn = imageExtensionForSource(img.path || '', img.data || '')

		if (strImgExtn === 'svg') {
			// SVG bullets consume *TWO* rels, mirroring addImage(): a PNG preview (referenced by the
			// `<a:buBlip><a:blip r:embed>`) plus the SVG itself (referenced by the `asvg:svgBlip` ext).
			// Auto-paging shares one bullet options object across the overflow slides, so a re-registration
			// has to keep the pair of ids that object already carries rather than mint a new one.
			const pinned = bullet._rId && bullet._rIdSvg ? { pngRid: bullet._rId, svgRid: bullet._rIdSvg } : undefined
			const { pngRid, svgRid } = registerSvgImageRels(target, { path: img.path ?? '', data: img.data ?? '' }, pinned)
			bullet._rId = pngRid
			bullet._rIdSvg = svgRid
		} else {
			const relId = bullet._rId || getNewRelId(target)
			target._relsMedia.push({
				path: img.path || 'preencoded.' + strImgExtn,
				type: imageContentType(strImgExtn),
				extn: strImgExtn,
				data: img.data || '',
				rId: relId,
				Target: nextMediaTarget(target, 'image', strImgExtn),
			})
			bullet._rId = relId
		}
	})
}
