/**
 * PptxGenJS: slide object serialization
 *
 * The per-shape `<p:spTree>` builder: `slideObjectToXml` walks a slide/layout's
 * objects (recursing into groups, allocating `<p:cNvPr>` ids as it goes) and the
 * `render*Object` helpers emit each shape kind (table, text, connector, image,
 * media, chart). `slideObjectRelationsToXml` emits the matching `.rels` targets.
 *
 * The group render mutates a slide-wide child-id counter shared across the
 * render helpers, so this cluster is kept co-located in one module.
 */

import { ChartType, isChartExType, SlideObjectType } from '../../core-enums.js'
import {
	CRLF,
	DEF_CELL_MARGIN_IN,
	DEF_PRES_LAYOUT_NAME,
	DEF_TEXT_SHADOW,
	SLDNUMFLDID,
	XML_DECL,
} from '../../core-enums-internal.js'
import type { ObjectOptions, HyperlinkProps, ShapeLineProps, TableCell, TableCellProps } from '../../core-interfaces.js'
import type {
	PresSlideInternal,
	SlideLayoutInternal,
	SlideObject,
	SlideRel,
	SlideRelChart,
	SlideRelMedia,
	ZoomInternal,
	ZoomTileInternal,
} from '../../types/internal.js'
import { encodeXmlEntities, getDuplicateObjectNames, isHyperlinkRel } from '../../gen-utils.js'
import { createColorElement } from '../drawingml/color.js'
import { createShadowEffectLst } from '../drawingml/effect.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { createLineCap, genXmlLineFill } from '../drawingml/line.js'
import { getImageSizeFromBase64 } from '../../media/image-size.js'
import {
	convertRotationDegrees,
	getSmartParseNumber,
	inch2Emu,
	lineWidthToEmu,
	marginToEmu,
	resolveTableColWidthsEmu,
} from '../../units-internal.js'
import { EMU_PER_INCH, FIXED_PCT_PER_PERCENT, PERCENT_SCALE, pixelsToEmu } from '../../units.js'
import { warn } from '../../log.js'
import { clampFontSizeSz } from '../drawingml/clamp.js'
import { genXmlCustGeom, genXmlPresetGeom } from '../drawingml/geometry.js'
import { genXmlImageCrop, ImageSizingXml } from '../drawingml/image.js'
import {
	genXmlObjectLock,
	GRAPHIC_FRAME_LOCK_ATTRS,
	GROUP_SHAPE_LOCK_ATTRS,
	PICTURE_LOCK_ATTRS,
	SHAPE_LOCK_ATTRS,
} from '../drawingml/locks.js'
import { genTableCellBorderXml } from '../drawingml/table-border.js'
import { genXmlPlaceholder, genXmlTextBody, objectHasMath } from '../drawingml/text-body.js'
import { el, raw, voidEl, type XmlAttrs, type XmlChild } from '../oxml/el.js'
import { collectSlideShapeIds, resolveObjectNameToId } from './shape-ids.js'

const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
/** The MS-2007 `media` rel that pairs with an ECMA audio/video/online rel on the same Target. */
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'
const DML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
/** VML namespace — declared by an OLE object's `mc:Choice Requires="v"` (no VML content is emitted). */
const VML_NS = 'urn:schemas-microsoft-com:vml'
/** graphicData URI for an embedded OLE object (`<p:oleObj>`). */
const OLE_NS = 'http://schemas.openxmlformats.org/presentationml/2006/ole'
/** graphicData URI + child namespace for chartEx charts (referenced via `<mc:AlternateContent>`). */
const CHARTEX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex'
/** Zoom (Slide/Section/Summary) graphicData URI + `mc:Choice Requires` prefix + element local-names, per variant. */
const ZOOM_VARIANTS = {
	slide: {
		uri: 'http://schemas.microsoft.com/office/powerpoint/2016/slidezoom',
		prefix: 'pslz',
		zm: 'sldZm',
		obj: 'sldZmObj',
	},
	section: {
		uri: 'http://schemas.microsoft.com/office/powerpoint/2016/sectionzoom',
		prefix: 'psez',
		zm: 'sectionZm',
		obj: 'sectionZmObj',
	},
	summary: {
		uri: 'http://schemas.microsoft.com/office/powerpoint/2016/summaryzoom',
		prefix: 'psuz',
		zm: 'summaryZm',
		obj: 'summaryZmObj',
	},
} as const
/** Namespace of the zoom preview `blipFill`/`spPr` (children of `zmPr`). */
const P166_NS = 'http://schemas.microsoft.com/office/powerpoint/2016/6/main'
/** Slide→chartEx-part relationship type (MS, not the ECMA `.../relationships/chart`). */
const CHARTEX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx'
/**
 * chartEx feature-version namespace declared on `<mc:Choice Requires>`. Each 2016 chart wave
 * introduced a feature level a consumer must "understand" to render the chart; a consumer that
 * doesn't falls through to `<mc:Fallback>`. Keyed by `ChartType`.
 */
const CHARTEX_FEATURE_NS: Partial<Record<ChartType, { prefix: string; uri: string }>> = {
	[ChartType.waterfall]: { prefix: 'cx1', uri: 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex' },
	[ChartType.funnel]: { prefix: 'cx2', uri: 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex' },
	[ChartType.regionMap]: { prefix: 'cx4', uri: 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex' },
}

/**
 * The `<p:cNvPr>` OPEN tag shared by every shape renderer. Callers append `/>` or
 * `>`+children+`</p:cNvPr>` — the element is self-closing for some shape kinds and paired
 * (hyperlink / media-action children) for others.
 *
 * NOT built with the element builder, deliberately. `descr` is escaped here but `name` is **not**,
 * and that asymmetry is intentional and load-bearing: `objectName` is caller-supplied free text,
 * but every `add*Definition` (text.ts, shape.ts, image.ts, chart.ts, media.ts, connector.ts,
 * group.ts, table.ts) already runs it through `encodeXmlEntities(validateObjectName(...))` once
 * before it reaches a slide object's `options`. Escaping it again here would double-encode it
 * (`'Q&A'` -> `Q&amp;A` upstream -> `Q&amp;amp;A` if escaped here too). This helper exists so that
 * single escape stays one line in one place instead of eight call sites re-deriving it.
 * @param id - the shape's `<p:cNvPr>` id, unique slide-wide
 * @param name - caller-supplied `objectName`, already escaped once upstream (emitted as-is)
 * @param descr - alt text (escaped)
 * @param openPrefix - byte-significant indentation before `<p:cNvPr`
 * @returns the open tag, without its closing delimiter
 */
function cNvPrOpen(id: number, name: string | undefined, descr: string, openPrefix = ''): string {
	return `${openPrefix}<p:cNvPr id="${id}" name="${name}" descr="${encodeXmlEntities(descr)}"`
}

/**
 * The `<a:hlinkClick>` children of a shape's `<p:cNvPr>` — a URL link and/or a jump to another
 * slide. Shared by the text and image renderers, which emitted identical copies.
 *
 * NOTE: the tooltip is passed through UNESCAPED and escaped once by the element builder. Escaping
 * it here as well would emit `&amp;amp;` for a tooltip containing `&`.
 * @param link - the shape's hyperlink, if any
 * @returns zero, one or two `<a:hlinkClick>` elements
 */
function cNvPrHyperlink(link: HyperlinkProps | undefined): string {
	if (!link) return ''
	const tooltip = link.tooltip ?? ''
	return (
		(link.url ? voidEl('a:hlinkClick', { 'r:id': `rId${link._rId}`, tooltip }) : '') +
		(link.slide
			? voidEl('a:hlinkClick', { 'r:id': `rId${link._rId}`, tooltip, action: 'ppaction://hlinksldjump' })
			: '') +
		// Action buttons: a self-contained slide-show navigation action. No relationship, so `r:id`
		// is emitted empty (schema-optional on CT_Hyperlink; matches PowerPoint's own output).
		(link.action
			? voidEl('a:hlinkClick', { 'r:id': '', tooltip, action: `ppaction://hlinkshowjump?jump=${link.action}` })
			: '')
	)
}

/**
 * A shape outline: `<a:ln>` with its fill, dash pattern and arrow ends.
 *
 * The text, connector and image renderers each carried a byte-identical copy of this block; they
 * now share one. Attribute and child order are byte-significant and preserved as written.
 *
 * FUTURE: arrow-size support via the `w`/`len` attrs on headEnd/tailEnd
 * (e.g. `<a:headEnd type="arrow" w="lg" len="lg"/>`; each is 'sm'|'med'|'lg', a 3x3 grid)
 * @param ln - the shape's line properties
 * @returns the `<a:ln>` element
 */
function genXmlShapeLine(ln: ShapeLineProps): string {
	return el('a:ln', { w: ln.width ? lineWidthToEmu(ln.width) : null, cap: ln.cap ? createLineCap(ln.cap) : null }, [
		raw(genXmlLineFill(ln)),
		ln.dashType ? raw(voidEl('a:prstDash', { val: ln.dashType })) : null,
		ln.beginArrowType ? raw(voidEl('a:headEnd', { type: ln.beginArrowType })) : null,
		ln.endArrowType ? raw(voidEl('a:tailEnd', { type: ln.endArrowType })) : null,
	])
}

type TableInheritableOption =
	| 'align'
	| 'bold'
	| 'border'
	| 'color'
	| 'fill'
	| 'fontFace'
	| 'fontSize'
	| 'margin'
	| 'textDirection'
	| 'underline'
	| 'valign'
type TableInheritableValue = ObjectOptions[TableInheritableOption]

/** The four axes that make up an explicit group frame. All or nothing — see `givenGroupFrameAxes`. */
const GROUP_FRAME_AXES = ['x', 'y', 'w', 'h'] as const

/**
 * Which of `x`/`y`/`w`/`h` a group's caller actually supplied.
 *
 * A group frame is explicit only when **all four** are given; anything less falls back to
 * auto-bounds on every axis. Partial frames used to be honoured per-axis, which let the unset ones
 * take the shared per-object defaults (`x=0`, `y=0`, `cx=75%` of the layout width, `cy=0`) — so
 * `addGroup([rect], { x: 5, y: 2 })` emitted a zero-height group whose width was a silent
 * slide-width fraction, and every child of it then re-read as `null` through the read path's
 * degenerate-`chExt` guard.
 *
 * Per-axis fallback was the other candidate, but it cannot mean what it reads like: the writer keeps
 * an identity child space (`chOff/chExt == off/ext`), so a group's frame never moves or scales its
 * children — it only places the selection handle and the rotate pivot. `{ x: 5 }` would leave the
 * children where they were and put the group's box somewhere they are not. Falling back whole, with
 * a warning, keeps the group box around its content and says so out loud.
 * @param options - the group object's options
 * @returns the supplied axis names, in `x`/`y`/`w`/`h` order
 */
const givenGroupFrameAxes = (options: ObjectOptions): Array<(typeof GROUP_FRAME_AXES)[number]> =>
	GROUP_FRAME_AXES.filter((axis) => typeof options[axis] !== 'undefined')

/**
 * Whether a group's frame is fully explicit (all four of `x`/`y`/`w`/`h`), and so should be used
 * verbatim instead of the children's bounding box. See `givenGroupFrameAxes`.
 * @param options - the group object's options
 * @returns true when every axis is supplied
 */
const hasCompleteGroupFrame = (options: ObjectOptions): boolean =>
	givenGroupFrameAxes(options).length === GROUP_FRAME_AXES.length

/**
 * Transforms a slide or slideLayout to resulting XML string - Creates `ppt/slide*.xml`
 * @param {PresSlideInternal|SlideLayoutInternal} slideObject - slide object created within createSlideObject
 * @return {string} XML string with <p:cSld> as the root
 */
export function slideObjectToXml(slide: PresSlideInternal | SlideLayoutInternal): string {
	// `_name` is escaped HERE, at emission, unlike `objectName`'s single-escape-upstream design
	// (see `cNvPrOpen`): `_name` doubles as the raw lookup key `addSlide({masterTitle})` matches
	// against the caller's `title` string (pptxgen.ts, `layout._name === masterSlideName`), so it
	// must stay unescaped until the last possible moment or that match breaks for any title
	// containing `&`/`<`/`"`. Plain slides' default `_name` ("Slide N", slide.ts) never contains
	// XML metacharacters, so escaping it here is a no-op for that path.
	// The element stays a template because it wraps the entire slide, built by append below.
	let strSlideXml: string = slide._name ? '<p:cSld name="' + encodeXmlEntities(slide._name) + '">' : '<p:cSld>'

	// Warn on duplicate Selection Pane identities within this slide. Unique `objectName`
	// values are what consumers (e.g. semantic manifests) rely on, so flag collisions loudly.
	// Groups are recursed into: a group's children are `<p:cNvPr>`-named on this same slide, so a
	// child colliding with a top-level object (or with a child of another group) is a collision the
	// Selection Pane shows — checking only the top level cannot see it.
	const collectObjectNames = (objects: SlideObject[]): string[] =>
		objects.flatMap((obj) => [
			...(typeof obj.options?.objectName === 'string' ? [obj.options.objectName] : []),
			...collectObjectNames(obj._groupObjects || []),
		])
	const duplicateObjectNames = getDuplicateObjectNames(collectObjectNames(slide._slideObjects))
	if (duplicateObjectNames.length > 0) {
		warn(
			`duplicate objectName value(s) emitted on a single slide: ${duplicateObjectNames.join(', ')}. Selection Pane identities should be unique.`
		)
	}

	// STEP 1: Add background color/image (ensure only a single `<p:bg>` tag is created, ex: when master-baskground has both `color` and `path`)
	if (slide._bkgdImgRid) {
		strSlideXml += el(
			'p:bg',
			null,
			raw(
				el('p:bgPr', null, [
					raw(
						el('a:blipFill', { dpi: '0', rotWithShape: '1' }, [
							raw(el('a:blip', { 'r:embed': `rId${slide._bkgdImgRid}` }, raw(voidEl('a:lum')))),
							raw(voidEl('a:srcRect')),
							raw(el('a:stretch', null, raw(voidEl('a:fillRect')))),
						])
					),
					raw(voidEl('a:effectLst')),
				])
			)
		)
	} else if (slide.background?.color || slide.background?.type === 'gradient') {
		strSlideXml += el(
			'p:bg',
			null,
			raw(el('p:bgPr', null, [raw(genXmlColorSelection(slide.background)), raw(voidEl('a:effectLst'))]))
		)
	} else if (!slide.background && slide._name && slide._name === DEF_PRES_LAYOUT_NAME) {
		// NOTE: Default [white] background is needed on slideMaster1.xml to avoid gray background in Keynote (and Finder previews)
		strSlideXml += el('p:bg', null, raw(el('p:bgRef', { idx: '1001' }, raw(voidEl('a:schemeClr', { val: 'bg1' })))))
	}

	// STEP 2: Continue slide by starting spTree node
	// spTree root — OOXML requires the shape tree to open with the implicit top-level group's
	// non-visual props (`<p:nvGrpSpPr>`, the reserved `cNvPr id="1"`) and an identity group
	// transform (off/ext and chOff/chExt all zero) before any child shape. This is the slide's
	// built-in root group, not a user-authored `addGroup` — hence the zeroed frame.
	strSlideXml += '<p:spTree>'
	strSlideXml += el('p:nvGrpSpPr', null, [
		raw(voidEl('p:cNvPr', { id: '1', name: '' })),
		raw(voidEl('p:cNvGrpSpPr')),
		raw(voidEl('p:nvPr')),
	])
	strSlideXml += el(
		'p:grpSpPr',
		null,
		raw(
			el('a:xfrm', null, [
				raw(voidEl('a:off', { x: '0', y: '0' })),
				raw(voidEl('a:ext', { cx: '0', cy: '0' })),
				raw(voidEl('a:chOff', { x: '0', y: '0' })),
				raw(voidEl('a:chExt', { cx: '0', cy: '0' })),
			])
		)
	)

	// Every object's <p:cNvPr> id, up front, for the references that cannot wait for the walk below
	// to reach their target (connector shape bindings, animation spids). `collectSlideShapeIds`
	// mirrors the allocation performed here — keep the two in step.
	const shapeIds = collectSlideShapeIds(slide._slideObjects)

	// STEP 3: Loop over all Slide.data objects and add them to this slide.
	// Allocates <p:cNvPr id> values for group children, which are not in `slide._slideObjects`
	// and so cannot reuse the top-level `idx + 2` scheme without colliding. Seeded past the last
	// top-level object id so child ids stay unique slide-wide.
	let childIdxAlloc = slide._slideObjects.length

	// Resolve an object's bounds in EMU. For an auto-sized group (a `group` object with no explicit
	// x/y/w/h) this recurses into `_groupObjects` and returns their bounding box — so a parent group
	// can size around a nested auto-sized child group. The group rendering below uses the same helper
	// for both a child's bounds and a group's own off/ext, keeping every level consistent.
	const resolveObjBounds = (obj: SlideObject): { x: number; y: number; cx: number; cy: number } => {
		const o = obj.options || {}
		// Shares `hasCompleteGroupFrame` with the group renderer below, so a partial frame resolves to
		// the same auto-bounds here (where a parent group sizes around this one) as it does there.
		if (obj._type === SlideObjectType.group && !hasCompleteGroupFrame(o)) {
			const kids = (obj._groupObjects || []).map(resolveObjBounds)
			if (kids.length === 0) return { x: 0, y: 0, cx: 0, cy: 0 }
			const minX = Math.min(...kids.map((b) => b.x))
			const minY = Math.min(...kids.map((b) => b.y))
			const maxX = Math.max(...kids.map((b) => b.x + b.cx))
			const maxY = Math.max(...kids.map((b) => b.y + b.cy))
			return { x: minX, y: minY, cx: maxX - minX, cy: maxY - minY }
		}
		return {
			x: typeof o.x !== 'undefined' ? getSmartParseNumber(o.x, 'X', slide._presLayout) : 0,
			y: typeof o.y !== 'undefined' ? getSmartParseNumber(o.y, 'Y', slide._presLayout) : 0,
			cx: typeof o.w !== 'undefined' ? getSmartParseNumber(o.w, 'X', slide._presLayout) : 0,
			cy: typeof o.h !== 'undefined' ? getSmartParseNumber(o.h, 'Y', slide._presLayout) : 0,
		}
	}

	// Render one slide object — and, for a group, its children recursively — to an XML fragment.
	// Closes over `slide` and `childIdxAlloc`. Uses a local
	// `strSlideXml` accumulator (shadowing the slide-level one) so the existing per-object
	// `strSlideXml +=` appends compose into the returned fragment unchanged.
	const renderSlideObjectXml = (slideItemObj: SlideObject, idx: number): string => {
		let strSlideXml = ''
		let x = 0
		let y = 0
		let cx = getSmartParseNumber('75%', 'X', slide._presLayout)
		let cy = 0
		let placeholderObj: SlideObject | null = null
		const sizing: ObjectOptions['sizing'] = slideItemObj.options?.sizing
		const rounding = slideItemObj.options?.rounding

		const slideLayout = (slide as PresSlideInternal)._slideLayout
		const wantedPlaceholder = slideItemObj.options?.placeholder
		if (slideLayout?._slideObjects !== undefined && wantedPlaceholder) {
			placeholderObj =
				slideLayout._slideObjects.filter(
					(object: SlideObject) => object.options?.placeholder === wantedPlaceholder
				)[0] ?? null
		}

		// A: Set option vars
		slideItemObj.options = slideItemObj.options || {}
		const itemOpts = slideItemObj.options

		if (typeof slideItemObj.options.x !== 'undefined')
			x = getSmartParseNumber(slideItemObj.options.x, 'X', slide._presLayout)
		if (typeof slideItemObj.options.y !== 'undefined')
			y = getSmartParseNumber(slideItemObj.options.y, 'Y', slide._presLayout)
		if (typeof slideItemObj.options.w !== 'undefined')
			cx = getSmartParseNumber(slideItemObj.options.w, 'X', slide._presLayout)
		if (typeof slideItemObj.options.h !== 'undefined')
			cy = getSmartParseNumber(slideItemObj.options.h, 'Y', slide._presLayout)

		// Set w/h now that smart parse is done
		const imgWidth = cx
		const imgHeight = cy

		// If using a placeholder then inherit it's position
		if (placeholderObj) {
			const phOpts = placeholderObj.options ?? {}
			if (phOpts.x || phOpts.x === 0) x = getSmartParseNumber(phOpts.x, 'X', slide._presLayout)
			if (phOpts.y || phOpts.y === 0) y = getSmartParseNumber(phOpts.y, 'Y', slide._presLayout)
			if (phOpts.w || phOpts.w === 0) cx = getSmartParseNumber(phOpts.w, 'X', slide._presLayout)
			if (phOpts.h || phOpts.h === 0) cy = getSmartParseNumber(phOpts.h, 'Y', slide._presLayout)
		}
		// The `<a:xfrm>` placement attributes, shared by every shape kind that has a transform.
		// NOTE: order is byte-significant (flipH, flipV, rot), and `null` means omitted — `rotate: 0`
		// stays absent, matching the truthiness test this replaced.
		const locationAttrs: XmlAttrs = {
			flipH: slideItemObj.options.flipH ? '1' : null,
			flipV: slideItemObj.options.flipV ? '1' : null,
			rot: slideItemObj.options.rotate ? convertRotationDegrees(slideItemObj.options.rotate) : null,
		}

		// B: Add OBJECT to the current Slide
		switch (slideItemObj._type) {
			case SlideObjectType.table:
				strSlideXml += renderTableObject(slideItemObj, idx, x, y, cx, cy, placeholderObj, itemOpts)
				break
			case SlideObjectType.text:
			case SlideObjectType.placeholder:
				strSlideXml += renderTextObject(slideItemObj, idx, slide, x, y, cx, cy, placeholderObj, locationAttrs)
				break
			case SlideObjectType.connector:
				strSlideXml += renderConnectorObject(slideItemObj, idx, x, y, cx, cy, locationAttrs, shapeIds)
				break
			case SlideObjectType.image:
				strSlideXml += renderImageObject(
					slideItemObj,
					idx,
					slide,
					x,
					y,
					cx,
					cy,
					imgWidth,
					imgHeight,
					placeholderObj,
					locationAttrs,
					sizing,
					rounding
				)
				break
			case SlideObjectType.media:
				strSlideXml += renderMediaObject(slideItemObj, idx, x, y, cx, cy, locationAttrs)
				break
			case SlideObjectType.chart:
				strSlideXml += renderChartObject(slideItemObj, idx, x, y, cx, cy, placeholderObj)
				break
			case SlideObjectType.oleObject:
				strSlideXml += renderOleObject(slideItemObj, idx, x, y, cx, cy)
				break
			case SlideObjectType.zoom:
				strSlideXml += renderZoomObject(slideItemObj, idx, x, y, cx, cy)
				break

			case SlideObjectType.group: {
				const groupChildren = slideItemObj._groupObjects || []

				// Render children (recursively for nested groups). Each child gets a unique id via
				// `childIdxAlloc` (children are not in `_slideObjects`); the shared counter keeps ids
				// collision-free across nesting depth.
				let innerXml = ''
				groupChildren.forEach((child) => {
					child.options = child.options || {}
					innerXml += renderSlideObjectXml(child, childIdxAlloc++)
				})

				// Identity child coordinate space (chOff/chExt == off/ext) at every depth, so children
				// keep their slide-absolute coordinates. Use explicit x/y/w/h when all four are given,
				// else the bounding box of the children (recursing into nested auto-sized groups).
				// A partial frame warns and falls back whole rather than letting the unset axes take the
				// per-object defaults above (`cy` = 0 among them) and emit a degenerate group.
				const givenAxes = givenGroupFrameAxes(slideItemObj.options)
				if (givenAxes.length > 0 && givenAxes.length < GROUP_FRAME_AXES.length) {
					const missingAxes = GROUP_FRAME_AXES.filter((axis) => !givenAxes.includes(axis))
					warn(
						`addGroup: group "${slideItemObj.options.objectName ?? ''}" has a partial frame (${givenAxes.join('/')} given, ${missingAxes.join('/')} missing); using auto-bounds (the bounding box of its children) instead. Pass all of x/y/w/h, or none.`
					)
				}
				const gb = hasCompleteGroupFrame(slideItemObj.options) ? { x, y, cx, cy } : resolveObjBounds(slideItemObj)
				const gx: number = gb.x
				const gy: number = gb.y
				const gcx: number = gb.cx
				const gcy: number = gb.cy

				const grpLockXml = genXmlObjectLock(
					'a:grpSpLocks',
					GROUP_SHAPE_LOCK_ATTRS,
					slideItemObj.options.objectLock,
					slideItemObj.options.objectName
				)
				strSlideXml += '<p:grpSp>'
				strSlideXml += el('p:nvGrpSpPr', null, [
					raw(cNvPrOpen(idx + 2, slideItemObj.options.objectName, slideItemObj.options.altText || '') + '/>'),
					// Paired only when there are locks to carry; otherwise self-closing.
					raw(grpLockXml ? el('p:cNvGrpSpPr', null, raw(grpLockXml)) : voidEl('p:cNvGrpSpPr')),
					raw(voidEl('p:nvPr')),
				])
				strSlideXml += el(
					'p:grpSpPr',
					null,
					raw(
						el('a:xfrm', locationAttrs, [
							raw(voidEl('a:off', { x: gx, y: gy })),
							raw(voidEl('a:ext', { cx: gcx, cy: gcy })),
							raw(voidEl('a:chOff', { x: gx, y: gy })),
							raw(voidEl('a:chExt', { cx: gcx, cy: gcy })),
						])
					)
				)
				strSlideXml += innerXml
				strSlideXml += '</p:grpSp>'
				break
			}

			default:
				strSlideXml += ''
				break
		}
		return strSlideXml
	}

	slide._slideObjects.forEach((slideItemObj: SlideObject, idx: number) => {
		strSlideXml += renderSlideObjectXml(slideItemObj, idx)
	})

	// STEP 4: Add slide numbers (if any) last
	if (slide._slideNumberProps) {
		// Set some defaults (done here b/c SlideNumber canbe added to masters or slides and has numerous entry points)
		if (!slide._slideNumberProps.align) slide._slideNumberProps.align = 'left'

		// Allocate this placeholder's <p:cNvPr> id from the same monotonic counter as every other
		// shape on the slide (top-level objects took `idx + 2`; group children advanced `childIdxAlloc`
		// past that). A hardcoded id here (formerly 25) aliases a shape or group-child id once a slide
		// holds enough objects — a duplicate `<p:cNvPr>` id PowerPoint repairs — so take the next free
		// slot instead. `childIdxAlloc` is the next unused index after the walk above; its id is `+ 2`.
		const slideNumberId = childIdxAlloc++ + 2

		const snProps = slide._slideNumberProps
		strSlideXml += '<p:sp>'
		strSlideXml += ' <p:nvSpPr>'
		strSlideXml +=
			voidEl('p:cNvPr', { id: slideNumberId, name: 'Slide Number Placeholder 0' }, { openPrefix: '  ' }) +
			el('p:cNvSpPr', null, raw(voidEl('a:spLocks', { noGrp: '1' })))
		strSlideXml += el('p:nvPr', null, raw(voidEl('p:ph', { type: 'sldNum', sz: 'quarter', idx: '4294967295' })), {
			openPrefix: '  ',
		})
		strSlideXml += ' </p:nvSpPr>'
		strSlideXml += ' <p:spPr>'
		strSlideXml +=
			el('a:xfrm', null, [
				raw(
					voidEl('a:off', {
						x: getSmartParseNumber(snProps.x, 'X', slide._presLayout),
						y: getSmartParseNumber(snProps.y, 'Y', slide._presLayout),
					})
				),
				raw(
					voidEl('a:ext', {
						cx: snProps.w ? getSmartParseNumber(snProps.w, 'X', slide._presLayout) : '800000',
						cy: snProps.h ? getSmartParseNumber(snProps.h, 'Y', slide._presLayout) : '300000',
					})
				),
			]) +
			el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')), { openPrefix: ' ' }) +
			el(
				'a:extLst',
				null,
				raw(
					el(
						'a:ext',
						{ uri: '{C572A759-6A51-4108-AA02-DFA0A04FC94B}' },
						raw(
							voidEl('ma14:wrappingTextBoxFlag', {
								val: '0',
								'xmlns:ma14': 'http://schemas.microsoft.com/office/mac/drawingml/2011/main',
							})
						)
					)
				),
				{ openPrefix: ' ' }
			) +
			'</p:spPr>'
		strSlideXml += '<p:txBody>'
		// Margins are inches (see `marginToEmu`), matching text-box and cell margins.
		// NOTE: attribute ORDER is byte-significant, and note the margin order is lIns/tIns/rIns/bIns
		// while the source array is [Top, Right, Bottom, Left] — hence the 3/0/1/2 indexing.
		const snMargin = snProps.margin
		const snMarginAt = (arrIdx: number): number | null =>
			Array.isArray(snMargin)
				? marginToEmu(snMargin[arrIdx] || 0)
				: typeof snMargin === 'number'
					? marginToEmu(snMargin || 0)
					: null
		strSlideXml += voidEl('a:bodyPr', {
			lIns: snMarginAt(3),
			tIns: snMarginAt(0),
			rIns: snMarginAt(1),
			bIns: snMarginAt(2),
			anchor: snProps.valign
				? snProps.valign.replace('top', 't').replace('middle', 'ctr').replace('bottom', 'b')
				: null,
		})
		let defRPr = ''
		if (snProps.fontFace || snProps.fontSize || snProps.color) {
			// The typeface is caller-supplied via `slide.slideNumber({ fontFace })`; the element builder
			// escapes it, so a `"`/`&` in the name cannot close the attribute early.
			const face = snProps.fontFace
			defRPr = el('a:defRPr', { sz: clampFontSizeSz(snProps.fontSize || 12) }, [
				snProps.color ? raw(genXmlColorSelection(snProps.color)) : null,
				face ? raw(voidEl('a:latin', { typeface: face })) : null,
				face ? raw(voidEl('a:ea', { typeface: face })) : null,
				face ? raw(voidEl('a:cs', { typeface: face })) : null,
			])
		}
		strSlideXml += el('a:lstStyle', null, raw(el('a:lvl1pPr', null, raw(defRPr))), { openPrefix: '  ' })

		// `align` is normalized to 'left' above when unset; anything not starting c/r falls back to 'l'.
		// `align` is defaulted to 'left' on the props object above; read it through a local so the
		// narrowing survives (the mutation is kept — other readers rely on it).
		const snAlignRaw = snProps.align ?? 'left'
		const snAlign = snAlignRaw.startsWith('c') ? 'ctr' : snAlignRaw.startsWith('r') ? 'r' : 'l'
		strSlideXml += el('a:p', null, [
			raw(voidEl('a:pPr', { algn: snAlign })),
			raw(
				el('a:fld', { id: SLDNUMFLDID, type: 'slidenum' }, [
					// NOTE: `b` is emitted as "0" when unset, unlike the run properties elsewhere which omit it.
					raw(voidEl('a:rPr', { b: snProps.bold ? 1 : 0, lang: 'en-US' })),
					// NOTE: `String(...)` is load-bearing. A slide MASTER has no `_slideNum`, and the
					// template this replaced stringified that `null` into the literal text "null"
					// (visible in slideMaster1.xml). `el()` skips a null child entirely, which would
					// silently drop it — a byte change. Preserved; fixing it is its own commit.
					raw(el('a:t', null, String(slide._slideNum))),
				])
			),
			raw(voidEl('a:endParaRPr', { lang: 'en-US' })),
		])
		strSlideXml += '</p:txBody></p:sp>'
	}

	// STEP 5: Close spTree and finalize slide XML
	strSlideXml += '</p:spTree>'
	strSlideXml += '</p:cSld>'

	// LAST: Return
	return strSlideXml
}

/**
 * Render a `table` slide object to its `<p:graphicFrame>` XML (merge-grid, row/col spans, per-cell styling).
 */
function renderTableObject(
	slideItemObj: SlideObject,
	idx: number,
	x: number,
	y: number,
	cx: number,
	cy: number,
	placeholderObj: SlideObject | null,
	itemOpts: ObjectOptions
): string {
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	let strXml: string
	let arrTabRows: TableCell[][] = []
	let objTabOpts: ObjectOptions = {}
	let intColCnt = 0
	let tblInner = ''
	let cellOpts: TableCellProps | null = null
	// Shallow-clone each row so splice() in the merge-grid builder does not mutate the stored
	// arrTabRows, which would corrupt output on repeated write()/writeFile() calls.
	arrTabRows = (slideItemObj.arrTabRows ?? []).map((row) => [...row])
	objTabOpts = slideItemObj.options
	intColCnt = 0

	// Calc number of columns
	// NOTE: Cells may have a colspan, so merely taking the length of the [0] (or any other) row is not
	// ....: sufficient to determine column count. Therefore, check each cell for a colspan and total cols as reqd
	;(arrTabRows[0] ?? []).forEach((cell) => {
		cellOpts = cell.options || null
		intColCnt += cellOpts?.colspan ? Number(cellOpts.colspan) : 1
	})

	// STEP 1: Start Table XML
	// NOTE: The cNvPr id must be unique among ALL shapes on the slide. A table is an
	// ordinary top-level slide object, so it uses the same `idx + 2` scheme as every other
	// object type below. The legacy `intTableNum * slide._slideNum + 1` formula could collide
	// with another shape's `idx + 2` on the same slide (e.g. a table plus enough sibling
	// shapes on slide 7), producing a duplicate id that makes PowerPoint report the file as
	// corrupt/unreadable (0x80070570) while LibreOffice silently tolerates it.
	strXml =
		'<p:graphicFrame><p:nvGraphicFramePr>' +
		cNvPrOpen(idx + 2, slideItemObj.options.objectName, slideItemObj.options.altText || '') +
		'/>'
	strXml +=
		el(
			'p:cNvGraphicFramePr',
			null,
			raw(
				genXmlObjectLock(
					'a:graphicFrameLocks',
					GRAPHIC_FRAME_LOCK_ATTRS,
					{ noGrp: true, ...slideItemObj.options.objectLock },
					slideItemObj.options.objectName
				)
			)
		) +
		// A table bound to a layout placeholder emits that placeholder's <p:ph> (idx/type) so
		// PowerPoint treats the graphicFrame as filling the placeholder. The <p:ph>
		// precedes <p:extLst> per CT_ApplicationNonVisualDrawingProps document order.
		el(
			'p:nvPr',
			null,
			[
				raw(genXmlPlaceholder(placeholderObj)),
				raw(
					el(
						'p:extLst',
						null,
						raw(
							el(
								'p:ext',
								{ uri: '{D42A27DB-BD31-4B8C-83A1-F6EECF244321}' },
								raw(voidEl('p14:modId', { 'xmlns:p14': P14_NS, val: '1579011935' }))
							)
						)
					)
				),
			],
			{ openPrefix: '  ' }
		) +
		'</p:nvGraphicFramePr>'
	strXml += el('p:xfrm', null, [
		raw(voidEl('a:off', { x: x || (x === 0 ? 0 : EMU_PER_INCH), y: y || (y === 0 ? 0 : EMU_PER_INCH) })),
		raw(voidEl('a:ext', { cx: cx || (cx === 0 ? 0 : EMU_PER_INCH), cy: cy || EMU_PER_INCH })),
	])
	{
		// NOTE: attribute ORDER is byte-significant. None of these flags appears in the byte-gate
		// baseline (zero parts each), so their emission is pinned by test/regression instead.
		const tblPrAttrs: XmlAttrs = {
			rtl: objTabOpts.rtl ? '1' : null,
			firstRow: objTabOpts.hasHeader ? '1' : null,
			lastRow: objTabOpts.hasFooter ? '1' : null,
			bandRow: objTabOpts.hasBandedRows ? '1' : null,
			bandCol: objTabOpts.hasBandedColumns ? '1' : null,
			firstCol: objTabOpts.hasFirstColumn ? '1' : null,
			lastCol: objTabOpts.hasLastColumn ? '1' : null,
		}
		// Paired when a style id is carried, else self-closing — an arity difference.
		const tblPr = objTabOpts.tableStyle
			? el('a:tblPr', tblPrAttrs, raw(el('a:tableStyleId', null, objTabOpts.tableStyle)))
			: voidEl('a:tblPr', tblPrAttrs)
		// The `<a:tbl>` children accumulate here and are wrapped once at STEP 5, so the byte-significant
		// (and non-depth-regular) indentation on the closing tags is described in one place.
		tblInner = tblPr
	}

	// STEP 2: Set column widths
	// Per-column inches from an explicit `colW` array, else split the table's
	// resolved EMU width (`cx`) evenly. `resolveTableColWidthsEmu` is the single
	// source of truth shared with the measured-fit pass. NOTE: divide the EMU
	// width, not the raw inches `options.w` — the latter collapsed auto-width
	// tables to ~0-EMU columns (e.g. `w=9` → `gridCol w="3"`).
	{
		const gridColsEmu = resolveTableColWidthsEmu(objTabOpts.colW, cx, intColCnt)
		tblInner += el(
			'a:tblGrid',
			null,
			gridColsEmu.map((w) => raw(voidEl('a:gridCol', { w })))
		)
	}

	// STEP 3: Build our row arrays into an actual grid to match the XML we will be building next
	// Note row arrays can arrive "lopsided" as in row1:[1,2,3] row2:[3] when first two cols rowspan!,
	// so a simple loop below in XML building wont suffice to build table correctly.
	// We have to build an actual grid now
	/*
					EX: (A0:rowspan=3, B1:rowspan=2, C1:colspan=2)

					/------|------|------|------\
					|  A0  |  B0  |  C0  |  D0  |
					|      |  B1  |  C1  |      |
					|      |      |  C2  |  D2  |
					\------|------|------|------/
				*/
	// A: add _hmerge cell for colspan. should reserve rowspan
	arrTabRows.forEach((cells) => {
		for (let cIdx = 0; cIdx < cells.length;) {
			const cell = cells[cIdx]
			if (!cell) break
			const colspan = cell.options?.colspan
			const rowspan = cell.options?.rowspan
			if (colspan && colspan > 1) {
				const vMergeCells = new Array(colspan - 1).fill(undefined).map(() => {
					return {
						_type: SlideObjectType.tablecell,
						options: { rowspan },
						_hmerge: true,
						_spanOrigin: cell,
					} as const
				})
				cells.splice(cIdx + 1, 0, ...vMergeCells)
				cIdx += colspan
			} else {
				cIdx += 1
			}
		}
	})
	// B: add _vmerge cell for rowspan. should reserve colspan/_hmerge
	arrTabRows.forEach((cells, rIdx) => {
		const nextRow = arrTabRows[rIdx + 1]
		if (!nextRow) return
		cells.forEach((cell, cIdx) => {
			const rowspan = cell._rowContinue || cell.options?.rowspan
			const colspan = cell.options?.colspan
			const _hmerge = cell._hmerge
			if (rowspan && rowspan > 1) {
				// Point back to the true origin cell: when `cell` is itself an `_hmerge` dummy
				// (combined colspan+rowspan), use its origin rather than the dummy.
				const _spanOrigin = cell._spanOrigin || cell
				const hMergeCell = {
					_type: SlideObjectType.tablecell,
					options: { colspan },
					_rowContinue: rowspan - 1,
					_vmerge: true,
					_hmerge,
					_spanOrigin,
				} as const
				nextRow.splice(cIdx, 0, hMergeCell)
			}
		})
	})

	// STEP 4: Build table rows/cells
	arrTabRows.forEach((cells, rIdx) => {
		// A: Table Height provided without rowH? Then distribute rows
		let intRowH = 0 // IMPORTANT: Default must be zero for auto-sizing to work
		if (Array.isArray(objTabOpts.rowH) && objTabOpts.rowH[rIdx]) intRowH = inch2Emu(Number(objTabOpts.rowH[rIdx]))
		else if (objTabOpts.rowH && !isNaN(Number(objTabOpts.rowH))) intRowH = inch2Emu(Number(objTabOpts.rowH))
		else if (itemOpts.cy || itemOpts.h) {
			// `cy` already holds the table height resolved to EMU (line ~276), correctly handling
			// inches/percent/unit-string inputs — reuse it rather than re-parsing options.h.
			intRowH = Math.round((itemOpts.h ? cy : typeof itemOpts.cy === 'number' ? itemOpts.cy : 1) / arrTabRows.length)
		}

		// B: Start row — cells accumulate here and the row wraps them once, below.
		const rowCells: string[] = []

		// C: Loop over each CELL
		cells.forEach((cellObj) => {
			const cell: TableCell = cellObj

			// NOTE: attribute ORDER is byte-significant; `undefined` omits the attribute entirely,
			// which is what the old `.filter(([, v]) => !!v)` did.
			const cellSpanAttrs: XmlAttrs = {
				rowSpan: cell.options?.rowspan && cell.options.rowspan > 1 ? cell.options.rowspan : undefined,
				gridSpan: cell.options?.colspan && cell.options.colspan > 1 ? cell.options.colspan : undefined,
				vMerge: cell._vmerge ? 1 : undefined,
				hMerge: cell._hmerge ? 1 : undefined,
			}

			// 1: COLSPAN/ROWSPAN: Emit the dummy covered cell for any active span. PowerPoint defines a
			// merged region's outer edges (e.g. the right border of a colspan, the bottom border of a
			// rowspan) on the *covered* cells, so inherit the origin cell's border + fill here instead of
			// emitting an empty `<a:tcPr/>` that drops those edges.
			if (cell._hmerge || cell._vmerge) {
				const origin = cell._spanOrigin
				let spanPrXml = ''
				if (origin) {
					const originOpts = origin.options || {}
					const originBorder = Array.isArray(originOpts.border) ? originOpts.border : null
					if (originBorder) spanPrXml += genTableCellBorderXml(originBorder)
					// Resolve the origin's fill with the same precedence the origin cell itself uses below,
					// so the whole merged region fills uniformly.
					const spanFill = originOpts.fill || ''
					if (spanFill) spanPrXml += genXmlColorSelection(spanFill)
				}
				// NOTE: the covered cell is FLAT, unlike the real cell below, which carries indentation
				// before its `</a:tcPr>` and `</a:tc>`.
				rowCells.push(el('a:tc', cellSpanAttrs, raw(el('a:tcPr', null, raw(spanPrXml)))))
				return
			}

			// 2: OPTIONS: Build/set cell options
			const cellOpts = cell.options || {}
			cell.options = cellOpts

			// B: Inherit some options from table when cell options dont exist
			// @see: http://officeopenxml.com/drwTableCellProperties-alignment.php
			const inheritedCellOpts = cellOpts as Partial<Record<TableInheritableOption, TableInheritableValue>>
			const inheritedTableOpts = objTabOpts as Partial<Record<TableInheritableOption, TableInheritableValue>>
			;(
				[
					'align',
					'bold',
					'border',
					'color',
					'fill',
					'fontFace',
					'fontSize',
					'margin',
					'textDirection',
					'underline',
					'valign',
				] as const
			).forEach((name) => {
				if (inheritedTableOpts[name] && !inheritedCellOpts[name] && inheritedCellOpts[name] !== 0)
					inheritedCellOpts[name] = inheritedTableOpts[name]
			})

			const cellValign = cellOpts.valign
				? cellOpts.valign
						.replace(/^c$/i, 'ctr')
						.replace(/^m$/i, 'ctr')
						.replace('center', 'ctr')
						.replace('middle', 'ctr')
						.replace('top', 't')
						.replace('btm', 'b')
						.replace('bottom', 'b')
				: null
			const cellTextDir = cellOpts.textDirection && cellOpts.textDirection !== 'horz' ? cellOpts.textDirection : null

			const fillColor = cellOpts.fill || ''
			const cellFill = fillColor ? genXmlColorSelection(fillColor) : ''

			let cellMargin = cellOpts.margin === 0 || cellOpts.margin ? cellOpts.margin : DEF_CELL_MARGIN_IN
			if (!Array.isArray(cellMargin) && typeof cellMargin === 'number')
				cellMargin = [cellMargin, cellMargin, cellMargin, cellMargin]
			// defensive fallback - if `cellMargin` is not a 4-element array of finite numbers, use defaults (prevents NaN in marL/R/T/B)
			if (
				!Array.isArray(cellMargin) ||
				cellMargin.length !== 4 ||
				cellMargin.some((v) => typeof v !== 'number' || !isFinite(v))
			) {
				cellMargin = DEF_CELL_MARGIN_IN
			}
			// Cell margins are inches (see `marginToEmu`); a `>= 1` value warns once as a likely legacy points value.
			// NOTE: attribute ORDER is byte-significant (margins, then anchor, then vert).
			const tcPrAttrs: XmlAttrs = {
				marL: marginToEmu(cellMargin[3]),
				marR: marginToEmu(cellMargin[1]),
				marT: marginToEmu(cellMargin[0]),
				marB: marginToEmu(cellMargin[2]),
				anchor: cellValign,
				vert: cellTextDir,
			}

			// FUTURE: cell no-wrap support (add `horzOverflow="overflow"` to the cell's `<a:tcPr>`)

			// 4: Set CELL content and properties; 5: borders; 6: fill ==============
			// The trailing indentation before `</a:tcPr>` and `</a:tc>` is byte-significant.
			const cellBorder = Array.isArray(cellOpts.border) ? cellOpts.border : null
			rowCells.push(
				el(
					'a:tc',
					cellSpanAttrs,
					[
						raw(genXmlTextBody(cell)),
						raw(
							el('a:tcPr', tcPrAttrs, [cellBorder ? raw(genTableCellBorderXml(cellBorder)) : null, raw(cellFill)], {
								closePrefix: '  ',
							})
						),
					],
					{ closePrefix: ' ' }
				)
			)
		})

		// D: Complete row
		tblInner += el('a:tr', { h: intRowH }, rowCells.map(raw))
	})

	// STEP 5: Complete table. NOTE: the closing tags carry indentation the opening tags do not,
	// so each `closePrefix` is stated explicitly rather than derived from depth.
	strXml += el(
		'a:graphic',
		null,
		raw(
			el(
				'a:graphicData',
				{ uri: 'http://schemas.openxmlformats.org/drawingml/2006/table' },
				raw(el('a:tbl', null, raw(tblInner), { closePrefix: '      ' })),
				{ closePrefix: '    ' }
			)
		),
		{ closePrefix: '  ' }
	)
	strXml += '</p:graphicFrame>'

	return strXml
}

/**
 * Render a `text` / `placeholder` slide object to its `<p:sp>` XML.
 */
function renderTextObject(
	slideItemObj: SlideObject,
	idx: number,
	slide: PresSlideInternal | SlideLayoutInternal,
	x: number,
	y: number,
	cx: number,
	cy: number,
	placeholderObj: SlideObject | null,
	locationAttrs: XmlAttrs
): string {
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	// Lines can have zero cy, but text should not
	if (!slideItemObj.options.line && cy === 0) cy = EMU_PER_INCH * 0.3

	// Margin/Padding/Inset for textboxes
	if (!slideItemObj.options._bodyProp) slideItemObj.options._bodyProp = {}
	if (slideItemObj.options.margin && Array.isArray(slideItemObj.options.margin)) {
		// Margin arrays are documented as [Top, Right, Bottom, Left] (CSS order) and table cells /
		// slide numbers already map them that way. Keep textboxes consistent: index 0=Top, 3=Left.
		// Margins are inches (see `marginToEmu`), matching cell margins and the PowerPoint dialog.
		slideItemObj.options._bodyProp.tIns = marginToEmu(slideItemObj.options.margin[0] || 0)
		slideItemObj.options._bodyProp.rIns = marginToEmu(slideItemObj.options.margin[1] || 0)
		slideItemObj.options._bodyProp.bIns = marginToEmu(slideItemObj.options.margin[2] || 0)
		slideItemObj.options._bodyProp.lIns = marginToEmu(slideItemObj.options.margin[3] || 0)
	} else if (typeof slideItemObj.options.margin === 'number') {
		slideItemObj.options._bodyProp.lIns = marginToEmu(slideItemObj.options.margin)
		slideItemObj.options._bodyProp.rIns = marginToEmu(slideItemObj.options.margin)
		slideItemObj.options._bodyProp.bIns = marginToEmu(slideItemObj.options.margin)
		slideItemObj.options._bodyProp.tIns = marginToEmu(slideItemObj.options.margin)
	}

	// A: Start SHAPE =======================================================
	// A native equation uses the `a14` (drawing-2010) markup-compatibility extension.
	// PowerPoint wraps the whole shape in <mc:AlternateContent><mc:Choice Requires="a14"> so
	// non-a14 consumers (and schema validators) treat the a14:m subtree as a known extension.
	if (objectHasMath(slideItemObj)) {
		strSlideXml += '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
		strSlideXml += '<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
	}
	strSlideXml += '<p:sp>'

	// B: The addition of the "txBox" attribute is the sole determiner of if an object is a shape or textbox
	const txtOpts = slideItemObj.options
	strSlideXml +=
		'<p:nvSpPr>' +
		cNvPrOpen(idx + 2, txtOpts.objectName, txtOpts.altText || '') +
		'>' +
		cNvPrHyperlink(txtOpts.hyperlink) +
		'</p:cNvPr>'
	{
		const spLockXml = genXmlObjectLock('a:spLocks', SHAPE_LOCK_ATTRS, txtOpts.objectLock, txtOpts.objectName)
		// NOTE: paired only when there are locks to carry; otherwise self-closing. That is an arity
		// difference, so it cannot be expressed as one `el()` call.
		const cNvSpPrAttrs: XmlAttrs = { txBox: txtOpts?.isTextBox ? '1' : null }
		strSlideXml += spLockXml ? el('p:cNvSpPr', cNvSpPrAttrs, raw(spLockXml)) : voidEl('p:cNvSpPr', cNvSpPrAttrs)
	}
	// Prefer the resolved slide-layout placeholder; otherwise fall back to the shape's own
	// placeholder type so a standalone title/body text box still emits a real <p:ph>.
	strSlideXml += el(
		'p:nvPr',
		null,
		raw(
			genXmlPlaceholder(
				slideItemObj._type === SlideObjectType.placeholder ||
					(placeholderObj == null && slideItemObj.options?._placeholderType)
					? slideItemObj
					: placeholderObj
			)
		)
	)
	strSlideXml += '</p:nvSpPr><p:spPr>'
	strSlideXml += el('a:xfrm', locationAttrs, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])

	if (slideItemObj.shape === 'custGeom') {
		strSlideXml += genXmlCustGeom(slideItemObj.options, cx, cy, slide._presLayout)
	} else {
		strSlideXml += genXmlPresetGeom(slideItemObj.shape ?? '', slideItemObj.options, cx, cy)
	}

	// Option: FILL
	strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>'

	// shape Type: LINE: line color
	if (slideItemObj.options.line) strSlideXml += genXmlShapeLine(slideItemObj.options.line)

	// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
	if (slideItemObj.options.shadow && slideItemObj.options.shadow.type !== 'none') {
		strSlideXml += createShadowEffectLst(slideItemObj.options.shadow, DEF_TEXT_SHADOW)
	}

	// B: Close shape Properties
	strSlideXml += '</p:spPr>'

	// C: Add formatted text (text body "bodyPr")
	strSlideXml += genXmlTextBody(slideItemObj)

	// LAST: Close SHAPE =======================================================
	strSlideXml += '</p:sp>'

	// Close the a14 markup-compatibility envelope for an equation-bearing shape.
	if (objectHasMath(slideItemObj)) strSlideXml += '</mc:Choice></mc:AlternateContent>'
	return strSlideXml
}

/**
 * Render a `connector` slide object to its `<p:cxnSp>` XML (start/end shape bindings via shapeIds).
 */
function renderConnectorObject(
	slideItemObj: SlideObject,
	idx: number,
	x: number,
	y: number,
	cx: number,
	cy: number,
	locationAttrs: XmlAttrs,
	shapeIds: Map<SlideObject, number>
): string {
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	// A connector is emitted as <p:cxnSp> (a connector shape) rather than <p:sp>, so
	// PowerPoint treats it as a connector. Geometry/flip come from the shared resolution
	// above; the preset (straightConnector1 / bentConnector3 / curvedConnector3) is on `shape`.
	strSlideXml += '<p:cxnSp><p:nvCxnSpPr>'
	strSlideXml += cNvPrOpen(idx + 2, slideItemObj.options.objectName, slideItemObj.options.altText || '') + '/>'
	{
		// Shape binding: resolve each bound target's objectName to its cNvPr id and emit
		// <a:stCxn>/<a:endCxn> in schema order. Resolution goes through `shapeIds`, so a shape
		// inside a group binds like any other (it is cNvPr-named on this slide); the old
		// `_slideObjects`-only lookup missed those and warned that the shape did not exist.
		// An unresolved name falls back to the static endpoint geometry (warn, don't corrupt)
		// rather than a dangling id.
		const cxnTag = (binding: { name: string; idx: number } | undefined, tag: 'a:stCxn' | 'a:endCxn'): string => {
			if (!binding) return ''
			const id = resolveObjectNameToId(shapeIds, binding.name)
			if (id === null) {
				warn(
					`addConnector could not bind to shape "${binding.name}" (no object with that objectName on the slide); using endpoint coordinates instead.`
				)
				return ''
			}
			return `<${tag} id="${id}" idx="${binding.idx}"/>`
		}
		const cxnSpPr = cxnTag(slideItemObj.options._startCxn, 'a:stCxn') + cxnTag(slideItemObj.options._endCxn, 'a:endCxn')
		strSlideXml += cxnSpPr ? `<p:cNvCxnSpPr>${cxnSpPr}</p:cNvCxnSpPr>` : '<p:cNvCxnSpPr/>'
	}
	strSlideXml += '<p:nvPr/></p:nvCxnSpPr><p:spPr>'
	strSlideXml += el('a:xfrm', locationAttrs, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])
	{
		// Bent/curved connectors carry adjustable jogs as `<a:gd name="adjN" fmla="val …"/>`
		// (1000ths-of-a-percent). With none, the empty `<a:avLst/>` leaves the preset default (50%).
		const adj = slideItemObj.options._connectorAdj || []
		const avLst = adj.map((val, i) => voidEl('a:gd', { name: `adj${i + 1}`, fmla: `val ${val}` })).join('')
		strSlideXml += el('a:prstGeom', { prst: slideItemObj.shape }, raw(el('a:avLst', null, raw(avLst))))
	}
	strSlideXml += genXmlShapeLine(slideItemObj.options.line || {})
	strSlideXml += '</p:spPr></p:cxnSp>'
	return strSlideXml
}

/**
 * Render an `image` slide object to its `<p:pic>` XML (sizing/crop, rounding, hyperlink, shadow).
 */
function renderImageObject(
	slideItemObj: SlideObject,
	idx: number,
	slide: PresSlideInternal | SlideLayoutInternal,
	x: number,
	y: number,
	cx: number,
	cy: number,
	imgWidth: number,
	imgHeight: number,
	placeholderObj: SlideObject | null,
	locationAttrs: XmlAttrs,
	sizing: ObjectOptions['sizing'],
	rounding: ObjectOptions['rounding']
): string {
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	// Backfill any omitted dimension of a path-based image from its natural pixel ratio.
	// The bytes weren't available synchronously in `addImage()`, but `_relsMedia[].data` is
	// populated by now, so measure it here and keep aspect ratio.
	// PowerPoint inserts images at 96 DPI, so natural pixels / 96 * EMU == display EMU.
	if (slideItemObj.options._szAuto) {
		const szAuto = slideItemObj.options._szAuto
		const relData = (slide._relsMedia || []).find((rel) => rel.rId === slideItemObj.imageRid)?.data
		const natural = typeof relData === 'string' ? getImageSizeFromBase64(relData) : null
		if (natural) {
			if (szAuto.w && szAuto.h) {
				cx = pixelsToEmu(natural.w, 96)
				cy = pixelsToEmu(natural.h, 96)
			} else if (szAuto.h) {
				// Width supplied, derive height
				cy = Math.round(cx * (natural.h / natural.w))
			} else if (szAuto.w) {
				// Height supplied, derive width
				cx = Math.round(cy * (natural.w / natural.h))
			}
			imgWidth = cx
			imgHeight = cy
		}
	}
	const imgOpts = slideItemObj.options
	const imgLink = slideItemObj.hyperlink
	strSlideXml += '<p:pic>'
	strSlideXml += '  <p:nvPicPr>'
	strSlideXml +=
		cNvPrOpen(idx + 2, imgOpts.objectName, imgOpts.altText || slideItemObj.image || '') + '>' + cNvPrHyperlink(imgLink)
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
			? raw(
					voidEl(
						'a:alphaModFix',
						{ amt: Math.round((100 - imgOpts.transparency) * FIXED_PCT_PER_PERCENT) },
						{ openPrefix: alphaPrefix }
					)
				)
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
			? raw(voidEl('a:biLevel', { thresh: Math.round(imgOpts.biLevel.threshold * PERCENT_SCALE) }))
			: null,
	]

	strSlideXml += '<p:blipFill>'
	// NOTE: This works for both cases: either `path` or `data` contains the SVG
	if ((slide._relsMedia || []).find((rel) => rel.rId === slideItemObj.imageRid)?.extn === 'svg') {
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
										'xmlns:asvg': 'http://schemas.microsoft.com/office/drawing/2016/SVG/main',
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
	if (slideItemObj.options.crop) {
		// Explicit OOXML srcRect (percentage edge insets), emitted verbatim. Crops the source
		// directly, so it wins over the inch-based `sizing` crop and works for SVG/unmeasurable
		// formats; the picture's normal w/h box stays the display extent.
		if (sizing?.type)
			warn(
				`addImage 'crop' and 'sizing' are mutually exclusive for image "${slideItemObj.options.objectName}"; 'sizing' was ignored.`
			)
		strSlideXml += genXmlImageCrop(slideItemObj.options.crop, slideItemObj.options.objectName)
	} else if (sizing?.type) {
		const boxW = sizing.w ? getSmartParseNumber(sizing.w, 'X', slide._presLayout) : cx
		const boxH = sizing.h ? getSmartParseNumber(sizing.h, 'Y', slide._presLayout) : cy
		const boxX = getSmartParseNumber(sizing.x || 0, 'X', slide._presLayout)
		const boxY = getSmartParseNumber(sizing.y || 0, 'Y', slide._presLayout)

		// `cover`/`contain` crop the *source* bitmap, so the srcRect must be derived from the
		// image's natural pixel ratio — not the displayed box (options.w/h). Measure it from the
		// embedded media bytes; if unmeasurable (SVG/unknown format) fall back to display dims + warn.
		// `crop` keeps display EMU: its contract treats the displayed extent as the crop frame.
		let cropSize: { w: number; h: number } = { w: imgWidth, h: imgHeight }
		if (sizing.type === 'cover' || sizing.type === 'contain') {
			const relData = (slide._relsMedia || []).find((rel) => rel.rId === slideItemObj.imageRid)?.data
			const natural = typeof relData === 'string' ? getImageSizeFromBase64(relData) : null
			if (natural) {
				cropSize = natural
			} else {
				warn(
					`sizing '${sizing.type}' could not measure natural dimensions for image "${slideItemObj.options.objectName}"; falling back to displayed aspect ratio (crop may be inexact). Provide a raster image (PNG/JPEG/GIF/BMP/WebP) or an SVG with width/height or a viewBox to enable an aspect-correct crop.`
				)
			}
		}

		strSlideXml += ImageSizingXml[sizing.type](cropSize, { w: boxW, h: boxH, x: boxX, y: boxY })
		imgWidth = boxW
		imgHeight = boxH
	} else {
		strSlideXml += el('a:stretch', null, raw(voidEl('a:fillRect')), { openPrefix: '  ' })
	}
	strSlideXml += '</p:blipFill>'
	strSlideXml += '<p:spPr>'
	strSlideXml += el(
		'a:xfrm',
		locationAttrs,
		[raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx: imgWidth, cy: imgHeight }))],
		{ openPrefix: ' ', childPrefix: '  ', closePrefix: ' ' }
	)
	// Clip the picture to a geometry. `points` (freeform custGeom) takes precedence over `shape`/`rounding`;
	// otherwise `shape` wins over `rounding` (shorthand for an ellipse), falling back to a plain rectangle.
	if (slideItemObj.options.points) {
		strSlideXml += ' ' + genXmlCustGeom(slideItemObj.options, imgWidth, imgHeight, slide._presLayout)
	} else {
		strSlideXml +=
			' ' +
			genXmlPresetGeom(
				slideItemObj.options.shape ?? (rounding ? 'ellipse' : 'rect'),
				slideItemObj.options,
				imgWidth,
				imgHeight
			)
	}

	// BORDER: `<a:ln>` outline (must precede `<a:effectLst>` per CT_ShapeProperties order)
	if (slideItemObj.options.line) strSlideXml += genXmlShapeLine(slideItemObj.options.line)

	// EFFECTS > SHADOW: REF: @see http://officeopenxml.com/drwSp-effects.php
	if (slideItemObj.options.shadow && slideItemObj.options.shadow.type !== 'none') {
		strSlideXml += createShadowEffectLst(slideItemObj.options.shadow, DEF_TEXT_SHADOW)
	}
	strSlideXml += '</p:spPr>'
	strSlideXml += '</p:pic>'
	return strSlideXml
}

/**
 * Render a `media` (audio/video/online) slide object to its `<p:pic>` XML.
 */
function renderMediaObject(
	slideItemObj: SlideObject,
	idx: number,
	x: number,
	y: number,
	cx: number,
	cy: number,
	locationAttrs: XmlAttrs
): string {
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	const opts = slideItemObj.options
	const mediaRid = slideItemObj.mediaRid ?? 0

	// The online (external-link) and embedded forms differ in exactly two tokens, so they share
	// one body rather than two near-identical branches:
	//  - EG_Media choice: audio embeds use <a:audioFile>, video and online video <a:videoFile>;
	//  - the paired <p14:media> rel is bound by r:link for an external link (no media binary
	//    part, sharing the link Target with the ECMA rel) and by r:embed for an embedded file.
	// Note the ECMA media element itself is r:link in BOTH forms.
	const mediaTag = slideItemObj.mtype === 'audio' ? 'a:audioFile' : 'a:videoFile'
	const p14MediaAttr = slideItemObj.mtype === 'online' ? 'r:link' : 'r:embed'

	return el('p:pic', null, [
		raw(
			el(
				'p:nvPicPr',
				null,
				[
					// cNvPr/@id must be unique across every shape on the slide, so it uses the slide-object
					// index (idx + 2) like all other shapes — NOT mediaRid, which lives in the relationship-id
					// space and collides with a sibling shape's idx (duplicate ids => PowerPoint reports the
					// file corrupt, 0x80070570). The preview image is still bound via <a:blip r:embed> below.
					raw(
						cNvPrOpen(idx + 2, opts.objectName, opts.altText || '') +
							'>' +
							voidEl('a:hlinkClick', { 'r:id': '', action: 'ppaction://media' }) +
							'</p:cNvPr>'
					),
					raw(
						el(
							'p:cNvPicPr',
							null,
							raw(
								genXmlObjectLock(
									'a:picLocks',
									PICTURE_LOCK_ATTRS,
									{ noChangeAspect: true, ...opts.objectLock },
									opts.objectName
								)
							),
							{ openPrefix: ' ' }
						)
					),
					raw(
						el(
							'p:nvPr',
							null,
							[
								raw(voidEl(mediaTag, { 'r:link': `rId${mediaRid}` }, { openPrefix: '  ' })),
								raw(
									el(
										'p:extLst',
										null,
										raw(
											el(
												'p:ext',
												{ uri: '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}' },
												raw(
													voidEl(
														'p14:media',
														{ 'xmlns:p14': P14_NS, [p14MediaAttr]: `rId${mediaRid + 1}` },
														{ openPrefix: '    ' }
													)
												),
												{ openPrefix: '   ', closePrefix: '   ' }
											)
										),
										{ openPrefix: '  ', closePrefix: '  ' }
									)
								),
							],
							{ openPrefix: ' ', closePrefix: ' ' }
						)
					),
				],
				{ openPrefix: ' ', closePrefix: ' ' }
			)
		),
		// NOTE: Preview image is required!
		raw(
			el(
				'p:blipFill',
				null,
				[
					raw(voidEl('a:blip', { 'r:embed': `rId${mediaRid + 2}` })),
					raw(el('a:stretch', null, raw(voidEl('a:fillRect')))),
				],
				{ openPrefix: ' ' }
			)
		),
		raw(
			el(
				'p:spPr',
				null,
				[
					raw(
						el('a:xfrm', locationAttrs, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))], {
							openPrefix: '  ',
						})
					),
					raw(el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')), { openPrefix: '  ' })),
				],
				{ openPrefix: ' ', closePrefix: ' ' }
			)
		),
	])
}

/**
 * Render an `oleObject` slide object to its `<p:graphicFrame>` XML.
 *
 * Mirrors what PowerPoint authors for Insert ▸ Object: a graphicFrame in the `.../ole` graphicData
 * namespace holding an `<mc:AlternateContent>`. The `mc:Choice` carries the bare `<p:oleObj>` (which
 * PowerPoint renders by drawing the live embedded document); the `mc:Fallback` repeats it with a
 * `<p:pic>` preview so every other consumer shows the cached picture.
 *
 * Note `mc:Choice Requires="v"` only *declares* the VML namespace — modern PowerPoint writes no
 * `spid` and no `vmlDrawing` part, so neither is emitted here.
 */
function renderOleObject(slideItemObj: SlideObject, idx: number, x: number, y: number, cx: number, cy: number): string {
	const ole = slideItemObj.ole
	if (!ole) return ''
	const opts = slideItemObj.options || {}

	// Shared by both branches; attribute order matches PowerPoint's.
	const oleAttrs: XmlAttrs = {
		name: ole.name,
		showAsIcon: ole.showAsIcon ? '1' : null,
		'r:id': `rId${ole.objectRid}`,
		imgW: ole.imgW ?? cx,
		imgH: ole.imgH ?? cy,
		progId: ole.progId,
	}
	const oleObj = (children: XmlChild[]): string => el('p:oleObj', oleAttrs, children)

	// The Fallback's cached picture. Its `cNvPr` is the id-less `0`/`""` form PowerPoint writes:
	// the picture is an alternate rendering of the graphicFrame above, never a sibling shape, so it
	// takes no id from the slide's shape-id space.
	const previewPic = el('p:pic', null, [
		raw(
			el('p:nvPicPr', null, [
				raw(voidEl('p:cNvPr', { id: 0, name: '' })),
				raw(voidEl('p:cNvPicPr')),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(
			el('p:blipFill', null, [
				raw(voidEl('a:blip', { 'r:embed': `rId${ole.previewRid}` })),
				raw(el('a:stretch', null, raw(voidEl('a:fillRect')))),
			])
		),
		raw(
			el('p:spPr', null, [
				raw(el('a:xfrm', null, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])),
				raw(el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')))),
			])
		),
	])

	const alternateContent = el('mc:AlternateContent', { 'xmlns:mc': MC_NS }, [
		raw(el('mc:Choice', { 'xmlns:v': VML_NS, Requires: 'v' }, raw(oleObj([raw(voidEl('p:embed'))])))),
		raw(el('mc:Fallback', null, raw(oleObj([raw(voidEl('p:embed')), raw(previewPic)])))),
	])

	return el('p:graphicFrame', null, [
		raw(
			el('p:nvGraphicFramePr', null, [
				raw(cNvPrOpen(idx + 2, opts.objectName, opts.altText || '') + '/>'),
				raw(
					el(
						'p:cNvGraphicFramePr',
						null,
						raw(
							genXmlObjectLock(
								'a:graphicFrameLocks',
								GRAPHIC_FRAME_LOCK_ATTRS,
								{ noChangeAspect: true, ...opts.objectLock },
								opts.objectName
							)
						)
					)
				),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(el('p:xfrm', null, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])),
		raw(el('a:graphic', null, raw(el('a:graphicData', { uri: OLE_NS }, raw(alternateContent))))),
	])
}

/**
 * Render a `chart` slide object to its `<p:graphicFrame>` XML referencing the chart part.
 *
 * Classic charts emit a bare `<p:graphicFrame>` pointing at a `<c:chart>`. chartEx charts
 * (waterfall, …) instead wrap that graphicFrame in `<mc:AlternateContent>`: an `<mc:Choice>`
 * carrying the `<cx:chart>` reference (rendered by PowerPoint 2016+/Microsoft 365) and an
 * `<mc:Fallback>` placeholder shape shown by every other consumer.
 */
function renderChartObject(
	slideItemObj: SlideObject,
	idx: number,
	x: number,
	y: number,
	cx: number,
	cy: number,
	placeholderObj: SlideObject | null
): string {
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	const opts = slideItemObj.options
	// A chart's options are really `ChartOptsInternal`; `_type` is not on the broad `ObjectOptions`.
	const chartType = (opts as { _type?: ChartType })._type
	const isChartEx = isChartExType(chartType)

	const graphicDataUri = isChartEx ? CHARTEX_NS : CHART_NS
	const chartChild = isChartEx
		? voidEl('cx:chart', { 'xmlns:cx': CHARTEX_NS, 'r:id': `rId${slideItemObj.chartRid}` }, { openPrefix: '   ' })
		: voidEl('c:chart', { 'r:id': `rId${slideItemObj.chartRid}`, 'xmlns:c': CHART_NS }, { openPrefix: '   ' })

	const graphicFrame = el('p:graphicFrame', null, [
		raw(
			el(
				'p:nvGraphicFramePr',
				null,
				[
					raw(cNvPrOpen(idx + 2, opts.objectName, opts.altText || '', '   ') + '/>'),
					raw(voidEl('p:cNvGraphicFramePr', null, { openPrefix: '   ' })),
					raw(el('p:nvPr', null, raw(genXmlPlaceholder(placeholderObj)), { openPrefix: '   ' })),
				],
				{ openPrefix: ' ', closePrefix: ' ' }
			)
		),
		raw(el('p:xfrm', null, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))], { openPrefix: ' ' })),
		raw(
			el(
				'a:graphic',
				{ 'xmlns:a': DML_NS },
				raw(el('a:graphicData', { uri: graphicDataUri }, raw(chartChild), { openPrefix: '  ', closePrefix: '  ' })),
				{ openPrefix: ' ', closePrefix: ' ' }
			)
		),
	])

	if (!isChartEx) return graphicFrame

	// chartEx: wrap the graphicFrame in <mc:AlternateContent>. The Choice declares the feature-level
	// namespace it Requires; the Fallback is a plain shape so non-2016 consumers show something.
	const feature = (chartType && CHARTEX_FEATURE_NS[chartType]) ?? CHARTEX_FEATURE_NS[ChartType.waterfall]
	const choice = el(
		'mc:Choice',
		{ [`xmlns:${feature?.prefix ?? 'cx1'}`]: feature?.uri, Requires: feature?.prefix ?? 'cx1' },
		raw(graphicFrame)
	)
	const fallback = el('mc:Fallback', null, raw(renderChartExFallback(idx, opts, x, y, cx, cy)))
	return el('mc:AlternateContent', { 'xmlns:mc': MC_NS }, [raw(choice), raw(fallback)])
}

/**
 * Build the `<mc:Fallback>` placeholder shape shown when a chartEx chart cannot be rendered
 * (any consumer older than PowerPoint 2016 / Microsoft 365). A light-grey rectangle at the chart's
 * position carrying a short explanatory note — enough that the slide reads sensibly rather than
 * showing a void where the chart would be.
 */
function renderChartExFallback(idx: number, opts: ObjectOptions, x: number, y: number, cx: number, cy: number): string {
	return el('p:sp', null, [
		raw(
			el('p:nvSpPr', null, [
				raw(cNvPrOpen(idx + 2, opts.objectName, opts.altText || '') + '/>'),
				raw(voidEl('p:cNvSpPr')),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(
			el('p:spPr', null, [
				raw(el('a:xfrm', null, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])),
				raw(el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')))),
				raw(el('a:solidFill', null, raw(voidEl('a:srgbClr', { val: 'F2F2F2' })))),
				raw(el('a:ln', null, raw(el('a:solidFill', null, raw(voidEl('a:srgbClr', { val: 'BFBFBF' })))))),
			])
		),
		raw(
			el('p:txBody', null, [
				raw(voidEl('a:bodyPr')),
				raw(voidEl('a:lstStyle')),
				raw(
					el('a:p', null, [
						raw(
							el('a:r', null, [
								raw(voidEl('a:rPr', { lang: 'en-US' })),
								raw(el('a:t', null, 'This chart requires PowerPoint 2016 or newer to display.')),
							])
						),
					])
				),
			])
		),
	])
}

/** The `p166:blipFill` + `p166:spPr` (preview image + framed tile) shared by a zoom tile's `zmPr`. */
function zoomBlipSpPr(previewRid: number, xf: { x: number; y: number; cx: number; cy: number }): string {
	return (
		el('p166:blipFill', { 'xmlns:p166': P166_NS }, [
			raw(voidEl('a:blip', { 'r:embed': `rId${previewRid}` })),
			raw(el('a:stretch', null, raw(voidEl('a:fillRect')))),
		]) +
		el('p166:spPr', { 'xmlns:p166': P166_NS }, [
			raw(
				el('a:xfrm', null, [raw(voidEl('a:off', { x: xf.x, y: xf.y })), raw(voidEl('a:ext', { cx: xf.cx, cy: xf.cy }))])
			),
			raw(el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')))),
			raw(el('a:ln', { w: '3175' }, raw(el('a:solidFill', null, raw(voidEl('a:prstClr', { val: 'ltGray' })))))),
		])
	)
}

/** One `{prefix}:{obj}` targeting element (sldZmObj/sectionZmObj/summaryZmObj) with its nested `zmPr`. */
function zoomObjEl(
	variant: ZoomInternal['variant'],
	tile: ZoomTileInternal,
	returnToParent: boolean,
	transitionDur: number,
	localXf: { x: number; y: number; cx: number; cy: number }
): string {
	const { prefix, obj } = ZOOM_VARIANTS[variant]
	const objAttrs: XmlAttrs = variant === 'slide' ? { sldId: tile.sldId ?? null } : { sectionId: tile.sectionId ?? null }
	// zmPr attrs: id, (Slide Zoom only) returnToParent, transitionDur — matching PowerPoint's order.
	const zmAttrs: XmlAttrs = { id: tile.zmPrId }
	if (variant === 'slide') zmAttrs['returnToParent'] = returnToParent ? '1' : '0'
	zmAttrs['transitionDur'] = transitionDur
	return el(
		`${prefix}:${obj}`,
		objAttrs,
		raw(el(`${prefix}:zmPr`, zmAttrs, raw(zoomBlipSpPr(tile.previewRid, localXf))))
	)
}

/** The `mc:Fallback` picture (or grouped pictures) — a hyperlinked thumbnail for pre-2016 consumers. */
function zoomFallbackPic(
	picId: number,
	objectName: string | undefined,
	tile: ZoomTileInternal,
	absXf: { x: number; y: number; cx: number; cy: number }
): string {
	return el('p:pic', null, [
		raw(
			el('p:nvPicPr', null, [
				raw(
					cNvPrOpen(picId, objectName, '') +
						'>' +
						voidEl('a:hlinkClick', { 'r:id': `rId${tile.fallbackSlideRid}`, action: 'ppaction://hlinksldjump' }) +
						'</p:cNvPr>'
				),
				raw(
					el(
						'p:cNvPicPr',
						null,
						raw(
							voidEl('a:picLocks', {
								noGrp: '1',
								noRot: '1',
								noChangeAspect: '1',
								noMove: '1',
								noResize: '1',
								noEditPoints: '1',
								noAdjustHandles: '1',
								noChangeArrowheads: '1',
								noChangeShapeType: '1',
							})
						)
					)
				),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(
			el('p:blipFill', null, [
				raw(voidEl('a:blip', { 'r:embed': `rId${tile.previewRid}` })),
				raw(el('a:stretch', null, raw(voidEl('a:fillRect')))),
			])
		),
		raw(
			el('p:spPr', null, [
				raw(
					el('a:xfrm', null, [
						raw(voidEl('a:off', { x: absXf.x, y: absXf.y })),
						raw(voidEl('a:ext', { cx: absXf.cx, cy: absXf.cy })),
					])
				),
				raw(el('a:prstGeom', { prst: 'rect' }, raw(voidEl('a:avLst')))),
				raw(el('a:ln', { w: '3175' }, raw(el('a:solidFill', null, raw(voidEl('a:prstClr', { val: 'ltGray' })))))),
			])
		),
	])
}

/**
 * Render a `zoom` slide object to its `<mc:AlternateContent>` XML. The `mc:Choice` carries the real
 * `<p:graphicFrame>` (Slide/Section/Summary Zoom in the 2016 zoom namespaces, read by PowerPoint
 * 2016+); the `mc:Fallback` carries a hyperlinked picture — or, for a Summary Zoom, a group of them —
 * so pre-2016 consumers still get a clickable navigation thumbnail. See `gen/define/zoom.ts`.
 */
function renderZoomObject(
	slideItemObj: SlideObject,
	idx: number,
	x: number,
	y: number,
	cx: number,
	cy: number
): string {
	const zoom = slideItemObj.zoom
	if (!zoom) return ''
	const opts = slideItemObj.options || {}
	const objectName = opts.objectName
	const { uri, prefix, zm } = ZOOM_VARIANTS[zoom.variant]
	const firstTile = zoom.tiles[0]
	if (!firstTile) return '' // every zoom is registered with >= 1 tile; keeps this a total function

	// Choice: the zoom `graphicData` body. Single-tile (slide/section) uses a frame-local 0,0 xfrm;
	// summary lays each tile out at its precomputed grid cell (frame-local EMU).
	const objectsXml =
		zoom.variant === 'summary'
			? zoom.tiles
					.map((t) =>
						raw(zoomObjEl('summary', t, zoom.returnToParent, zoom.transitionDur, t.grid ?? { x: 0, y: 0, cx, cy }))
					)
					.concat(raw(voidEl(`${prefix}:gridLayout`)))
			: [raw(zoomObjEl(zoom.variant, firstTile, zoom.returnToParent, zoom.transitionDur, { x: 0, y: 0, cx, cy }))]

	const graphicFrame = el('p:graphicFrame', null, [
		raw(
			el('p:nvGraphicFramePr', null, [
				raw(cNvPrOpen(idx + 2, objectName, '') + '/>'),
				raw(el('p:cNvGraphicFramePr', null, raw(voidEl('a:graphicFrameLocks', { noChangeAspect: '1' })))),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(el('p:xfrm', null, [raw(voidEl('a:off', { x, y })), raw(voidEl('a:ext', { cx, cy }))])),
		raw(el('a:graphic', null, raw(el('a:graphicData', { uri }, raw(el(`${prefix}:${zm}`, null, objectsXml)))))),
	])
	const choice = el('mc:Choice', { [`xmlns:${prefix}`]: uri, Requires: prefix }, raw(graphicFrame))

	// Fallback: a hyperlinked picture per tile at its slide-absolute position.
	let fallbackInner: string
	if (zoom.variant === 'summary') {
		const pics = zoom.tiles.map((t, k) => {
			const g = t.grid ?? { x: 0, y: 0, cx, cy }
			return raw(zoomFallbackPic(idx + 3 + k, objectName, t, { x: x + g.x, y: y + g.y, cx: g.cx, cy: g.cy }))
		})
		fallbackInner = el('p:grpSp', null, [
			raw(
				el('p:nvGrpSpPr', null, [
					raw(cNvPrOpen(idx + 2, objectName, '') + '/>'),
					raw(voidEl('p:cNvGrpSpPr')),
					raw(voidEl('p:nvPr')),
				])
			),
			raw(
				el(
					'p:grpSpPr',
					null,
					raw(
						el('a:xfrm', null, [
							raw(voidEl('a:off', { x, y })),
							raw(voidEl('a:ext', { cx, cy })),
							raw(voidEl('a:chOff', { x, y })),
							raw(voidEl('a:chExt', { cx, cy })),
						])
					)
				)
			),
			...pics,
		])
	} else {
		fallbackInner = zoomFallbackPic(idx + 2, objectName, firstTile, { x, y, cx, cy })
	}

	return el('mc:AlternateContent', { 'xmlns:mc': MC_NS }, [
		raw(choice),
		raw(el('mc:Fallback', null, raw(fallbackInner))),
	])
}

/**
 * Transforms slide relations to XML string.
 * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
 * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
 * @param {PresSlideInternal | SlideLayoutInternal} slide - slide object whose relations are being transformed
 * @param {{ target: string; type: string }[]} defaultRels - array of default relations
 * @return {string} XML
 */
export function slideObjectRelationsToXml(
	slide: PresSlideInternal | SlideLayoutInternal,
	defaultRels: Array<{ target: string; type: string }>
): string {
	let lastRid = 0 // stores maximum rId used for dynamic relations
	const rels: string[] = []

	/**
	 * Has a rel with this Target already been emitted? Media items produce *TWO* rels
	 * sharing one Target, and the second is told from the first by the Target already
	 * being present. `target` must therefore be the ESCAPED form — what actually got
	 * written — or an online-video link carrying `&` never matches its own first rel
	 * and the pair is mistyped (media emitted as video). See `SlideRel.Target`.
	 */
	const hasTarget = (target: string): boolean => rels.some((xml) => xml.includes(` Target="${target}"`))

	// STEP 1: Add all rels for this Slide
	slide._rels.forEach((rel: SlideRel) => {
		lastRid = Math.max(lastRid, rel.rId)
		if (isHyperlinkRel(rel)) {
			if (rel.data === 'slide') {
				rels.push(
					voidEl('Relationship', { Id: `rId${rel.rId}`, Type: OFFICE_REL + 'slide', Target: `slide${rel.Target}.xml` })
				)
			} else {
				rels.push(
					voidEl('Relationship', {
						Id: `rId${rel.rId}`,
						Type: OFFICE_REL + 'hyperlink',
						Target: rel.Target,
						TargetMode: 'External',
					})
				)
			}
		}
		// NOTE: there is no `else` here on purpose. `_rels` only ever holds hyperlinks —
		// it is initialized empty (`slide.ts`) and the three sites that append to it
		// (`define/hyperlinks.ts` ×2, `define/image.ts`) all push `SlideObjectType.hyperlink`.
		// The notesSlide rel every slide needs comes from `defaultRels` in `makeXmlSlideRel`.
	})
	;(slide._relsChart || []).forEach((rel: SlideRelChart) => {
		lastRid = Math.max(lastRid, rel.rId)
		// chartEx parts use the MS chartEx rel type; classic charts use the ECMA `chart` rel.
		rels.push(
			voidEl('Relationship', {
				Id: `rId${rel.rId}`,
				Type: rel.isChartEx ? CHARTEX_REL : OFFICE_REL + 'chart',
				Target: rel.Target,
			})
		)
	})
	;(slide._relsMedia || []).forEach((rel: SlideRelMedia) => {
		const relType = rel.type.toLowerCase()
		// `voidEl` escapes the Target on the way out; the probe has to compare against
		// those emitted bytes, so it needs the escaped form computed separately here.
		const relTarget = encodeXmlEntities(rel.Target)
		const media = (type: string, targetMode?: string): string =>
			voidEl('Relationship', { Id: `rId${rel.rId}`, Type: type, Target: rel.Target, TargetMode: targetMode })
		lastRid = Math.max(lastRid, rel.rId)
		if (rel.oleRelType) {
			// An OLE payload part carries its rel type verbatim (`.../package` or `.../oleObject`);
			// its `type` is the part's content type, which the sniffing below would misread.
			rels.push(media(rel.oleRelType))
		} else if (relType.includes('image')) {
			rels.push(media(OFFICE_REL + 'image'))
		} else if (relType.includes('audio')) {
			rels.push(hasTarget(relTarget) ? media(MS_MEDIA_REL) : media(OFFICE_REL + 'audio'))
		} else if (relType.includes('video')) {
			rels.push(hasTarget(relTarget) ? media(MS_MEDIA_REL) : media(OFFICE_REL + 'video'))
		} else if (relType.includes('online')) {
			// Online video has *TWO* external rels sharing the link Target: the ECMA video
			// rel (first) and the MS-2007 media rel (second). Both TargetMode="External",
			// no media binary part.
			rels.push(hasTarget(relTarget) ? media(MS_MEDIA_REL, 'External') : media(OFFICE_REL + 'video', 'External'))
		}
	})

	// STEP 2: Add default rels
	defaultRels.forEach((rel, idx) => {
		rels.push(voidEl('Relationship', { Id: `rId${lastRid + idx + 1}`, Type: rel.type, Target: rel.target }))
	})

	return XML_DECL + CRLF + el('Relationships', { xmlns: PACKAGE_REL_NS }, rels.map(raw))
}
