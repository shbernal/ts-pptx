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

import {
	CRLF,
	DEF_CELL_MARGIN_IN,
	DEF_PRES_LAYOUT_NAME,
	DEF_TEXT_SHADOW,
	EMU,
	SLDNUMFLDID,
	SlideObjectType,
	XML_DECL,
} from '../../core-enums.js'
import type {
	ObjectOptions,
	PresSlideInternal,
	SlideLayoutInternal,
	SlideObject,
	SlideRel,
	SlideRelChart,
	SlideRelMedia,
	TableCell,
	TableCellProps,
} from '../../core-interfaces.js'
import {
	convertRotationDegrees,
	createColorElement,
	createLineCap,
	createShadowEffectLst,
	encodeXmlEntities,
	genXmlColorSelection,
	genXmlLineFill,
	getDuplicateObjectNames,
	getImageSizeFromBase64,
	getSmartParseNumber,
	inch2Emu,
	isHyperlinkRel,
	lineWidthToEmu,
	marginToEmu,
	resolveTableColWidthsEmu,
} from '../../gen-utils.js'
import { FIXED_PCT_PER_PERCENT, pixelsToEmu } from '../../units.js'
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
import { el, raw, voidEl } from '../oxml/el.js'
import { collectSlideShapeIds, resolveObjectNameToId } from './shape-ids.js'

const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
/** The MS-2007 `media` rel that pairs with an ECMA audio/video/online rel on the same Target. */
const MS_MEDIA_REL = 'http://schemas.microsoft.com/office/2007/relationships/media'

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
	let strSlideXml: string = slide._name ? '<p:cSld name="' + slide._name + '">' : '<p:cSld>'

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
		strSlideXml += `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId${slide._bkgdImgRid}"><a:lum/></a:blip><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`
	} else if (slide.background?.color || slide.background?.type === 'gradient') {
		strSlideXml += `<p:bg><p:bgPr>${genXmlColorSelection(slide.background)}<a:effectLst/></p:bgPr></p:bg>`
	} else if (!slide.background && slide._name && slide._name === DEF_PRES_LAYOUT_NAME) {
		// NOTE: Default [white] background is needed on slideMaster1.xml to avoid gray background in Keynote (and Finder previews)
		strSlideXml += '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
	}

	// STEP 2: Continue slide by starting spTree node
	// spTree root — OOXML requires the shape tree to open with the implicit top-level group's
	// non-visual props (`<p:nvGrpSpPr>`, the reserved `cNvPr id="1"`) and an identity group
	// transform (off/ext and chOff/chExt all zero) before any child shape. This is the slide's
	// built-in root group, not a user-authored `addGroup` — hence the zeroed frame.
	strSlideXml += '<p:spTree>'
	strSlideXml += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
	strSlideXml += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
	strSlideXml += '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'

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
		let locationAttr = ''
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
		//
		if (slideItemObj.options.flipH) locationAttr += ' flipH="1"'
		if (slideItemObj.options.flipV) locationAttr += ' flipV="1"'
		if (slideItemObj.options.rotate) locationAttr += ` rot="${convertRotationDegrees(slideItemObj.options.rotate)}"`

		// B: Add OBJECT to the current Slide
		switch (slideItemObj._type) {
			case SlideObjectType.table:
				strSlideXml += renderTableObject(slideItemObj, idx, x, y, cx, cy, placeholderObj, itemOpts)
				break
			case SlideObjectType.text:
			case SlideObjectType.placeholder:
				strSlideXml += renderTextObject(slideItemObj, idx, slide, x, y, cx, cy, placeholderObj, locationAttr)
				break
			case SlideObjectType.connector:
				strSlideXml += renderConnectorObject(slideItemObj, idx, x, y, cx, cy, locationAttr, shapeIds)
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
					locationAttr,
					sizing,
					rounding
				)
				break
			case SlideObjectType.media:
				strSlideXml += renderMediaObject(slideItemObj, idx, x, y, cx, cy, locationAttr)
				break
			case SlideObjectType.chart:
				strSlideXml += renderChartObject(slideItemObj, idx, x, y, cx, cy, placeholderObj)
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
				strSlideXml += '<p:nvGrpSpPr>'
				strSlideXml += `<p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}"/>`
				strSlideXml += grpLockXml ? `<p:cNvGrpSpPr>${grpLockXml}</p:cNvGrpSpPr>` : '<p:cNvGrpSpPr/>'
				strSlideXml += '<p:nvPr/>'
				strSlideXml += '</p:nvGrpSpPr>'
				strSlideXml += `<p:grpSpPr><a:xfrm${locationAttr}>`
				strSlideXml += `<a:off x="${gx}" y="${gy}"/><a:ext cx="${gcx}" cy="${gcy}"/>`
				strSlideXml += `<a:chOff x="${gx}" y="${gy}"/><a:chExt cx="${gcx}" cy="${gcy}"/>`
				strSlideXml += '</a:xfrm></p:grpSpPr>'
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

		strSlideXml += '<p:sp>'
		strSlideXml += ' <p:nvSpPr>'
		strSlideXml += `  <p:cNvPr id="${slideNumberId}" name="Slide Number Placeholder 0"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>`
		strSlideXml += '  <p:nvPr><p:ph type="sldNum" sz="quarter" idx="4294967295"/></p:nvPr>'
		strSlideXml += ' </p:nvSpPr>'
		strSlideXml += ' <p:spPr>'
		strSlideXml +=
			'<a:xfrm>' +
			`<a:off x="${getSmartParseNumber(slide._slideNumberProps.x, 'X', slide._presLayout)}" y="${getSmartParseNumber(slide._slideNumberProps.y, 'Y', slide._presLayout)}"/>` +
			`<a:ext cx="${slide._slideNumberProps.w ? getSmartParseNumber(slide._slideNumberProps.w, 'X', slide._presLayout) : '800000'}" cy="${slide._slideNumberProps.h ? getSmartParseNumber(slide._slideNumberProps.h, 'Y', slide._presLayout) : '300000'}"/>` +
			'</a:xfrm>' +
			' <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
			' <a:extLst><a:ext uri="{C572A759-6A51-4108-AA02-DFA0A04FC94B}"><ma14:wrappingTextBoxFlag val="0" xmlns:ma14="http://schemas.microsoft.com/office/mac/drawingml/2011/main"/></a:ext></a:extLst>' +
			'</p:spPr>'
		strSlideXml += '<p:txBody>'
		strSlideXml += '<a:bodyPr'
		if (slide._slideNumberProps.margin && Array.isArray(slide._slideNumberProps.margin)) {
			// Margins are inches (see `marginToEmu`), matching text-box and cell margins.
			strSlideXml += ` lIns="${marginToEmu(slide._slideNumberProps.margin[3] || 0)}"`
			strSlideXml += ` tIns="${marginToEmu(slide._slideNumberProps.margin[0] || 0)}"`
			strSlideXml += ` rIns="${marginToEmu(slide._slideNumberProps.margin[1] || 0)}"`
			strSlideXml += ` bIns="${marginToEmu(slide._slideNumberProps.margin[2] || 0)}"`
		} else if (typeof slide._slideNumberProps.margin === 'number') {
			strSlideXml += ` lIns="${marginToEmu(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` tIns="${marginToEmu(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` rIns="${marginToEmu(slide._slideNumberProps.margin || 0)}"`
			strSlideXml += ` bIns="${marginToEmu(slide._slideNumberProps.margin || 0)}"`
		}
		if (slide._slideNumberProps.valign) {
			strSlideXml += ` anchor="${slide._slideNumberProps.valign.replace('top', 't').replace('middle', 'ctr').replace('bottom', 'b')}"`
		}
		strSlideXml += '/>'
		strSlideXml += '  <a:lstStyle><a:lvl1pPr>'
		if (slide._slideNumberProps.fontFace || slide._slideNumberProps.fontSize || slide._slideNumberProps.color) {
			strSlideXml += `<a:defRPr sz="${clampFontSizeSz(slide._slideNumberProps.fontSize || 12)}">`
			if (slide._slideNumberProps.color) strSlideXml += genXmlColorSelection(slide._slideNumberProps.color)
			if (slide._slideNumberProps.fontFace) {
				// Caller-supplied via `slide.slideNumber({ fontFace })`; escaped so a `"`/`&` in the
				// name cannot close the attribute early and emit a non-parseable slide part.
				const slideNumFace = encodeXmlEntities(slide._slideNumberProps.fontFace)
				strSlideXml += `<a:latin typeface="${slideNumFace}"/><a:ea typeface="${slideNumFace}"/><a:cs typeface="${slideNumFace}"/>`
			}
			strSlideXml += '</a:defRPr>'
		}
		strSlideXml += '</a:lvl1pPr></a:lstStyle>'
		strSlideXml += '<a:p>'
		if (slide._slideNumberProps.align.startsWith('l')) strSlideXml += '<a:pPr algn="l"/>'
		else if (slide._slideNumberProps.align.startsWith('c')) strSlideXml += '<a:pPr algn="ctr"/>'
		else if (slide._slideNumberProps.align.startsWith('r')) strSlideXml += '<a:pPr algn="r"/>'
		else strSlideXml += '<a:pPr algn="l"/>'
		strSlideXml += `<a:fld id="${SLDNUMFLDID}" type="slidenum"><a:rPr b="${slide._slideNumberProps.bold ? 1 : 0}" lang="en-US"/>`
		strSlideXml += `<a:t>${slide._slideNum}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>`
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
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	let strXml = ''
	let arrTabRows: TableCell[][] = []
	let objTabOpts: ObjectOptions = {}
	let intColCnt = 0
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
	strXml = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}"/>`
	strXml +=
		`<p:cNvGraphicFramePr>${genXmlObjectLock('a:graphicFrameLocks', GRAPHIC_FRAME_LOCK_ATTRS, { noGrp: true, ...slideItemObj.options.objectLock }, slideItemObj.options.objectName)}</p:cNvGraphicFramePr>` +
		// A table bound to a layout placeholder emits that placeholder's <p:ph> (idx/type) so
		// PowerPoint treats the graphicFrame as filling the placeholder. The <p:ph>
		// precedes <p:extLst> per CT_ApplicationNonVisualDrawingProps document order.
		`  <p:nvPr>${genXmlPlaceholder(placeholderObj)}<p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1579011935"/></p:ext></p:extLst></p:nvPr>` +
		'</p:nvGraphicFramePr>'
	strXml += `<p:xfrm><a:off x="${x || (x === 0 ? 0 : EMU)}" y="${y || (y === 0 ? 0 : EMU)}"/><a:ext cx="${cx || (cx === 0 ? 0 : EMU)}" cy="${
		cy || EMU
	}"/></p:xfrm>`
	{
		const tblPrAttrs =
			(objTabOpts.rtl ? ' rtl="1"' : '') +
			(objTabOpts.hasHeader ? ' firstRow="1"' : '') +
			(objTabOpts.hasFooter ? ' lastRow="1"' : '') +
			(objTabOpts.hasBandedRows ? ' bandRow="1"' : '') +
			(objTabOpts.hasBandedColumns ? ' bandCol="1"' : '') +
			(objTabOpts.hasFirstColumn ? ' firstCol="1"' : '') +
			(objTabOpts.hasLastColumn ? ' lastCol="1"' : '')
		const tblPr = objTabOpts.tableStyle
			? `<a:tblPr${tblPrAttrs}><a:tableStyleId>${objTabOpts.tableStyle}</a:tableStyleId></a:tblPr>`
			: `<a:tblPr${tblPrAttrs}/>`
		strXml += `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>${tblPr}`
	}

	// STEP 2: Set column widths
	// Per-column inches from an explicit `colW` array, else split the table's
	// resolved EMU width (`cx`) evenly. `resolveTableColWidthsEmu` is the single
	// source of truth shared with the measured-fit pass. NOTE: divide the EMU
	// width, not the raw inches `options.w` — the latter collapsed auto-width
	// tables to ~0-EMU columns (e.g. `w=9` → `gridCol w="3"`).
	{
		const gridColsEmu = resolveTableColWidthsEmu(objTabOpts.colW, cx, intColCnt)
		strXml += '<a:tblGrid>'
		for (const w of gridColsEmu) strXml += `<a:gridCol w="${w}"/>`
		strXml += '</a:tblGrid>'
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

		// B: Start row
		strXml += `<a:tr h="${intRowH}">`

		// C: Loop over each CELL
		cells.forEach((cellObj) => {
			const cell: TableCell = cellObj

			const cellSpanAttrs = {
				rowSpan: cell.options?.rowspan && cell.options.rowspan > 1 ? cell.options.rowspan : undefined,
				gridSpan: cell.options?.colspan && cell.options.colspan > 1 ? cell.options.colspan : undefined,
				vMerge: cell._vmerge ? 1 : undefined,
				hMerge: cell._hmerge ? 1 : undefined,
			}
			let cellSpanAttrStr = Object.entries(cellSpanAttrs)
				.filter(([, v]) => !!v)
				.map(([k, v]) => `${String(k)}="${String(v)}"`)
				.join(' ')
			if (cellSpanAttrStr) cellSpanAttrStr = ' ' + cellSpanAttrStr

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
				strXml += `<a:tc${cellSpanAttrStr}><a:tcPr>${spanPrXml}</a:tcPr></a:tc>`
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
				? ` anchor="${cellOpts.valign.replace(/^c$/i, 'ctr').replace(/^m$/i, 'ctr').replace('center', 'ctr').replace('middle', 'ctr').replace('top', 't').replace('btm', 'b').replace('bottom', 'b')}"`
				: ''
			const cellTextDir =
				cellOpts.textDirection && cellOpts.textDirection !== 'horz' ? ` vert="${cellOpts.textDirection}"` : ''

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
			const cellMarginXml = ` marL="${marginToEmu(cellMargin[3])}" marR="${marginToEmu(cellMargin[1])}" marT="${marginToEmu(
				cellMargin[0]
			)}" marB="${marginToEmu(cellMargin[2])}"`

			// FUTURE: cell no-wrap support (add `horzOverflow="overflow"` to the cell's `<a:tcPr>`)

			// 4: Set CELL content and properties ==================================
			strXml += `<a:tc${cellSpanAttrStr}>${genXmlTextBody(cell)}<a:tcPr${cellMarginXml}${cellValign}${cellTextDir}>`

			// 5: Borders: Add any borders
			const cellBorder = Array.isArray(cellOpts.border) ? cellOpts.border : null
			if (cellBorder) strXml += genTableCellBorderXml(cellBorder)

			// 6: Close cell Properties & Cell
			strXml += cellFill
			strXml += '  </a:tcPr>'
			strXml += ' </a:tc>'
		})

		// D: Complete row
		strXml += '</a:tr>'
	})

	// STEP 5: Complete table
	strXml += '      </a:tbl>'
	strXml += '    </a:graphicData>'
	strXml += '  </a:graphic>'
	strXml += '</p:graphicFrame>'

	// STEP 6: Set table XML
	strSlideXml += strXml
	return strSlideXml
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
	locationAttr: string
): string {
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	// Lines can have zero cy, but text should not
	if (!slideItemObj.options.line && cy === 0) cy = EMU * 0.3

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
	strSlideXml += `<p:nvSpPr><p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}">`
	// <Hyperlink>
	if (slideItemObj.options.hyperlink?.url) {
		strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.options.hyperlink._rId}" tooltip="${slideItemObj.options.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.options.hyperlink.tooltip) : ''}"/>`
	}
	if (slideItemObj.options.hyperlink?.slide) {
		strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.options.hyperlink._rId}" tooltip="${slideItemObj.options.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.options.hyperlink.tooltip) : ''}" action="ppaction://hlinksldjump"/>`
	}
	// </Hyperlink>
	strSlideXml += '</p:cNvPr>'
	{
		const spLockXml = genXmlObjectLock(
			'a:spLocks',
			SHAPE_LOCK_ATTRS,
			slideItemObj.options.objectLock,
			slideItemObj.options.objectName
		)
		strSlideXml += '<p:cNvSpPr' + (slideItemObj.options?.isTextBox ? ' txBox="1"' : '')
		strSlideXml += spLockXml ? `>${spLockXml}</p:cNvSpPr>` : '/>'
	}
	// Prefer the resolved slide-layout placeholder; otherwise fall back to the shape's own
	// placeholder type so a standalone title/body text box still emits a real <p:ph>.
	strSlideXml += `<p:nvPr>${genXmlPlaceholder(slideItemObj._type === SlideObjectType.placeholder || (placeholderObj == null && slideItemObj.options?._placeholderType) ? slideItemObj : placeholderObj)}</p:nvPr>`
	strSlideXml += '</p:nvSpPr><p:spPr>'
	strSlideXml += `<a:xfrm${locationAttr}>`
	strSlideXml += `<a:off x="${x}" y="${y}"/>`
	strSlideXml += `<a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`

	if (slideItemObj.shape === 'custGeom') {
		strSlideXml += genXmlCustGeom(slideItemObj.options.points, cx, cy, slide._presLayout)
	} else {
		strSlideXml += genXmlPresetGeom(slideItemObj.shape ?? '', slideItemObj.options, cx, cy)
	}

	// Option: FILL
	strSlideXml += slideItemObj.options.fill ? genXmlColorSelection(slideItemObj.options.fill) : '<a:noFill/>'

	// shape Type: LINE: line color
	if (slideItemObj.options.line) {
		const lnAttrs =
			(slideItemObj.options.line.width ? ` w="${lineWidthToEmu(slideItemObj.options.line.width)}"` : '') +
			(slideItemObj.options.line.cap ? ` cap="${createLineCap(slideItemObj.options.line.cap)}"` : '')
		strSlideXml += `<a:ln${lnAttrs}>`
		strSlideXml += genXmlLineFill(slideItemObj.options.line)
		if (slideItemObj.options.line.dashType) strSlideXml += `<a:prstDash val="${slideItemObj.options.line.dashType}"/>`
		if (slideItemObj.options.line.beginArrowType)
			strSlideXml += `<a:headEnd type="${slideItemObj.options.line.beginArrowType}"/>`
		if (slideItemObj.options.line.endArrowType)
			strSlideXml += `<a:tailEnd type="${slideItemObj.options.line.endArrowType}"/>`
		// FUTURE: arrow-size support via the `w`/`len` attrs on headEnd/tailEnd
		// (e.g. `<a:headEnd type="arrow" w="lg" len="lg"/>`; each is 'sm'|'med'|'lg', a 3x3 grid)
		strSlideXml += '</a:ln>'
	}

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
	locationAttr: string,
	shapeIds: Map<SlideObject, number>
): string {
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	// A connector is emitted as <p:cxnSp> (a connector shape) rather than <p:sp>, so
	// PowerPoint treats it as a connector. Geometry/flip come from the shared resolution
	// above; the preset (straightConnector1 / bentConnector3 / curvedConnector3) is on `shape`.
	strSlideXml += '<p:cxnSp><p:nvCxnSpPr>'
	strSlideXml += `<p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}"/>`
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
	strSlideXml += `<a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
	{
		// Bent/curved connectors carry adjustable jogs as `<a:gd name="adjN" fmla="val …"/>`
		// (1000ths-of-a-percent). With none, the empty `<a:avLst/>` leaves the preset default (50%).
		const adj = slideItemObj.options._connectorAdj || []
		const avLst = adj.map((val, i) => `<a:gd name="adj${i + 1}" fmla="val ${val}"/>`).join('')
		strSlideXml += `<a:prstGeom prst="${slideItemObj.shape}"><a:avLst>${avLst}</a:avLst></a:prstGeom>`
	}
	{
		const ln = slideItemObj.options.line || {}
		const lnAttrs =
			(ln.width ? ` w="${lineWidthToEmu(ln.width)}"` : '') + (ln.cap ? ` cap="${createLineCap(ln.cap)}"` : '')
		strSlideXml += `<a:ln${lnAttrs}>`
		strSlideXml += genXmlLineFill(ln)
		if (ln.dashType) strSlideXml += `<a:prstDash val="${ln.dashType}"/>`
		if (ln.beginArrowType) strSlideXml += `<a:headEnd type="${ln.beginArrowType}"/>`
		if (ln.endArrowType) strSlideXml += `<a:tailEnd type="${ln.endArrowType}"/>`
		strSlideXml += '</a:ln>'
	}
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
	locationAttr: string,
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
	strSlideXml += '<p:pic>'
	strSlideXml += '  <p:nvPicPr>'
	strSlideXml += `<p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(
		slideItemObj.options.altText || slideItemObj.image || ''
	)}">`
	if (slideItemObj.hyperlink?.url) {
		strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.hyperlink._rId}" tooltip="${
			slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : ''
		}"/>`
	}
	if (slideItemObj.hyperlink?.slide) {
		strSlideXml += `<a:hlinkClick r:id="rId${slideItemObj.hyperlink._rId}" tooltip="${
			slideItemObj.hyperlink.tooltip ? encodeXmlEntities(slideItemObj.hyperlink.tooltip) : ''
		}" action="ppaction://hlinksldjump"/>`
	}
	strSlideXml += '    </p:cNvPr>'
	// Default to locking aspect ratio (PowerPoint's own behavior); user `objectLock` overrides any flag, incl. noChangeAspect.
	strSlideXml += `    <p:cNvPicPr>${genXmlObjectLock('a:picLocks', PICTURE_LOCK_ATTRS, { noChangeAspect: true, ...slideItemObj.options.objectLock }, slideItemObj.options.objectName)}</p:cNvPicPr>`
	strSlideXml += '    <p:nvPr>' + genXmlPlaceholder(placeholderObj) + '</p:nvPr>'
	strSlideXml += '  </p:nvPicPr>'
	// Duotone recolor: maps shadows→shadow color, highlights→highlight color.
	// `<a:duotone>` is one of the `<a:blip>` image-effect children (CT_Blip);
	// it sits alongside `alphaModFix` and before any `extLst`.
	strSlideXml += '<p:blipFill>'
	// NOTE: This works for both cases: either `path` or `data` contains the SVG
	if ((slide._relsMedia || []).find((rel) => rel.rId === slideItemObj.imageRid)?.extn === 'svg') {
		strSlideXml += `<a:blip r:embed="rId${(slideItemObj.imageRid ?? 0) - 1}">`
		strSlideXml += slideItemObj.options.transparency
			? ` <a:alphaModFix amt="${Math.round((100 - slideItemObj.options.transparency) * FIXED_PCT_PER_PERCENT)}"/>`
			: ''
		strSlideXml += slideItemObj.options.duotone
			? `<a:duotone>${createColorElement(slideItemObj.options.duotone.shadow)}${createColorElement(slideItemObj.options.duotone.highlight)}</a:duotone>`
			: ''
		strSlideXml += ' <a:extLst>'
		strSlideXml += '  <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
		strSlideXml += `   <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId${slideItemObj.imageRid}"/>`
		strSlideXml += '  </a:ext>'
		strSlideXml += ' </a:extLst>'
		strSlideXml += '</a:blip>'
	} else {
		strSlideXml += `<a:blip r:embed="rId${slideItemObj.imageRid}">`
		strSlideXml += slideItemObj.options.transparency
			? `<a:alphaModFix amt="${Math.round((100 - slideItemObj.options.transparency) * FIXED_PCT_PER_PERCENT)}"/>`
			: ''
		strSlideXml += slideItemObj.options.duotone
			? `<a:duotone>${createColorElement(slideItemObj.options.duotone.shadow)}${createColorElement(slideItemObj.options.duotone.highlight)}</a:duotone>`
			: ''
		strSlideXml += '</a:blip>'
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
		strSlideXml += '  <a:stretch><a:fillRect/></a:stretch>'
	}
	strSlideXml += '</p:blipFill>'
	strSlideXml += '<p:spPr>'
	strSlideXml += ' <a:xfrm' + locationAttr + '>'
	strSlideXml += `  <a:off x="${x}" y="${y}"/>`
	strSlideXml += `  <a:ext cx="${imgWidth}" cy="${imgHeight}"/>`
	strSlideXml += ' </a:xfrm>'
	// Clip the picture to a geometry. `points` (freeform custGeom) takes precedence over `shape`/`rounding`;
	// otherwise `shape` wins over `rounding` (shorthand for an ellipse), falling back to a plain rectangle.
	if (slideItemObj.options.points) {
		strSlideXml += ' ' + genXmlCustGeom(slideItemObj.options.points, imgWidth, imgHeight, slide._presLayout)
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
	if (slideItemObj.options.line) {
		const imgLine = slideItemObj.options.line
		const lnAttrs =
			(imgLine.width ? ` w="${lineWidthToEmu(imgLine.width)}"` : '') +
			(imgLine.cap ? ` cap="${createLineCap(imgLine.cap)}"` : '')
		strSlideXml += `<a:ln${lnAttrs}>`
		strSlideXml += genXmlLineFill(imgLine)
		if (imgLine.dashType) strSlideXml += `<a:prstDash val="${imgLine.dashType}"/>`
		if (imgLine.beginArrowType) strSlideXml += `<a:headEnd type="${imgLine.beginArrowType}"/>`
		if (imgLine.endArrowType) strSlideXml += `<a:tailEnd type="${imgLine.endArrowType}"/>`
		strSlideXml += '</a:ln>'
	}

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
	locationAttr: string
): string {
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	if (slideItemObj.mtype === 'online') {
		strSlideXml += '<p:pic>'
		strSlideXml += ' <p:nvPicPr>'
		// cNvPr/@id must be unique across every shape on the slide, so it uses the slide-object
		// index (idx + 2) like all other shapes — NOT mediaRid, which lives in the relationship-id
		// space and collides with a sibling shape's idx (duplicate ids => PowerPoint reports the
		// file corrupt, 0x80070570). The preview image is still bound via <a:blip r:embed> below.
		strSlideXml += `<p:cNvPr id="${idx + 2}" name="${
			slideItemObj.options.objectName
		}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>`
		strSlideXml += ` <p:cNvPicPr>${genXmlObjectLock('a:picLocks', PICTURE_LOCK_ATTRS, { noChangeAspect: true, ...slideItemObj.options.objectLock }, slideItemObj.options.objectName)}</p:cNvPicPr>`
		strSlideXml += ' <p:nvPr>'
		// External-link video: <a:videoFile r:link> at the ECMA rel, <p14:media r:link>
		// at the MS-2007 media rel (both External, sharing the link Target). Mirrors the
		// embedded branch but uses r:link (no media binary part).
		strSlideXml += `  <a:videoFile r:link="rId${slideItemObj.mediaRid}"/>`
		strSlideXml += '  <p:extLst>'
		strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
		strSlideXml += `    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:link="rId${(slideItemObj.mediaRid ?? 0) + 1}"/>`
		strSlideXml += '   </p:ext>'
		strSlideXml += '  </p:extLst>'
		strSlideXml += ' </p:nvPr>'
		strSlideXml += ' </p:nvPicPr>'
		strSlideXml += ` <p:blipFill><a:blip r:embed="rId${(slideItemObj.mediaRid ?? 0) + 2}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` // NOTE: Preview image is required!
		strSlideXml += ' <p:spPr>'
		strSlideXml += `  <a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
		strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
		strSlideXml += ' </p:spPr>'
		strSlideXml += '</p:pic>'
	} else {
		strSlideXml += '<p:pic>'
		strSlideXml += ' <p:nvPicPr>'
		// cNvPr/@id must be unique across every shape on the slide, so it uses the slide-object
		// index (idx + 2) like all other shapes — NOT mediaRid, which lives in the relationship-id
		// space and collides with a sibling shape's idx (duplicate ids => PowerPoint reports the
		// file corrupt, 0x80070570). The preview image is still bound via <a:blip r:embed> below.
		strSlideXml += `<p:cNvPr id="${idx + 2}" name="${
			slideItemObj.options.objectName
		}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>`
		strSlideXml += ` <p:cNvPicPr>${genXmlObjectLock('a:picLocks', PICTURE_LOCK_ATTRS, { noChangeAspect: true, ...slideItemObj.options.objectLock }, slideItemObj.options.objectName)}</p:cNvPicPr>`
		strSlideXml += ' <p:nvPr>'
		// EG_Media choice: audio embeds use <a:audioFile>, video uses <a:videoFile>
		strSlideXml += `  <a:${slideItemObj.mtype === 'audio' ? 'audioFile' : 'videoFile'} r:link="rId${slideItemObj.mediaRid}"/>`
		strSlideXml += '  <p:extLst>'
		strSlideXml += '   <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">'
		strSlideXml += `    <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="rId${(slideItemObj.mediaRid ?? 0) + 1}"/>`
		strSlideXml += '   </p:ext>'
		strSlideXml += '  </p:extLst>'
		strSlideXml += ' </p:nvPr>'
		strSlideXml += ' </p:nvPicPr>'
		strSlideXml += ` <p:blipFill><a:blip r:embed="rId${(slideItemObj.mediaRid ?? 0) + 2}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` // NOTE: Preview image is required!
		strSlideXml += ' <p:spPr>'
		strSlideXml += `  <a:xfrm${locationAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
		strSlideXml += '  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
		strSlideXml += ' </p:spPr>'
		strSlideXml += '</p:pic>'
	}
	return strSlideXml
}

/**
 * Render a `chart` slide object to its `<p:graphicFrame>` XML referencing the chart part.
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
	let strSlideXml = ''
	// Caller guarantees options is set (see slideObjectToXml); re-narrow for this scope.
	slideItemObj.options = slideItemObj.options || {}
	strSlideXml += '<p:graphicFrame>'
	strSlideXml += ' <p:nvGraphicFramePr>'
	strSlideXml += `   <p:cNvPr id="${idx + 2}" name="${slideItemObj.options.objectName}" descr="${encodeXmlEntities(slideItemObj.options.altText || '')}"/>`
	strSlideXml += '   <p:cNvGraphicFramePr/>'
	strSlideXml += `   <p:nvPr>${genXmlPlaceholder(placeholderObj)}</p:nvPr>`
	strSlideXml += ' </p:nvGraphicFramePr>'
	strSlideXml += ` <p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm>`
	strSlideXml += ' <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
	strSlideXml += '  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
	strSlideXml += `   <c:chart r:id="rId${slideItemObj.chartRid}" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`
	strSlideXml += '  </a:graphicData>'
	strSlideXml += ' </a:graphic>'
	strSlideXml += '</p:graphicFrame>'
	return strSlideXml
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
		rels.push(voidEl('Relationship', { Id: `rId${rel.rId}`, Type: OFFICE_REL + 'chart', Target: rel.Target }))
	})
	;(slide._relsMedia || []).forEach((rel: SlideRelMedia) => {
		const relType = rel.type.toLowerCase()
		// `voidEl` escapes the Target on the way out; the probe has to compare against
		// those emitted bytes, so it needs the escaped form computed separately here.
		const relTarget = encodeXmlEntities(rel.Target)
		const media = (type: string, targetMode?: string): string =>
			voidEl('Relationship', { Id: `rId${rel.rId}`, Type: type, Target: rel.Target, TargetMode: targetMode })
		lastRid = Math.max(lastRid, rel.rId)
		if (relType.includes('image')) {
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
