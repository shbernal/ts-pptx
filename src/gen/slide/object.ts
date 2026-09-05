/**
 * ts-pptx: slide object serialization
 *
 * The per-shape `<p:spTree>` builder: `slideObjectToXml` walks a slide/layout's objects
 * (recursing into groups, allocating `<p:cNvPr>` ids as it goes) and dispatches each one to the
 * matching renderer under `objects/`. The group render and the slide-number placeholder stay
 * here because both consume the slide-wide child-id counter the walk maintains.
 * `slideObjectRelationsToXml` emits the matching `.rels` targets.
 */

import { SlideObjectType } from '../../enums.js'
import { STRETCH_FILL_RECT } from '../drawingml/src-rect.js'
import { prstGeomRect } from '../drawingml/geometry.js'
import { DEF_PRES_LAYOUT_NAME, SLDNUM_PLACEHOLDER_TEXT, SLDNUMFLDID } from '../../constants-internal.js'
import type { ObjectOptions, SlideNumberProps } from '../../types/index.js'
import type {
	PresSlideInternal,
	SlideLayoutInternal,
	SlideObject,
	SlideRel,
	SlideRelChart,
	SlideRelMedia,
} from '../../types/internal.js'
import { encodeXmlAttrValue, getDuplicateObjectNames, isHyperlinkRel } from '../utils.js'
import { fillNamesPaint, genXmlColorSelection } from '../drawingml/fill.js'
import { rejectEmptyColor } from '../drawingml/color.js'
import { convertRotationDegrees, getSmartParseNumber, resolveInsetsEmu } from '../../units-internal.js'
import { warn } from '../../diagnostics.js'
import { clampFontSizeSz } from '../drawingml/clamp.js'
import { resolveTextAnchor } from '../drawingml/text-body.js'
import { genXmlObjectLock, GROUP_SHAPE_LOCK_ATTRS } from '../drawingml/locks.js'
import { el, raw, voidEl, type XmlAttrs } from '../oxml/el.js'
import { cNvPrOpen, grpXfrmEl, type RenderContext } from './objects/shared.js'
import { renderChartObject } from './objects/chart.js'
import { renderConnectorObject } from './objects/connector.js'
import { renderImageObject } from './objects/image.js'
import { renderMediaObject } from './objects/media.js'
import { renderModel3dObject } from './objects/model3d.js'
import { renderOleObject } from './objects/ole.js'
import { renderTableObject } from './objects/table.js'
import { renderTextObject } from './objects/text.js'
import { renderZoomObject } from './objects/zoom.js'
import { collectSlideShapeIds } from './shape-ids.js'
import {
	AUDIO_REL,
	CHART_REL,
	CHARTEX_REL,
	IMAGE_REL,
	MS_MEDIA_REL,
	SLIDE_REL,
	VIDEO_REL,
} from '../../ooxml/rel-types.js'
import { externalHyperlinkRel, relationshipEl, relationshipsPart } from '../opc/rels.js'
import { xsdBool, xsdBoolIfTrue } from '../../ooxml/xsd-boolean.js'

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
 * Normalize one axis of a placement box so its extent is never negative.
 *
 * `<a:ext cx>`/`<a:ext cy>` are `ST_PositiveCoordinate` (ECMA-376 Part 1, `CT_PositiveSize2D`), so a
 * negative extent is schema-invalid. PowerPoint rejects the *whole* presentation with "The file or
 * directory is corrupted and unreadable" (0x80070570) and names no shape, part, or slide, while
 * LibreOffice renders the same package happily — so the defect is invisible until the deck reaches
 * PowerPoint.
 *
 * A signed delta is the natural way to write "draw from A to B" (`{ x: x0, y: y0, w: x1 - x0, h: y1 - y0 }`),
 * and it goes negative the moment the line runs leftward or upward. The correct encoding is the
 * normalized bounding box plus a flip, which is what `addConnector` already derives from its own
 * endpoints (`src/gen/define/connector.ts`); this gives every other shape kind the same treatment at
 * the one point where every `Coord` form ('50%', '2in', a number) has been resolved to EMU.
 *
 * The caller composes `flip` onto any `flipH`/`flipV` the author set, so an explicit flip plus a
 * negative extent cancel out rather than double-applying.
 * @param off - the axis origin in EMU (`a:off` `x` or `y`)
 * @param ext - the axis extent in EMU (`a:ext` `cx` or `cy`), possibly negative
 * @returns the min-corner origin, the absolute extent, and whether the axis was mirrored
 */
const normalizeAxisExtent = (off: number, ext: number): { off: number; ext: number; flip: boolean } =>
	ext < 0 ? { off: off + ext, ext: -ext, flip: true } : { off, ext, flip: false }

/**
 * Whether this part has a real slide number to cache in a `slidenum` field.
 *
 * Only a slide does. A master leaves `_slideNum` null and a layout carries the internal 1000+
 * counter that keeps layout media keys from colliding with slide media — neither is a page
 * number, and neither belongs in the field's cached text.
 */
function hasRealSlideNumber(slide: PresSlideInternal | SlideLayoutInternal): boolean {
	return slide._slideNum != null && slide._slideNum < 1000
}

/**
 * Warn about duplicate Selection Pane identities on this slide.
 *
 * Unique `objectName` values are what consumers (e.g. semantic manifests) rely on, so
 * collisions are flagged loudly. Deliberately NOT `warnOnce`: this runs per emit, and two
 * `write()` calls are two separate exports — a caller capturing diagnostics around the second
 * would otherwise be told its deck is clean because the first one had already said otherwise. Groups are recursed into: a group's children are
 * `<p:cNvPr>`-named on this same slide, so a child colliding with a top-level object (or with a
 * child of another group) is a collision the Selection Pane shows, and checking only the top
 * level cannot see it.
 */
function warnDuplicateObjectNames(slide: PresSlideInternal | SlideLayoutInternal): void {
	const collectObjectNames = (objects: SlideObject[]): string[] =>
		objects.flatMap((obj) => [
			...(typeof obj.options?.objectName === 'string' ? [obj.options.objectName] : []),
			...collectObjectNames(obj._groupObjects || []),
		])
	const duplicateObjectNames = getDuplicateObjectNames(collectObjectNames(slide._slideObjects))
	if (duplicateObjectNames.length > 0) {
		warn(
			'object-name/duplicate',
			`duplicate objectName value(s) emitted on a single slide: ${duplicateObjectNames.join(', ')}. Selection Pane identities should be unique.`
		)
	}
}

/**
 * The slide's `<p:bg>`, or `''` when it has none.
 *
 * Exactly one of the three forms is emitted: an image background (a rel the media pass resolved),
 * a colour or gradient, or the white default that only the built-in layout gets. They are
 * mutually exclusive because a master carrying both a `color` and a `path` must still produce a
 * single `<p:bg>`.
 */
function slideBackgroundXml(slide: PresSlideInternal | SlideLayoutInternal): string {
	let strSlideXml = ''
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
							raw(STRETCH_FILL_RECT),
						])
					),
					raw(voidEl('a:effectLst')),
				])
			)
		)
		// A background paints whenever it names one — `fillNamesPaint`, the same test the stroke
		// side uses. This gate used to accept a `color` or the literal `type: 'gradient'`, so a
		// background authored as `{ gradient }`, `{ pattern }` or `{ type: 'pattern', … }` was
		// dropped without a word and the slide came out inheriting the master.
	} else if (fillNamesPaint(slide.background)) {
		strSlideXml += el(
			'p:bg',
			null,
			raw(el('p:bgPr', null, [raw(genXmlColorSelection(slide.background)), raw(voidEl('a:effectLst'))]))
		)
	} else if (!slide.background && slide._name && slide._name === DEF_PRES_LAYOUT_NAME) {
		// NOTE: Default [white] background is needed on slideMaster1.xml to avoid gray background in Keynote (and Finder previews)
		strSlideXml += el('p:bg', null, raw(el('p:bgRef', { idx: '1001' }, raw(voidEl('a:schemeClr', { val: 'bg1' })))))
	}
	return strSlideXml
}

/**
 * The fixed opening of a slide's `<p:spTree>`.
 *
 * OOXML requires the shape tree to open with the implicit top-level group's non-visual props
 * (`<p:nvGrpSpPr>`, the reserved `cNvPr id="1"`) and an identity group transform (off/ext and
 * chOff/chExt all zero) before any child shape. This is the slide's built-in root group, not a
 * user-authored `addGroup` — hence the zeroed frame, and hence taking no arguments: every slide,
 * layout and master opens its tree with exactly these bytes.
 */
function spTreeOpenXml(): string {
	let strSlideXml = ''
	strSlideXml += '<p:spTree>'
	strSlideXml += el('p:nvGrpSpPr', null, [
		raw(voidEl('p:cNvPr', { id: '1', name: '' })),
		raw(voidEl('p:cNvGrpSpPr')),
		raw(voidEl('p:nvPr')),
	])
	strSlideXml += el('p:grpSpPr', null, raw(grpXfrmEl({ x: '0', y: '0', cx: '0', cy: '0' })))
	return strSlideXml
}

/**
 * The slide-number placeholder shape.
 *
 * Emitted after every authored object, so its `<p:cNvPr>` id can come from the same monotonic
 * counter the walk has been advancing — the caller allocates it and passes it in.
 *
 * @param slide - the slide, layout or master being emitted
 * @param snProps - the part's slide-number properties, read but never written: a serializer
 *   does not normalize the authored model (see the contract on `RenderContext.itemOpts`)
 * @param slideNumberId - the `<p:cNvPr>` id to use, already allocated by the caller
 */
function slideNumberPlaceholderXml(
	slide: PresSlideInternal | SlideLayoutInternal,
	snProps: SlideNumberProps,
	slideNumberId: number
): string {
	let strSlideXml = ''
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
		prstGeomRect({ openPrefix: ' ' }) +
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
	// NOTE: attribute ORDER is byte-significant; `resolveInsetsEmu` owns the array's own
	// [Top, Right, Bottom, Left] order.
	const snInsets = resolveInsetsEmu(snProps.margin)
	strSlideXml += voidEl('a:bodyPr', {
		lIns: snInsets?.l ?? null,
		tIns: snInsets?.t ?? null,
		rIns: snInsets?.r ?? null,
		bIns: snInsets?.b ?? null,
		anchor: resolveTextAnchor(snProps.valign),
	})
	let defRPr = ''
	rejectEmptyColor(snProps.color, 'slideNumber color')
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

	// Anything not starting c/r falls back to 'l'. This used to be preceded by a write of
	// `'left'` back onto `snProps`, justified as "other readers rely on it" — the only other
	// reader is this line, and it already defaults.
	const snAlignRaw = snProps.align ?? 'left'
	const snAlign = snAlignRaw.startsWith('c') ? 'ctr' : snAlignRaw.startsWith('r') ? 'r' : 'l'
	strSlideXml += el('a:p', null, [
		raw(voidEl('a:pPr', { algn: snAlign })),
		raw(
			el('a:fld', { id: SLDNUMFLDID, type: 'slidenum' }, [
				// NOTE: `b` is emitted as "0" when unset, unlike the run properties elsewhere which omit it.
				raw(voidEl('a:rPr', { b: xsdBool(snProps.bold), lang: 'en-US' })),
				// `<a:t>` inside an `a:fld` is the *cached* rendering of the field, so it is only a
				// slide number where there is a slide number to cache. A master has none
				// (`_slideNum` is null) and a layout carries the internal 1000+ counter, so this
				// used to ship `<a:t>null</a:t>` in every master and `<a:t>1004</a:t>` in a layout —
				// invisible in PowerPoint, which recomputes the field on open, but read straight
				// out by anything that takes the cache at face value (a text extractor, a search
				// indexer, this library's own read path). What PowerPoint itself caches on a master
				// or layout is the placeholder glyph, so emit that.
				//
				// The child must stay non-null: `el()` drops a null child entirely, and an `a:fld`
				// with no `a:t` is a different construct from one with placeholder text.
				raw(el('a:t', null, hasRealSlideNumber(slide) ? String(slide._slideNum) : SLDNUM_PLACEHOLDER_TEXT)),
			])
		),
		raw(voidEl('a:endParaRPr', { lang: 'en-US' })),
	])
	strSlideXml += '</p:txBody></p:sp>'
	return strSlideXml
}

/**
 * Transforms a slide or slideLayout to resulting XML string - Creates `ppt/slide*.xml`
 * @param {PresSlideInternal|SlideLayoutInternal} slideObject - slide object created within createSlideObject
 * @return {string} XML string with <p:cSld> as the root
 */
export function slideObjectToXml(slide: PresSlideInternal | SlideLayoutInternal): string {
	// `_name` is escaped HERE, at emission, unlike `objectName`'s single-escape-upstream design
	// (see `cNvPrOpen`): `_name` doubles as the raw lookup key `addSlide({masterTitle})` matches
	// against the caller's `title` string (presentation.ts, `layout._name === masterTitle`), so it
	// must stay unescaped until the last possible moment or that match breaks for any title
	// containing `&`/`<`/`"`. Plain slides' default `_name` ("Slide N", slide.ts) never contains
	// XML metacharacters, so escaping it here is a no-op for that path.
	// The element stays a template because it wraps the entire slide, built by append below.
	let strSlideXml: string = slide._name ? '<p:cSld name="' + encodeXmlAttrValue(slide._name) + '">' : '<p:cSld>'

	warnDuplicateObjectNames(slide)

	// STEP 1: Add the background, if this part defines one
	strSlideXml += slideBackgroundXml(slide)

	// STEP 2: Continue slide by starting spTree node
	strSlideXml += spTreeOpenXml()

	// Every object's <p:cNvPr> id, allocated once, up front. It has to be up front for the
	// references that cannot wait for the walk below to reach their target (connector shape
	// bindings, animation spids), and it is the ONLY allocation: the walk reads this map rather
	// than recomputing the same arithmetic, so there is nothing left to keep in step.
	const shapeIds = collectSlideShapeIds(slide._slideObjects)

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
		// Normalized the same way the render path below normalizes its own box, so a child with a
		// negative extent contributes the box it will actually be emitted with — an un-normalized
		// `x + cx` would put the group's `maxX` *left* of its `minX` and size the group wrong.
		const bx = normalizeAxisExtent(
			typeof o.x !== 'undefined' ? getSmartParseNumber(o.x, 'X', slide._presLayout) : 0,
			typeof o.w !== 'undefined' ? getSmartParseNumber(o.w, 'X', slide._presLayout) : 0
		)
		const by = normalizeAxisExtent(
			typeof o.y !== 'undefined' ? getSmartParseNumber(o.y, 'Y', slide._presLayout) : 0,
			typeof o.h !== 'undefined' ? getSmartParseNumber(o.h, 'Y', slide._presLayout) : 0
		)
		return { x: bx.off, y: by.off, cx: bx.ext, cy: by.ext }
	}

	// Render one slide object — and, for a group, its children recursively — to an XML fragment.
	// Closes over `slide` and `childIdxAlloc`. Uses a local
	// `strSlideXml` accumulator (shadowing the slide-level one) so the existing per-object
	// `strSlideXml +=` appends compose into the returned fragment unchanged.
	const renderSlideObjectXml = (slideItemObj: SlideObject): string => {
		// The one allocator's answer. Every renderer used to recompute `idx + 2` and the group walk
		// kept a second counter; a `p:cNvPr` id is now read from the same map the forward
		// references resolve through, so the two cannot drift.
		const shapeId = shapeIds.get(slideItemObj) ?? 2
		let strSlideXml = ''
		let x = 0
		let y = 0
		// Annotated `number` rather than inferring the branded `Emu` this initializer returns: `x`, `y`,
		// and `cy` are all plain numbers, every `render*Object` takes `number`, and normalization below
		// reassigns all four from one helper.
		let cx: number = getSmartParseNumber('75%', 'X', slide._presLayout)
		let cy = 0
		let placeholderObj: SlideObject | null = null

		const slideLayout = (slide as PresSlideInternal)._slideLayout
		const wantedPlaceholder = slideItemObj.options?.placeholder
		if (slideLayout?._slideObjects !== undefined && wantedPlaceholder) {
			placeholderObj =
				slideLayout._slideObjects.filter(
					(object: SlideObject) => object.options?.placeholder === wantedPlaceholder
				)[0] ?? null
		}

		// A: Set option vars. Resolved to a LOCAL, never assigned back: a serializer does not
		// normalize the authored model, which is the contract `RenderContext.itemOpts` and
		// `slideNumberPlaceholderXml` both state, and these were the last two writes against it.
		const itemOpts = slideItemObj.options ?? {}

		// Each axis, most specific source first: what the caller stated on this object, else what
		// the layout placeholder it names states, else the default already in the variable.
		//
		// The placeholder used to be applied AFTER this, and unconditionally — so
		// `addText('own coords', { placeholder: 'body', x: 5, y: 3, w: 2, h: 1 })` had all four of
		// its stated values thrown away with no diagnostic, while the same object with a
		// *partial* frame and no placeholder warns loudly. An explicit option beats an inherited
		// one everywhere else in this library; a placeholder is an inherited one.
		//
		// It also has to happen before the normalization below rather than after it: the
		// placeholder's extents used to skip normalization entirely while the flip flags were
		// derived from the object's own signs, so a negative extent on either side composed wrong.
		const phOpts = placeholderObj?.options ?? {}
		const inherited = <T>(own: T | undefined, ph: T | undefined): T | undefined =>
			own !== undefined ? own : (ph ?? undefined)
		const ownX = inherited(itemOpts.x, phOpts.x)
		const ownY = inherited(itemOpts.y, phOpts.y)
		const ownW = inherited(itemOpts.w, phOpts.w)
		const ownH = inherited(itemOpts.h, phOpts.h)
		if (ownX !== undefined) x = getSmartParseNumber(ownX, 'X', slide._presLayout)
		if (ownY !== undefined) y = getSmartParseNumber(ownY, 'Y', slide._presLayout)
		if (ownW !== undefined) cx = getSmartParseNumber(ownW, 'X', slide._presLayout)
		if (ownH !== undefined) cy = getSmartParseNumber(ownH, 'Y', slide._presLayout)

		// A negative `w`/`h` becomes a min-corner origin, an absolute extent, and a flip — never a
		// negative `<a:ext>`, which is out of range for `ST_PositiveCoordinate` and costs the whole
		// package (see `normalizeAxisExtent`). Done here, after every `Coord` form has resolved to EMU,
		// so `'-2in'` and `'-25%'` normalize alongside a plain negative number.
		const normX = normalizeAxisExtent(x, cx)
		const normY = normalizeAxisExtent(y, cy)
		x = normX.off
		cx = normX.ext
		y = normY.off
		cy = normY.ext

		// Set w/h now that smart parse is done
		const imgWidth = cx
		const imgHeight = cy

		// The `<a:xfrm>` placement attributes, shared by every shape kind that has a transform.
		// NOTE: order is byte-significant (flipH, flipV, rot), and `null` means omitted — `rotate: 0`
		// stays absent, matching the truthiness test this replaced.
		// A flip derived from a negative extent XORs with the author's own: `{ w: -2, flipH: true }`
		// is a box mirrored twice, i.e. not mirrored at all.
		const locationAttrs: XmlAttrs = {
			flipH: xsdBoolIfTrue(Boolean(itemOpts.flipH) !== normX.flip),
			flipV: xsdBoolIfTrue(Boolean(itemOpts.flipV) !== normY.flip),
			rot: itemOpts.rotate ? convertRotationDegrees(itemOpts.rotate) : null,
		}

		// B: Add OBJECT to the current Slide.
		// Each renderer below is called from here and nowhere else, so everything resolved above
		// travels as one `RenderContext` — including `itemOpts`, already normalized in step A. Six
		// of them used to take the whole `slideItemObj` and re-run `options = options || {}` as
		// their first statement, on an argument their only caller had already fixed; that
		// assignment existed to re-narrow the type, and threading the narrowed value in does the
		// same job without a mutation in the middle of an emit pass. If one of them ever gains a
		// second call site, normalizing is that caller's job — a contract stated once at the
		// boundary beats a defensive copy of it in every callee.
		const ctx: RenderContext = {
			obj: slideItemObj,
			shapeId,
			slide,
			frame: { x, y, cx, cy },
			placeholder: placeholderObj,
			locationAttrs,
			itemOpts,
		}

		switch (slideItemObj._type) {
			case SlideObjectType.table:
				strSlideXml += renderTableObject(ctx)
				break
			case SlideObjectType.text:
			case SlideObjectType.placeholder:
				strSlideXml += renderTextObject(ctx)
				break
			case SlideObjectType.connector:
				strSlideXml += renderConnectorObject(ctx, shapeIds)
				break
			case SlideObjectType.image:
				strSlideXml += renderImageObject(ctx, { imgWidth, imgHeight })
				break
			case SlideObjectType.media:
				strSlideXml += renderMediaObject(ctx)
				break
			case SlideObjectType.chart:
				strSlideXml += renderChartObject(ctx)
				break
			case SlideObjectType.oleObject:
				strSlideXml += renderOleObject(ctx)
				break
			case SlideObjectType.zoom:
				strSlideXml += renderZoomObject(ctx)
				break
			case SlideObjectType.model3d:
				strSlideXml += renderModel3dObject(ctx)
				break

			case SlideObjectType.group: {
				const groupChildren = slideItemObj._groupObjects || []

				// Render children (recursively for nested groups). Each child gets a unique id via
				// `childIdxAlloc` (children are not in `_slideObjects`); the shared counter keeps ids
				// collision-free across nesting depth.
				let innerXml = ''
				// No `child.options = child.options || {}` here: `renderSlideObjectXml` resolves its
				// own options on entry, so the parent has nothing to prepare for the child.
				groupChildren.forEach((child) => {
					innerXml += renderSlideObjectXml(child)
				})

				// Identity child coordinate space (chOff/chExt == off/ext) at every depth, so children
				// keep their slide-absolute coordinates. Use explicit x/y/w/h when all four are given,
				// else the bounding box of the children (recursing into nested auto-sized groups).
				// A partial frame warns and falls back whole rather than letting the unset axes take the
				// per-object defaults above (`cy` = 0 among them) and emit a degenerate group.
				const givenAxes = givenGroupFrameAxes(itemOpts)
				if (givenAxes.length > 0 && givenAxes.length < GROUP_FRAME_AXES.length) {
					const missingAxes = GROUP_FRAME_AXES.filter((axis) => !givenAxes.includes(axis))
					warn(
						'group/partial-frame',
						`addGroup: group "${itemOpts.objectName ?? ''}" has a partial frame (${givenAxes.join('/')} given, ${missingAxes.join('/')} missing); using auto-bounds (the bounding box of its children) instead. Pass all of x/y/w/h, or none.`
					)
				}
				const gb = hasCompleteGroupFrame(itemOpts) ? { x, y, cx, cy } : resolveObjBounds(slideItemObj)
				const gx: number = gb.x
				const gy: number = gb.y
				const gcx: number = gb.cx
				const gcy: number = gb.cy

				const grpLockXml = genXmlObjectLock(
					'a:grpSpLocks',
					GROUP_SHAPE_LOCK_ATTRS,
					itemOpts.objectLock,
					itemOpts.objectName
				)
				strSlideXml += '<p:grpSp>'
				strSlideXml += el('p:nvGrpSpPr', null, [
					raw(cNvPrOpen(shapeId, itemOpts.objectName, itemOpts.altText || '') + '/>'),
					// Paired only when there are locks to carry; otherwise self-closing.
					raw(grpLockXml ? el('p:cNvGrpSpPr', null, raw(grpLockXml)) : voidEl('p:cNvGrpSpPr')),
					raw(voidEl('p:nvPr')),
				])
				strSlideXml += el('p:grpSpPr', null, raw(grpXfrmEl({ x: gx, y: gy, cx: gcx, cy: gcy }, locationAttrs)))
				strSlideXml += innerXml
				strSlideXml += '</p:grpSp>'
				break
			}

			// The four `SlideObjectType` members that are not slide shapes, spelled out so adding a
			// member to the enum fails to compile here rather than silently emitting nothing.
			case SlideObjectType.notes:
				// Speaker notes live in `_slideObjects` but belong to the notes part, which
				// `gen/slide/notes.ts` builds from the same entry. Emitting nothing here is
				// load-bearing, not a gap.
				break
			case SlideObjectType.tablecell:
				// Rendered by its owning table, through `renderTableObject` above.
				break
			case SlideObjectType.hyperlink:
				// A relationship type, not a shape — it reaches the shape tree only as the `a:hlinkClick`
				// carried by whatever shape owns the link.
				break
			case SlideObjectType.online:
				// Unused anywhere in `src/`; online video is `SlideObjectType.media` with `mtype: 'online'`.
				break
			default: {
				// Compile-time only: with every member routed above, this arm's type is `never`, so a new
				// `SlideObjectType` fails to typecheck here until it is either rendered or listed as one
				// of the non-shape members. Nothing is emitted, which is what the old bare default did.
				const unrouted: never = slideItemObj._type
				void unrouted
				break
			}
		}
		return strSlideXml
	}

	slide._slideObjects.forEach((slideItemObj: SlideObject) => {
		strSlideXml += renderSlideObjectXml(slideItemObj)
	})

	// STEP 4: Add slide numbers (if any) last.
	// One past the highest id the allocator handed out — the same slot the group-child counter
	// used to reach. A hardcoded id here (formerly 25) aliases a shape or group-child id once a
	// slide holds enough objects, which PowerPoint repairs. Resolved here rather than inside the
	// emitter so a part with no slide number does not consume an id.
	//
	// A loop rather than `Math.max(1, ...shapeIds.values())`: spreading a collection as arguments
	// is an argument-count ceiling written as arithmetic, and `shapeIds` is sized by the slide.
	if (slide._slideNumberProps) {
		let maxShapeId = 1
		for (const id of shapeIds.values()) if (id > maxShapeId) maxShapeId = id
		strSlideXml += slideNumberPlaceholderXml(slide, slide._slideNumberProps, maxShapeId + 1)
	}

	// STEP 5: Close spTree and finalize slide XML
	strSlideXml += '</p:spTree>'
	strSlideXml += '</p:cSld>'

	// LAST: Return
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
	 * The Targets emitted so far *from the media loop below*, which is the only place the
	 * question is asked: a media item produces TWO rels sharing one Target, and the second is
	 * told from the first by the first already being there.
	 *
	 * Two things about the shape. It holds raw Targets rather than escaped ones, because this
	 * asks about the model and not about markup — the probe used to substring-scan the emitted
	 * XML for ` Target="…"`, which forced the caller to escape by hand with the same escaper the
	 * builder uses or the two drifted apart, and cost O(rels squared) in string scanning.
	 *
	 * And it is scoped to the media loop rather than to every rel. Scanning all of them meant a
	 * hyperlink could answer a question about media, which is not a hypothetical: an online video
	 * and a hyperlink to the same URL is an ordinary thing to author, and the hyperlink came
	 * first, so the video pair came out as two MS-media rels with no ECMA `video` rel at all and
	 * `<a:videoFile r:link>` pointing at the wrong type.
	 */
	const mediaTargets = new Set<string>()

	// STEP 1: Add all rels for this Slide
	slide._rels.forEach((rel: SlideRel) => {
		lastRid = Math.max(lastRid, rel.rId)
		if (isHyperlinkRel(rel)) {
			if (rel.data === 'slide') {
				rels.push(relationshipEl(rel.rId, SLIDE_REL, `slide${rel.Target}.xml`))
			} else {
				rels.push(externalHyperlinkRel(rel.rId, rel.Target))
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
		rels.push(relationshipEl(rel.rId, rel.isChartEx ? CHARTEX_REL : CHART_REL, rel.Target))
	})
	;(slide._relsMedia || []).forEach((rel: SlideRelMedia) => {
		const relType = rel.type.toLowerCase()
		// Recorded at the point of emission, never earlier: "first one wins the ECMA type, second
		// gets the MS type" is the whole pairing rule for audio, video and online video, so the
		// probe must see only what is already out.
		const seen = mediaTargets.has(rel.Target)
		mediaTargets.add(rel.Target)
		const media = (type: string, targetMode?: string): string =>
			relationshipEl(rel.rId, type, rel.Target, { targetMode })
		lastRid = Math.max(lastRid, rel.rId)
		if (rel.oleRelType) {
			// An OLE payload part carries its rel type verbatim (`.../package` or `.../oleObject`);
			// its `type` is the part's content type, which the sniffing below would misread.
			rels.push(media(rel.oleRelType))
		} else if (rel.model3dRelType) {
			// Likewise a 3D model payload: its content type is `model/gltf.binary`, which matches none
			// of the sniffs below — and they have no `else`, so without this branch the rel would be
			// silently dropped and the slide would carry a dangling `r:embed`.
			rels.push(media(rel.model3dRelType))
		} else if (relType.includes('image')) {
			rels.push(media(IMAGE_REL))
		} else if (relType.includes('audio')) {
			rels.push(seen ? media(MS_MEDIA_REL) : media(AUDIO_REL))
		} else if (relType.includes('video')) {
			rels.push(seen ? media(MS_MEDIA_REL) : media(VIDEO_REL))
		} else if (relType.includes('online')) {
			// Online video has *TWO* external rels sharing the link Target: the ECMA video
			// rel (first) and the MS-2007 media rel (second). Both TargetMode="External",
			// no media binary part.
			rels.push(seen ? media(MS_MEDIA_REL, 'External') : media(VIDEO_REL, 'External'))
		}
	})

	// STEP 2: Add default rels
	defaultRels.forEach((rel, idx) => {
		rels.push(relationshipEl(lastRid + idx + 1, rel.type, rel.target))
	})

	return relationshipsPart(rels)
}
