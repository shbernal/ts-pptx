/**
 * ts-pptx: helpers shared by more than one slide-object renderer
 *
 * The `<p:cNvPr>` open tag, its `<a:hlinkClick>` children and the `<a:ln>` outline are each
 * emitted by several shape kinds. They live here — rather than in the dispatch module — so a
 * renderer never has to import from its own caller.
 */

import type { ShapeLineProps } from '../../../types/index.js'
import type {
	HyperlinkPropsInternal,
	ObjectOptionsInternal,
	PresSlideInternal,
	SlideLayoutInternal,
	SlideObject,
} from '../../../types/internal.js'
import { encodeXmlAttrValue } from '../../utils.js'
import { createLineCap, genXmlLineFill, lineEndEl, resolveDash } from '../../drawingml/line.js'
import { lineWidthToEmu } from '../../../units-internal.js'
import { el, raw, voidEl, type XmlAttrs, type XmlFmt } from '../../oxml/el.js'
import { STRETCH_FILL_RECT } from '../../drawingml/src-rect.js'
import { prstGeomRect } from '../../drawingml/geometry.js'

/**
 * Everything the dispatch in `gen/slide/object.ts` has already resolved for one slide object,
 * handed to whichever renderer that object's `_type` selects.
 *
 * One object rather than a positional prefix because the prefix was not actually shared: `slide`
 * sat in position 3 for two of the nine renderers and was absent from the other seven, and
 * `renderImageObject` took six adjacent `number` parameters that no type could tell apart.
 * Transposing `cx` and `cy` there compiled, shipped, and produced a stretched picture that
 * nothing outside a rendered slide would have flagged. Named fields make that transposition
 * impossible to write.
 *
 * A renderer takes a second argument only for what genuinely is not here: the connector's
 * `shapeIds` map, and the image's natural pixel size (captured *before* a placeholder overrides
 * the frame, so it is not `frame.cx`/`frame.cy`).
 */
export interface RenderContext {
	/** The slide object being emitted. */
	obj: SlideObject
	/**
	 * Its `<p:cNvPr>` id, from the one allocator (`collectSlideShapeIds`).
	 *
	 * Every renderer used to recompute it as `idx + 2` off the walk index, and the group walk
	 * kept a second counter for children — so the map that answers a forward reference (a
	 * connector's `a:stCxn`, an animation's `p:spTgt spid`) and the ids actually emitted were two
	 * derivations that had to be kept in step by hand.
	 */
	shapeId: number
	/** The page it belongs to — a slide or a layout. */
	slide: PresSlideInternal | SlideLayoutInternal
	/** The resolved box in EMU: normalized for negative extents, then overridden by a placeholder. */
	frame: { x: number; y: number; cx: number; cy: number }
	/** The layout placeholder this object inherits from, or `null`. */
	placeholder: SlideObject | null
	/** `<a:xfrm>` placement attributes (`flipH`, `flipV`, `rot`); attribute order is byte-significant. */
	locationAttrs: XmlAttrs
	/**
	 * The object's options, already normalized by the dispatch. Read these rather than
	 * re-narrowing `obj.options`: each renderer has exactly one call site, and a contract stated
	 * there beats a defensive re-assignment in every callee.
	 */
	itemOpts: ObjectOptionsInternal
}

/**
 * The `<p:cNvPr>` OPEN tag shared by every shape renderer. Callers append `/>` or
 * `>`+children+`</p:cNvPr>` — the element is self-closing for some shape kinds and paired
 * (hyperlink / media-action children) for others.
 *
 * NOT built with the element builder, deliberately. `descr` is escaped here but `name` is **not**,
 * and that asymmetry is intentional and load-bearing: `objectName` is caller-supplied free text,
 * but every `add*Definition` (text.ts, shape.ts, image.ts, chart.ts, media.ts, connector.ts,
 * group.ts, table.ts) already runs it through `encodeXmlAttrValue(validateObjectName(...))` once
 * before it reaches a slide object's `options`. Escaping it again here would double-encode it
 * (`'Q&A'` -> `Q&amp;A` upstream -> `Q&amp;amp;A` if escaped here too). This helper exists so that
 * single escape stays one line in one place instead of eight call sites re-deriving it.
 * @param id - the shape's `<p:cNvPr>` id, unique slide-wide
 * @param name - caller-supplied `objectName`, already escaped once upstream (emitted as-is)
 * @param descr - alt text (escaped)
 * @param openPrefix - byte-significant indentation before `<p:cNvPr`
 * @returns the open tag, without its closing delimiter
 */
export function cNvPrOpen(id: number, name: string | undefined, descr: string, openPrefix = ''): string {
	return `${openPrefix}<p:cNvPr id="${id}" name="${name}" descr="${encodeXmlAttrValue(descr)}"`
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
export function cNvPrHyperlink(link: HyperlinkPropsInternal | undefined): string {
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
export function genXmlShapeLine(ln: ShapeLineProps): string {
	return el('a:ln', { w: ln.width ? lineWidthToEmu(ln.width) : null, cap: ln.cap ? createLineCap(ln.cap) : null }, [
		raw(genXmlLineFill(ln)),
		ln.dashType ? raw(voidEl('a:prstDash', { val: resolveDash(ln.dashType, 'solid', 'line: dashType') })) : null,
		raw(lineEndEl('a:headEnd', ln.beginArrowType, 'line: beginArrowType')),
		raw(lineEndEl('a:tailEnd', ln.endArrowType, 'line: endArrowType')),
	])
}

/** A box in EMU, the one argument every transform emitter takes. */
export interface XfrmFrame {
	x: number | string
	y: number | string
	cx: number | string
	cy: number | string
}

/**
 * An `<a:off>`/`<a:ext>` transform, under whichever wrapper the context calls for: `p:xfrm` on a
 * `<p:graphicFrame>`, `a:xfrm` inside a `<p:spPr>`. Fourteen sites across the shape, chart,
 * connector, image, media, notes-master, text and zoom emitters wrote the same six-line pair;
 * `attrs` is what the last of them needed (a shape's `rot`/`flipH`/`flipV` ride on the `a:xfrm`
 * itself, not on its children).
 * @param tag - `p:xfrm` or `a:xfrm`
 * @param frame - the box, in EMU
 * @param attrs - attributes on the transform element itself, e.g. rotation and flips
 * @param fmt - byte-significant layout, where the caller's part indents
 */
export function xfrmEl(
	tag: 'p:xfrm' | 'a:xfrm',
	frame: XfrmFrame,
	attrs: XmlAttrs | null = null,
	fmt?: XmlFmt
): string {
	return el(
		tag,
		attrs,
		[raw(voidEl('a:off', { x: frame.x, y: frame.y })), raw(voidEl('a:ext', { cx: frame.cx, cy: frame.cy }))],
		fmt
	)
}

/**
 * A group's `<a:xfrm>`: the box, then an IDENTITY child coordinate space (`chOff`/`chExt` equal
 * to `off`/`ext`).
 *
 * The identity is the contract `docs/groups.md` and the measured-fit solver both rest on — a
 * group never scales its children, so a grouped shape's authored size is its rendered size.
 * Three group emitters spelled the four-child block out.
 * @param frame - the box, in EMU
 * @param attrs - attributes on the `a:xfrm` itself
 */
export function grpXfrmEl(frame: XfrmFrame, attrs: XmlAttrs | null = null): string {
	return el('a:xfrm', attrs, [
		raw(voidEl('a:off', { x: frame.x, y: frame.y })),
		raw(voidEl('a:ext', { cx: frame.cx, cy: frame.cy })),
		raw(voidEl('a:chOff', { x: frame.x, y: frame.y })),
		raw(voidEl('a:chExt', { cx: frame.cx, cy: frame.cy })),
	])
}

/**
 * A `<p:graphicFrame>`: its non-visual properties, its transform, and the `<a:graphic>` /
 * `<a:graphicData>` envelope around a payload.
 *
 * The child order is the point. `CT_GraphicalObjectFrame` sequences
 * `nvGraphicFramePr` → `xfrm` → `graphic`, and PowerPoint reports a frame that gets it wrong as
 * a corrupt file (0x80070570) rather than as a bad element — a failure the five emitters that
 * build one (chart, table, OLE, zoom, 3D model) each had to be right about separately.
 *
 * `nvGraphicFramePr` stays the caller's to build and is *not* folded in here. The five differ in
 * every part of it — an empty `<p:cNvGraphicFramePr>` for a chart against four different
 * `<a:graphicFrameLocks>` default sets, a `<p:nvPr>` that is empty for three and carries a
 * placeholder (and, for a table, a `p14:modId` extension) for the other two — and each of those
 * differences is deliberate and documented where it is made.
 *
 * @param opts.nvGraphicFramePr - the already-serialized `<p:nvGraphicFramePr>`
 * @param opts.frame - the frame's slide-absolute box, in EMU
 * @param opts.uri - the `<a:graphicData>` payload namespace
 * @param opts.payload - the already-serialized `<a:graphicData>` content
 * @param opts.fmt - per-element byte layout; `graphicAttrs` adds attributes to `<a:graphic>`
 *   (the chart frame alone redeclares `xmlns:a` there, as PowerPoint writes it)
 */
export function graphicFrameEl(opts: {
	nvGraphicFramePr: string
	frame: { x: number; y: number; cx: number; cy: number }
	uri: string
	payload: string
	fmt?: { xfrm?: XmlFmt; graphic?: XmlFmt; graphicData?: XmlFmt; graphicAttrs?: XmlAttrs }
}): string {
	const { nvGraphicFramePr, frame, uri, payload, fmt } = opts
	return el('p:graphicFrame', null, [
		raw(nvGraphicFramePr),
		raw(xfrmEl('p:xfrm', frame, null, fmt?.xfrm)),
		raw(
			el(
				'a:graphic',
				fmt?.graphicAttrs ?? null,
				raw(el('a:graphicData', { uri }, raw(payload), fmt?.graphicData)),
				fmt?.graphic
			)
		),
	])
}

/**
 * The `<a:picLocks>` set PowerPoint fixes on an `mc:Fallback` preview picture.
 *
 * Fixed rather than taken from the caller's `objectLock`, and that is true of every construct
 * that emits one: the fallback picture is what a consumer *without* the feature draws, while the
 * caller's locks belong on the `mc:Choice` frame — the object PowerPoint itself manipulates. The
 * two element types also accept different flags, so folding a `graphicFrameLocks` set onto a
 * `picLocks` would warn about every flag they do not share. Routed through `genXmlObjectLock` so
 * attribute order comes from `PICTURE_LOCK_ATTRS` (`gen/drawingml/locks.ts`) rather than from a literal, and a flag
 * added to the table lands in the right place.
 *
 * A 3D model spreads `noCrop` on top: it is reframed by its camera, never by cropping the cached
 * raster.
 */
export const FALLBACK_PICTURE_LOCKS = Object.freeze({
	noGrp: true,
	noRot: true,
	noChangeAspect: true,
	noMove: true,
	noResize: true,
	noEditPoints: true,
	noAdjustHandles: true,
	noChangeArrowheads: true,
	noChangeShapeType: true,
})

/**
 * The drawn half of an `mc:Fallback` preview picture: the `<p:blipFill>` onto the cached image
 * and the `<p:spPr>` that places it. Identical in the OLE, zoom and 3D-model emitters, which is
 * everything a preview picture is apart from its `<p:nvPicPr>` — and there the three differ on
 * purpose (an id-less `cNvPr` for OLE, a hyperlinked one for a zoom tile, `noCrop` for a model),
 * so each keeps its own.
 * @param previewRid - the relationship id of the cached image
 * @param frame - where the picture is drawn, in EMU
 * @param outline - `true` for the hairline grey border a zoom tile carries
 * @param prefix - which namespace the two wrapper elements take; a zoom tile's live in `p166`
 */
/** Which namespace `previewPicBody`'s two wrapper elements take, and whether to declare it. */
interface PreviewPicPrefix {
	ns: string
	nsUri?: string
}

export function previewPicBody(
	previewRid: number,
	frame: XfrmFrame,
	outline = false,
	prefix: PreviewPicPrefix = { ns: 'p' }
): string {
	const attrs = prefix.nsUri === undefined ? null : { [`xmlns:${prefix.ns}`]: prefix.nsUri }
	return (
		el(`${prefix.ns}:blipFill`, attrs, [
			raw(voidEl('a:blip', { 'r:embed': `rId${previewRid}` })),
			raw(STRETCH_FILL_RECT),
		]) +
		el(`${prefix.ns}:spPr`, attrs, [
			raw(xfrmEl('a:xfrm', frame)),
			raw(prstGeomRect()),
			outline
				? raw(el('a:ln', { w: '3175' }, raw(el('a:solidFill', null, raw(voidEl('a:prstClr', { val: 'ltGray' }))))))
				: null,
		])
	)
}
