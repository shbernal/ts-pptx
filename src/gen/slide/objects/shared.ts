/**
 * ts-pptx: helpers shared by more than one slide-object renderer
 *
 * The `<p:cNvPr>` open tag, its `<a:hlinkClick>` children and the `<a:ln>` outline are each
 * emitted by several shape kinds. They live here — rather than in the dispatch module — so a
 * renderer never has to import from its own caller.
 */

import type { SlideObjectType } from '../../../enums.js'
import type { ShapeLineProps } from '../../../types/index.js'
import type {
	HyperlinkPropsInternal,
	ObjectOptionsInternal,
	PresSlideInternal,
	SlideLayoutInternal,
	SlideMasterInternal,
	SlideObject,
} from '../../../types/internal.js'
import { encodeXmlAttrValue } from '../../utils.js'
import { createLineCap, genXmlLineFill, lineEndEl, resolveDash } from '../../drawingml/line.js'
import { lineWidthToEmu, mapStated } from '../../../units-internal.js'
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
 * Everything a renderer needs is here, and that is now a requirement rather than an observation:
 * a renderer is called through {@link RendererTable}, so it takes this and nothing else. Two used
 * to take a second argument. The connector's `shapeIds` map is genuine context and moved onto this
 * interface; the image's `{ imgWidth, imgHeight }` was `frame.cx`/`frame.cy` under another name —
 * the dispatch captured them from the same two variables the frame is built from, two statements
 * apart — so it reads the frame instead.
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
	/**
	 * Every object's id, not just this one's — the whole map `shapeId` was read out of.
	 *
	 * A forward reference resolves through it: a connector names its endpoints by `objectName`
	 * and has to answer with the id the target will be emitted with, which the walk has not
	 * reached yet.
	 */
	shapeIds: ReadonlyMap<SlideObject, number>
	/** The page it belongs to — a slide, a layout or the master. */
	slide: PresSlideInternal | SlideLayoutInternal | SlideMasterInternal
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
 * The `SlideObjectType` members that a renderer emits a shape for.
 *
 * Not every member is one. `group` recurses back into the walk and owns the frame and bounds
 * logic that belongs to the dispatch, so it stays there; `notes`, `tablecell`, `hyperlink` and
 * `online` are not shapes at all and emit nothing. Those five are spelled out in the dispatch
 * beside this union's members, which is what makes a new enum member fail to compile until
 * someone says which kind it is.
 */
export type RenderedObjectType =
	| SlideObjectType.chart
	| SlideObjectType.connector
	| SlideObjectType.image
	| SlideObjectType.media
	| SlideObjectType.model3d
	| SlideObjectType.oleObject
	| SlideObjectType.placeholder
	| SlideObjectType.table
	| SlideObjectType.text
	| SlideObjectType.zoom

/** One shape family's emitter: everything it needs is in the {@link RenderContext}. */
export type ObjectRenderer = (ctx: RenderContext) => string

/**
 * Which renderer emits each shape family — the dispatch's whole knowledge of what a family's XML
 * looks like, supplied to it from outside.
 *
 * It has to come from outside. A named import inside a reachable function body is retained
 * unconditionally, so a `switch` that calls `renderChartObject` links the chart emitter into every
 * program that writes a slide, text-only ones included; a table built in the dispatch module would
 * import the same ten names and change nothing. Passed down the write path alongside the deck
 * state rather than registered into a module-level map: a mutable global populated at import time
 * is how the same input came to produce different bytes once already (see the chart-part-id
 * counter in `package/assemble.ts`), and it would also stop two decks built from different sets of
 * families in one process from each getting their own.
 *
 * `Partial` because a caller may legitimately supply fewer than ten. An object whose family is
 * absent is a routing bug rather than a silent omission, and the dispatch throws.
 */
export type RendererTable = Partial<Record<RenderedObjectType, ObjectRenderer>>

/**
 * The `<p:cNvPr>` OPEN tag shared by every shape renderer. Callers append `/>` or
 * `>`+children+`</p:cNvPr>` — the element is self-closing for some shape kinds and paired
 * (hyperlink / media-action children) for others.
 *
 * Not built with the element builder, because callers append the closing delimiter. Both `name` and
 * `descr` are escaped here and nowhere else: a slide object stores its `objectName` as the caller
 * wrote it, so lookups by name and `slide.objects` compare and report the caller's own spelling.
 * @param id - the shape's `<p:cNvPr>` id, unique slide-wide
 * @param name - the object's `objectName`, raw (escaped here)
 * @param descr - alt text, raw (escaped here)
 * @param openPrefix - byte-significant indentation before `<p:cNvPr`
 * @returns the open tag, without its closing delimiter
 */
export function cNvPrOpen(id: number, name: string | undefined, descr: string, openPrefix = ''): string {
	return `${openPrefix}<p:cNvPr id="${id}" name="${encodeXmlAttrValue(name ?? '')}" descr="${encodeXmlAttrValue(descr)}"`
}

/**
 * The `<a:hlinkClick>` child of a shape's `<p:cNvPr>`: a URL link, a jump to another slide, or a
 * slide-show action. Shared by the text and image renderers, which emitted identical copies.
 *
 * `<p:cNvPr>` holds at most one, so one destination is written, in that order of precedence. The
 * three used to be concatenated, which put two in one element for a link stating both `url` and
 * `slide`; `validateHyperlink` now refuses that link before it gets here.
 *
 * NOTE: the tooltip is passed through UNESCAPED and escaped once by the element builder. Escaping
 * it here as well would emit `&amp;amp;` for a tooltip containing `&`.
 * @param link - the shape's hyperlink, if any
 * @returns zero or one `<a:hlinkClick>` element
 */
export function cNvPrHyperlink(link: HyperlinkPropsInternal | undefined): string {
	if (!link) return ''
	const tooltip = link.tooltip ?? ''
	if (link.url) return voidEl('a:hlinkClick', { 'r:id': `rId${link._rId}`, tooltip })
	if (link.slide)
		return voidEl('a:hlinkClick', { 'r:id': `rId${link._rId}`, tooltip, action: 'ppaction://hlinksldjump' })
	// Action buttons: a self-contained slide-show navigation action. No relationship, so `r:id`
	// is emitted empty (schema-optional on CT_Hyperlink; matches PowerPoint's own output).
	if (link.action)
		return voidEl('a:hlinkClick', { 'r:id': '', tooltip, action: `ppaction://hlinkshowjump?jump=${link.action}` })
	return ''
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
	return el('a:ln', { w: mapStated(ln.width, lineWidthToEmu) ?? null, cap: ln.cap ? createLineCap(ln.cap) : null }, [
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
 * A part's `<p:spTree>`: the implicit top-level group's non-visual properties (the reserved
 * `cNvPr id="1"`) and an identity group transform, zeroed, then the shapes.
 *
 * This is the part's built-in root group, not a user-authored `addGroup`, hence the zeroed frame.
 * Every slide, layout, master, notes slide and notes master opens its tree with exactly these
 * bytes, and the slide emitter and the notes emitter each used to write them out.
 * @param children - the shapes, already serialized, in document order
 */
export function spTreeEl(children: readonly string[]): string {
	return el('p:spTree', null, [
		raw(
			el('p:nvGrpSpPr', null, [
				raw(voidEl('p:cNvPr', { id: 1, name: '' })),
				raw(voidEl('p:cNvGrpSpPr')),
				raw(voidEl('p:nvPr')),
			])
		),
		raw(el('p:grpSpPr', null, raw(grpXfrmEl({ x: 0, y: 0, cx: 0, cy: 0 })))),
		...children.map((child) => raw(child)),
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
