/**
 * ts-pptx: Zoom (Slide / Section / Summary) slide-object serialization
 *
 * Emits a `zoom` slide object as an `<mc:AlternateContent>`. The `mc:Choice` carries the real
 * `<p:graphicFrame>` in the 2016 zoom namespaces, read by PowerPoint 2016+; the `mc:Fallback`
 * carries a hyperlinked preview picture — or, for a Summary Zoom, a group of them — so older
 * consumers still get a clickable navigation thumbnail. See `gen/define/zoom.ts`.
 */

import type { SlideObject, ZoomInternal, ZoomTileInternal } from '../../../types/internal.js'
import { el, raw, voidEl, type XmlAttrs } from '../../oxml/el.js'
import { cNvPrOpen, MC_NS } from './shared.js'
import { PICTURE_LOCK_ATTRS, genXmlObjectLock } from '../../drawingml/locks.js'

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
						// PowerPoint's zoom-tile lock set. Fixed rather than caller-supplied: `SlideZoomProps`
						// has no `objectLock`, unlike image/media/OLE/table/shape/group. Routed through
						// `genXmlObjectLock` anyway, so the attribute order is the table's rather than this
						// literal's, and a flag added to the set lands in the right place.
						raw(
							genXmlObjectLock('a:picLocks', PICTURE_LOCK_ATTRS, {
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
export function renderZoomObject(
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
