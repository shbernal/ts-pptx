/**
 * ts-pptx: OLE-object slide-object serialization
 */

import { genXmlObjectLock, GRAPHIC_FRAME_LOCK_ATTRS } from '../../drawingml/locks.js'
import { el, raw, voidEl, type XmlAttrs, type XmlChild } from '../../oxml/el.js'
import { type RenderContext, cNvPrOpen, graphicFrameEl, previewPicBody } from './shared.js'
import { OOXML_NS } from '../../../ooxml/namespaces.js'

/** VML namespace — declared by an OLE object's `mc:Choice Requires="v"` (no VML content is emitted). */
const VML_NS = 'urn:schemas-microsoft-com:vml'
/** graphicData URI for an embedded OLE object (`<p:oleObj>`). */
const OLE_NS = 'http://schemas.openxmlformats.org/presentationml/2006/ole'

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
export function renderOleObject(ctx: RenderContext): string {
	const {
		obj: slideItemObj,
		idx,
		frame: { x, y, cx, cy },
	} = ctx
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
		raw(previewPicBody(ole.previewRid, { x, y, cx, cy })),
	])

	const alternateContent = el('mc:AlternateContent', { 'xmlns:mc': OOXML_NS.mc }, [
		raw(el('mc:Choice', { 'xmlns:v': VML_NS, Requires: 'v' }, raw(oleObj([raw(voidEl('p:embed'))])))),
		raw(el('mc:Fallback', null, raw(oleObj([raw(voidEl('p:embed')), raw(previewPic)])))),
	])

	const nvGraphicFramePr = el('p:nvGraphicFramePr', null, [
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
	return graphicFrameEl({ nvGraphicFramePr, frame: { x, y, cx, cy }, uri: OLE_NS, payload: alternateContent })
}
