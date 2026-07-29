/**
 * ts-pptx: media (audio / video / online video) slide-object serialization
 *
 * Emits a `media` slide object as a `<p:pic>` carrying the required preview image. The online
 * and embedded forms differ in exactly two tokens — the EG_Media choice element and whether the
 * paired `<p14:media>` rel binds by `r:link` or `r:embed` — so they share one body.
 */

import type { SlideObject } from '../../../types/internal.js'
import { genXmlObjectLock, PICTURE_LOCK_ATTRS } from '../../drawingml/locks.js'
import { el, raw, voidEl, type XmlAttrs } from '../../oxml/el.js'
import { cNvPrOpen, P14_NS } from './shared.js'

/**
 * Render a `media` (audio/video/online) slide object to its `<p:pic>` XML.
 */
export function renderMediaObject(
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
