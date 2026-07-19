/**
 * PptxGenJS: slide `<p:timing>` tree
 *
 * Build the slide-level timing tree that drives looping embedded media and the
 * preset build-animation `mainSeq`. A slide has at most one `<p:timing>`; all
 * looping media and animation effects share its `tmRoot` node.
 */

import { SlideObjectType } from '../../core-enums.js'
import type { AnimationProps, PresSlideInternal, SlideObject } from '../../core-interfaces.js'
import { el, raw, voidEl } from '../oxml/el.js'
import { collectSlideShapeIds } from '../slide/shape-ids.js'
import { buildAnimationSeq, buildBldList, resolveAnimationSpid } from './animation.js'

/**
 * Build the slide-level `<p:timing>` tree that makes embedded media loop.
 * - PowerPoint stores playback looping as `repeatCount` on the media node's `<p:cTn>`
 *   (`indefinite` for "Loop until Stopped", or `N*1000` for a finite N plays), inside
 *   the slide timing tree rather than on the `<p:pic>` itself.
 * - A slide has at most one `<p:timing>`; all looping media share its `tmRoot` node.
 * - Audio loops via `<p:audio>`, video via `<p:video>` (both `CT_TLCommonMediaNodeData`).
 * - The media node targets the picture by `spid` (its `<p:cNvPr>` id = slide-object index + 2).
 * @param {PresSlideInternal} slide - the slide to inspect for looping media
 * @returns {string} the `<p:timing>` XML, or `''` when no media loops
 */
export function slideTimingToXml(slide: PresSlideInternal): string {
	const loopMedia = slide._slideObjects.filter(
		(obj) =>
			obj._type === SlideObjectType.media &&
			obj.mtype !== 'online' &&
			typeof obj.mediaRid === 'number' &&
			(obj.loop === true || (typeof obj.loopCount === 'number' && obj.loopCount > 0))
	)

	// Resolve preset build animations to their target shape ids. `collectSlideShapeIds` covers group
	// children as well as top-level objects, so `objectName` addresses any shape on the slide.
	const shapeIds = collectSlideShapeIds(slide._slideObjects)
	const animations = (slide._animations ?? [])
		.map((anim) => ({ anim, spid: resolveAnimationSpid(shapeIds, slide._slideObjects.length, anim) }))
		.filter((entry): entry is { anim: AnimationProps; spid: number } => entry.spid !== null)

	const mediaNode = (obj: SlideObject, nodeId: number): string => {
		// spid must equal the picture's <p:cNvPr> id, which is the slide-object index + 2
		// (same basis as animation spids). Using mediaRid + 2 here desyncs from the shape id
		// and targets the wrong/nonexistent shape => PowerPoint reports the file corrupt.
		const spid = slide._slideObjects.indexOf(obj) + 2
		const repeatCount = obj.loop === true ? 'indefinite' : String(Math.round((obj.loopCount as number) * 1000))
		// EG_TimeNodeChoice: audio loops via <p:audio>, video via <p:video> (both CT_TLCommonMediaNodeData)
		const mediaTag = obj.mtype === 'audio' ? 'p:audio' : 'p:video'
		return el(
			mediaTag,
			null,
			raw(
				el('p:cMediaNode', null, [
					raw(
						el(
							'p:cTn',
							{ id: nodeId, repeatCount, fill: 'hold', display: 0 },
							raw(el('p:stCondLst', null, raw(voidEl('p:cond', { delay: 'indefinite' }))))
						)
					),
					raw(el('p:tgtEl', null, raw(voidEl('p:spTgt', { spid })))),
				])
			)
		)
	}

	// Media-only path: unchanged legacy output (`<p:cTn id="1">` tmRoot, media nodes follow).
	if (animations.length === 0) {
		if (loopMedia.length === 0) return ''
		let nodeId = 1
		const mediaNodes = loopMedia.map((obj) => mediaNode(obj, (nodeId += 1))).join('')
		return el(
			'p:timing',
			null,
			raw(
				el(
					'p:tnLst',
					null,
					raw(
						el(
							'p:par',
							null,
							raw(
								el(
									'p:cTn',
									{ id: 1, dur: 'indefinite', restart: 'never', nodeType: 'tmRoot' },
									raw(el('p:childTnLst', null, raw(mediaNodes)))
								)
							)
						)
					)
				)
			)
		)
	}

	// Animation path: a `mainSeq` of preset effects (plus any looping media as sibling nodes).
	// A monotonically increasing id is assigned to every `<p:cTn>` in document order, matching
	// how PowerPoint numbers the tree (tmRoot=1, mainSeq=2, then wrappers/effects/behaviors).
	let id = 2 // tmRoot=1, mainSeq=2 are fixed
	const next = (): number => (id += 1)

	const seq = buildAnimationSeq(animations, next)
	const mediaNodes = loopMedia.map((obj) => mediaNode(obj, next())).join('')
	const bldLst = buildBldList(animations)

	return el('p:timing', null, [
		raw(
			el(
				'p:tnLst',
				null,
				raw(
					el(
						'p:par',
						null,
						raw(
							el(
								'p:cTn',
								{ id: 1, dur: 'indefinite', restart: 'never', nodeType: 'tmRoot' },
								raw(el('p:childTnLst', null, raw(seq + mediaNodes)))
							)
						)
					)
				)
			)
		),
		raw(bldLst),
	])
}
