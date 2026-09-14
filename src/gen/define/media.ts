/**
 * ts-pptx: Media Definition
 *
 * `addMediaDefinition` registers an `addMedia()` audio / video (or online video) source: each
 * embedded A/V consumes two rels (ECMA video + MS-2007 media) plus a cover image; online videos
 * use the external-link variant. The timing / `<p:pic>` XML is emitted later.
 */
import { SlideObjectType } from '../../enums.js'
import { warn } from '../../diagnostics.js'
import type { MediaProps } from '../../types/index.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { getNewRelId, preencodedPath } from '../utils.js'
import { hasBase64Header } from '../../media/base64.js'
import { pushMediaRel } from './image-rel.js'
import { framedObjectOptions, requirePayloadSource } from './object-options.js'
import { InternalError, InvalidOptionError } from '../../errors.js'

/**
 * One media item costs three consecutive slide rels, allocated together by {@link addMedia}:
 * the ECMA `audio`/`video` rel (the id kept on the object as `mediaRid`), the MS-2007 `media`
 * rel sharing its Target, and the preview/poster image.
 *
 * The two offsets used to be bare arithmetic wherever a body or a descriptor needed one of the
 * other two rels, in three modules; `gen/anim/timing.ts` even carries a comment warning against
 * writing `mediaRid + 2` in a place where it means something else entirely. Naming them puts
 * the layout in one place, and {@link assertConsecutiveMediaRids} makes the assumption they
 * rest on fail loudly instead of quietly emitting a body that points at the wrong rel.
 */
export const msMediaRid = (mediaRid: number): number => mediaRid + 1

/** The preview/poster image rel's id; see {@link msMediaRid}. */
export const previewRid = (mediaRid: number): number => mediaRid + 2

/**
 * Check that a media item's three rels really did come out consecutive. The ids come from
 * three separate `getNewRelId` calls, so nothing but call order makes them so — and every
 * reader of the triple assumes it.
 */
function assertConsecutiveMediaRids(base: number, second: number, third: number): void {
	if (second !== msMediaRid(base) || third !== previewRid(base))
		throw new InternalError(
			'media/rel-ids-not-consecutive',
			`addMedia expected rel ids ${base}, ${msMediaRid(base)}, ${previewRid(base)}; got ${base}, ${second}, ${third}`
		)
}

/**
 * Adds a media object to a slide definition.
 * @param {PresSlideInternal} `target` - slide object that the media will be added to
 * @param {MediaProps} `opt` - media options
 */
export function addMediaDefinition(target: PresSlideInternal, opt: MediaProps): void {
	const strData = opt.data || ''
	const strLink = opt.link || ''
	const strPath = opt.path || ''
	const strType = opt.type || 'audio'
	// Empty when the caller passed no cover: the default play-button poster is resolved later,
	// by the async media pass, so that a deck carrying no media never links the artwork. The
	// base64-header check below still guards a cover the caller *did* pass.
	const strCover = opt.cover || ''
	const slideData: SlideObject = { _type: SlideObjectType.media }

	// STEP 1: REALITY-CHECK, before the object name below takes its index
	if (strType !== 'online') requirePayloadSource(opt, 'media/missing-source', 'addMedia')
	if (strData && !hasBase64Header(strData)) {
		throw new InvalidOptionError(
			'media/missing-base64-header',
			"addMedia(): `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')"
		)
	} else if (strCover && !hasBase64Header(strCover)) {
		throw new InvalidOptionError(
			'media/cover-missing-base64-header',
			"addMedia(): `cover` value lacks a base64 header! Ex: 'data:image/png;base64,iV[...]')"
		)
	}
	// Online Video: requires `link`
	if (strType === 'online' && !strLink) {
		throw new InvalidOptionError('media/online-missing-link', 'addMedia(): online videos require `link` value')
	}
	const options = framedObjectOptions(target, SlideObjectType.media, opt, {
		label: 'Media',
		kind: 'media',
		api: 'addMedia',
		defaults: { x: 0, y: 0, w: 2, h: 2 },
	})

	const strExtn =
		opt.extn || (strData ? (strData.split(';')[0] ?? '').split('/')[1] : strPath.split('.').pop()) || 'mp3'

	// STEP 2: Set type, media and the object's options
	slideData.mtype = strType
	slideData.media = strPath || preencodedPath('mov')
	slideData.options = options

	// Playback looping (embedded audio/video only; online embeds have no timing tree)
	if (strType !== 'online') {
		if (opt.loop) slideData.loop = true
		else if (typeof opt.loopCount === 'number' && Number.isFinite(opt.loopCount) && opt.loopCount > 0)
			slideData.loopCount = opt.loopCount
		// A stated count that plays nothing used to be dropped without a word, and `Infinity` was
		// written as `repeatCount="Infinity"`.
		else if (opt.loopCount != null)
			warn(
				'media/invalid-loop-count',
				`addMedia(): \`loopCount\` is how many times to play, a finite number above 0; got ${String(opt.loopCount)}, so the media plays once. Use \`loop: true\` to repeat it without end.`
			)
	}

	// STEP 3: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
	/**
	 * NOTE:
	 * - rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
	 *
	 * NOTE:
	 * - Audio/Video files consume *TWO* rId's:
	 * <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
	 * <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
	 */
	// A and B are the two rels an online video registers differently from an embedded one.
	let mediaRid: number
	let msRid: number
	if (strType === 'online') {
		// A: ECMA video rel (external link) — referenced by <a:videoFile r:link>.
		const online = { kind: 'media', extn: strExtn, type: 'online', path: strPath, data: 'dummy' }
		mediaRid = getNewRelId(target)
		pushMediaRel(target, { ...online, rId: mediaRid, target: { external: strLink } })

		// B: MS-2007 media rel — PowerPoint authors a second external rel sharing the
		// same link Target; the body points at it via <p14:media r:link>. (Mirrors the
		// embedded A/V pair, but External and with no media binary part.)
		msRid = getNewRelId(target)
		pushMediaRel(target, { ...online, rId: msRid, target: { external: strLink } })
	} else {
		// A: "relationships/video". An identical source already on the slide — the same path, or the
		// same inline payload — lends its part rather than embedding the bytes twice.
		const embedded = { kind: 'media', extn: strExtn, type: strType + '/' + strExtn, path: strPath, data: strData }
		mediaRid = getNewRelId(target)
		const video = pushMediaRel(target, { ...embedded, rId: mediaRid, dedupe: true })

		// B: "relationships/media", against the one part the `video` rel above named.
		msRid = getNewRelId(target)
		pushMediaRel(target, { ...embedded, rId: msRid, target: { sameAs: video } })
	}
	slideData.mediaRid = mediaRid

	// C: Add cover (preview/overlay) image
	const coverRid = getNewRelId(target)
	pushMediaRel(target, {
		kind: 'image',
		extn: 'png',
		type: 'image/png',
		data: strCover,
		rId: coverRid,
		extra: { isDefaultCover: !strCover },
	})
	assertConsecutiveMediaRids(mediaRid, msRid, coverRid)

	// LAST
	target._slideObjects.push(slideData)
}
