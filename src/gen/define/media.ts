/**
 * ts-pptx: Media Definition
 *
 * `addMediaDefinition` registers an `addMedia()` audio / video (or online video) source: each
 * embedded A/V consumes two rels (ECMA video + MS-2007 media) plus a cover image; online videos
 * use the external-link variant. The timing / `<p:pic>` XML is emitted later.
 */
import { SlideObjectType } from '../../enums.js'
import { IMG_PLAYBTN } from '../../media/placeholders.js'
import type { MediaProps } from '../../types/index.js'
import type { PresSlideInternal, SlideObject } from '../../types/internal.js'
import { encodeXmlAttrValue, getNewRelId, validateObjectName } from '../utils.js'
import { nextObjectNameIdx } from './object-name.js'
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
	const intPosX = opt.x || 0
	const intPosY = opt.y || 0
	const intSizeX = opt.w || 2
	const intSizeY = opt.h || 2
	const strData = opt.data || ''
	const strLink = opt.link || ''
	const strPath = opt.path || ''
	const strType = opt.type || 'audio'
	let strExtn = ''
	const strCover = opt.cover || IMG_PLAYBTN
	const mediaNameIdx = nextObjectNameIdx(target, SlideObjectType.media)
	const objectName = opt.objectName
		? encodeXmlAttrValue(validateObjectName(opt.objectName, 'media'))
		: `Media ${mediaNameIdx}`
	const slideData: SlideObject = { _type: SlideObjectType.media }

	// STEP 1: REALITY-CHECK
	if (!strPath && !strData && strType !== 'online') {
		throw new InvalidOptionError('media/missing-source', 'addMedia(): either `data` or `path` are required!')
	} else if (strData && !strData.toLowerCase().includes('base64,')) {
		throw new InvalidOptionError(
			'media/missing-base64-header',
			"addMedia(): `data` value lacks a base64 header! Ex: 'video/mpeg;base64,NMP[...]')"
		)
	} else if (strCover && !strCover.toLowerCase().includes('base64,')) {
		throw new InvalidOptionError(
			'media/cover-missing-base64-header',
			"addMedia(): `cover` value lacks a base64 header! Ex: 'data:image/png;base64,iV[...]')"
		)
	}
	// Online Video: requires `link`
	if (strType === 'online' && !strLink) {
		throw new InvalidOptionError('media/online-missing-link', 'addMedia(): online videos require `link` value')
	}

	strExtn = opt.extn || (strData ? (strData.split(';')[0] ?? '').split('/')[1] : strPath.split('.').pop()) || 'mp3'

	// STEP 2: Set type, media
	slideData.mtype = strType
	slideData.media = strPath || 'preencoded.mov'
	slideData.options = {}

	// Playback looping (embedded audio/video only; online embeds have no timing tree)
	if (strType !== 'online') {
		if (opt.loop) slideData.loop = true
		else if (typeof opt.loopCount === 'number' && opt.loopCount > 0) slideData.loopCount = opt.loopCount
	}

	// STEP 3: Set media properties & options
	slideData.options.x = intPosX
	slideData.options.y = intPosY
	slideData.options.w = intSizeX
	slideData.options.h = intSizeY
	slideData.options.objectName = objectName
	if (opt.altText) slideData.options.altText = opt.altText
	if (opt.objectLock) slideData.options.objectLock = opt.objectLock

	// STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
	/**
	 * NOTE:
	 * - rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
	 *
	 * NOTE:
	 * - Audio/Video files consume *TWO* rId's:
	 * <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
	 * <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
	 */
	if (strType === 'online') {
		const relId1 = getNewRelId(target)
		// A: ECMA video rel (external link) — referenced by <a:videoFile r:link>.
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			data: 'dummy',
			type: 'online',
			extn: strExtn,
			rId: relId1,
			Target: strLink,
		})
		slideData.mediaRid = relId1

		// B: MS-2007 media rel — PowerPoint authors a second external rel sharing the
		// same link Target; the body points at it via <p14:media r:link>. (Mirrors the
		// embedded A/V pair, but External and with no media binary part.)
		const relId2 = getNewRelId(target)
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			data: 'dummy',
			type: 'online',
			extn: strExtn,
			rId: relId2,
			Target: strLink,
		})

		// C: Add cover (preview/overlay) image
		const relId3 = getNewRelId(target)
		target._relsMedia.push({
			path: 'preencoded.png',
			data: strCover,
			type: 'image/png',
			extn: 'png',
			rId: relId3,
			Target: `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.png`,
		})
		assertConsecutiveMediaRids(relId1, relId2, relId3)
	} else {
		// PERF: Duplicate media should reuse existing `Target` value and not create an additional copy.
		// Path-based media match by `path`; base64/`data` media (which share the `preencoded`
		// placeholder path) match by their data payload so identical inline media embed once.
		const dupeItem = target._relsMedia.find((item) => {
			if (item.isDuplicate || !item.Target || item.type !== strType + '/' + strExtn) return false
			return strPath ? item.path === strPath : !!strData && item.data === strData
		})

		// A: "relationships/video"
		const relId1 = getNewRelId(target)
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: relId1,
			isDuplicate: !!dupeItem?.Target,
			Target: dupeItem?.Target
				? dupeItem.Target
				: `../media/media-${target._slideNum}-${target._relsMedia.length + 1}.${strExtn}`,
		})
		slideData.mediaRid = relId1

		// B: "relationships/media"
		const relId2 = getNewRelId(target)
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: relId2,
			isDuplicate: !!dupeItem?.Target,
			Target: dupeItem?.Target
				? dupeItem.Target
				: `../media/media-${target._slideNum}-${target._relsMedia.length + 0}.${strExtn}`,
		})

		// C: Add cover (preview/overlay) image
		const relId3 = getNewRelId(target)
		target._relsMedia.push({
			path: 'preencoded.png',
			type: 'image/png',
			extn: 'png',
			data: strCover,
			rId: relId3,
			Target: `../media/image-${target._slideNum}-${target._relsMedia.length + 1}.png`,
		})
		assertConsecutiveMediaRids(relId1, relId2, relId3)
	}

	// LAST
	target._slideObjects.push(slideData)
}
