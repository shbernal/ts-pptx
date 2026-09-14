/**
 * ts-pptx: the one place a media relationship is registered on a slide, layout or master.
 *
 * Every definer that embeds bytes pushes a media rel — images, image fills, previews, picture
 * bullets, backgrounds, audio and video, OLE payloads, 3D models, and the transition sound the
 * packager adds. They pushed the record by hand, each spelled slightly differently: two copies of
 * the dedupe match below, three spellings of the placeholder path, and one that put a base64
 * payload in the `path` field. {@link pushMediaRel} is the record, spelled once.
 */
import type { PresSlideInternal, SlideRelMedia } from '../../types/internal.js'
import { getNewRelId, nextMediaTarget, preencodedPath } from '../utils.js'
import { imageContentType, imageExtensionForSource } from '../../media/content-type.js'
import { hasBase64Header } from '../../media/base64.js'

/** An image source a media rel is registered from: the caller's `path` and `data`, and the extension its part is named with. */
export interface ImageSource {
	readonly path: string
	readonly data: string
	readonly extn: string
}

/** Why a caller's image source cannot be registered. */
export type ImageSourceProblem = 'missing-source' | 'missing-base64-header'

/**
 * Resolve a caller's `{ path, data }` to the source an image media rel is registered from, or say
 * why there is none.
 *
 * Every definer that embeds an image reads its source here: `addImage`, image fills, picture
 * bullets, backgrounds, and the covers and previews Zoom, OLE and 3D model objects are drawn from.
 * Each used to check its own subset, and three checked no header at all, so a cover of
 * `{ data: 'hello world' }` wrote a 7-byte `.png` part with nothing said.
 *
 * What a problem costs is the caller's to decide, because it depends on what the image is for: a
 * picture with nothing to degrade to throws, a cover falls back to a placeholder, and a fill, a
 * bullet or a background is dropped. This only reports it.
 * @param source - the caller's `path` and `data`, either of which may be absent
 */
export function resolveImageSource(
	source: { path?: string | undefined; data?: string | undefined } | undefined
): ImageSource | ImageSourceProblem {
	const path = source?.path || ''
	const data = source?.data || ''
	if (!path && !data) return 'missing-source'
	if (data && !hasBase64Header(data)) return 'missing-base64-header'
	return { path, data, extn: imageExtensionForSource(path, data) }
}

/** One media source, as {@link pushMediaRel} registers it. */
export interface MediaRelSource {
	/** The part name's leading segment: `image`, `media`, `audio`, `model3d`, `oleObject`. */
	kind: string
	/** The part's file extension, without the dot. */
	extn: string
	/** The part's content type — or `online` for an external video link, which writes no part. */
	type: string
	/** A file or URL the media pass loads the bytes from, when the source is one. */
	path?: string | undefined
	/** The bytes, as base64 or a `data:` URI, when they came inline. */
	data?: string | undefined
	/** The relationship id, already allocated. */
	rId: number
	/** The sibling directory the part lands in. */
	dir?: 'media' | 'embeddings'
	/**
	 * Point at the part an identical earlier source already has, rather than naming a new one.
	 *
	 * De-dup is per target and by source, not by bytes: a file-path source matches on `path`,
	 * while inline `data` sources have no real path — they all share the {@link preencodedPath}
	 * placeholder — so they match on their payload instead, which is what stops the same inline
	 * image being embedded once per use. A rel already marked `isDuplicate` is never matched
	 * against, so every duplicate points at the one original rather than at a chain.
	 *
	 * This is a *slide-local* optimization. `package/assemble.ts` runs a second, deck-wide collapse
	 * keyed on extension + bytes once every rel's data is loaded, which subsumes this one for reuse
	 * across slides and for sources that only turn out identical after loading.
	 */
	dedupe?: boolean
	/**
	 * Name no new part. `sameAs` points at the part another rel of this source already named, for
	 * the second of two rels one part needs (a video's ECMA and MS-2007 rels). `external` is the
	 * link an online video's rel carries as its Target, with no part behind it.
	 */
	target?: { sameAs: SlideRelMedia } | { external: string }
	/** Flags a few kinds carry for the passes after this one, copied onto the record as given. */
	extra?: Pick<SlideRelMedia, 'isSvgPng' | 'svgSize' | 'isDefaultCover' | 'oleRelType' | 'model3dRelType'>
}

/**
 * Register one media rel on `target` and return the record.
 *
 * The part name is read before the push: {@link nextMediaTarget} counts the media rels already on
 * the target, so the first one lands on `image-<key>-1`, and computing it after the push would
 * rename every part.
 * @param target - slide, layout or master the rel is registered on
 * @param source - what the rel points at, and how
 */
export function pushMediaRel(target: PresSlideInternal, source: MediaRelSource): SlideRelMedia {
	const path = source.path || ''
	const data = source.data ?? ''
	const linked = source.target && 'external' in source.target ? source.target.external : undefined
	const shared = source.target && 'sameAs' in source.target ? source.target.sameAs : undefined
	const dupe = source.dedupe
		? target._relsMedia.find((item) => {
				if (item.isDuplicate || !item.Target || item.type !== source.type) return false
				return path ? item.path === path : !!data && item.data === data
			})
		: undefined

	const rel: SlideRelMedia = {
		path: path || preencodedPath(source.extn),
		type: source.type,
		extn: source.extn,
		data,
		rId: source.rId,
		...(source.dedupe ? { isDuplicate: !!dupe?.Target } : shared ? { isDuplicate: !!shared.isDuplicate } : {}),
		Target:
			linked ??
			shared?.Target ??
			(dupe?.Target ? dupe.Target : nextMediaTarget(target, source.kind, source.extn, source.dir)),
		...source.extra,
	}
	target._relsMedia.push(rel)
	return rel
}

/**
 * Push an image media rel onto `target`, reusing an identical source's package part.
 *
 * Three definers need it — an image *fill* on a shape or text box (`registerImageFillMedia`), an
 * `addImage()` raster (`addImageDefinition`), and the cached preview raster a Zoom tile or OLE
 * object is drawn from (`registerPreviewImage`). See {@link MediaRelSource.dedupe} for the match.
 * @param target - slide (or layout/master) the rel is registered on
 * @param source - the resolved image source: a `path`, a base64 `data` payload, or both
 * @param relId - the relationship id already allocated for this use
 */
export function registerImageMediaRel(
	target: PresSlideInternal,
	source: { path?: string; data?: string; extn: string },
	relId: number
): void {
	pushMediaRel(target, {
		kind: 'image',
		extn: source.extn,
		type: imageContentType(source.extn),
		path: source.path || '',
		data: source.data || '',
		rId: relId,
		dedupe: true,
	})
}

/**
 * Push the *pair* of media rels an SVG source consumes: a rasterized PNG fallback (what a
 * renderer without SVG support paints, and what `<a:blip r:embed>` points at) and the SVG
 * itself (`asvg:svgBlip`).
 *
 * `addImage` and the picture-bullet definer both need this, and both used to take the SVG's
 * id as `pngRid + 1` on faith. That held only while nothing else allocated in between, and
 * `addImage`'s own hyperlink then took the same number a third time. Both ids now come from
 * {@link getNewRelId}, which skips every id the slide already holds.
 *
 * Neither push dedupes and neither needs to. The PNG fallback is rasterized per call from a
 * per-call `svgSize`, so two uses at different sizes are genuinely two different images; the SVG
 * source has no such excuse, but the deck-wide collapse in `package/assemble.ts` keys on
 * extension + bytes and merges them, so the same SVG placed twice measures as a single
 * `ppt/media/*.svg` part.
 *
 * @param target - slide (or layout/master) the rels are registered on
 * @param source - the resolved SVG source: a `path`, a base64 `data` payload, or both, plus
 *   the display size the PNG fallback is rasterized at when the caller knows it
 * @param pinned - ids to reuse rather than allocate. Auto-paging re-registers the same bullet
 *   on each overflow slide while sharing one options object by reference, so the pair has to
 *   keep the ids that object already carries.
 * @returns the two allocated relationship ids
 */
export function registerSvgImageRels(
	target: PresSlideInternal,
	source: { path: string; data: string; svgSize?: { w: number; h: number } },
	pinned?: { pngRid: number; svgRid: number }
): { pngRid: number; svgRid: number } {
	const { path, data } = source

	// The PNG fallback's name is read before the SVG's push, so the pair lands on consecutive names.
	const pngRid = pinned ? pinned.pngRid : getNewRelId(target)
	pushMediaRel(target, {
		kind: 'image',
		extn: 'png',
		type: 'image/png',
		path,
		data,
		rId: pngRid,
		extra: { isSvgPng: true, ...(source.svgSize ? { svgSize: source.svgSize } : {}) },
	})

	const svgRid = pinned ? pinned.svgRid : getNewRelId(target)
	pushMediaRel(target, { kind: 'image', extn: 'svg', type: 'image/svg+xml', path, data, rId: svgRid })

	return { pngRid, svgRid }
}
