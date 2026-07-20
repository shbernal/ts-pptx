/**
 * OPC content types for embedded media parts, keyed by file extension.
 *
 * Both mappings reproduce what PowerPoint itself authors rather than the IANA-preferred
 * spelling — a package that disagrees (e.g. `audio/mp3` for an `.mp3`) is the kind of thing
 * PowerPoint offers to repair.
 */

/**
 * Map an image file extension to its OOXML content type.
 * Inverse of the read-side `IMAGE_EXTENSION_BY_CONTENT_TYPE` (src/read/api/shapes.ts):
 * EMF/WMF use the `x-`-prefixed forms PowerPoint authors (and that the read side
 * expects), `jpg`/`jpeg` normalize to `image/jpeg`, and `svg` to `image/svg+xml`.
 * Only the content type is derived here; the file extension (used for the media
 * Target filename) is left to the caller.
 * @param {string} extn - image file extension (e.g. `png`, `jpg`, `emf`)
 * @returns {string} OOXML content type (e.g. `image/png`, `image/x-emf`)
 */
export function imageContentType(extn: string): string {
	switch ((extn || '').toLowerCase()) {
		case 'emf':
			return 'image/x-emf'
		case 'wmf':
			return 'image/x-wmf'
		case 'svg':
			return 'image/svg+xml'
		case 'jpg':
		case 'jpeg':
			return 'image/jpeg'
		default:
			return 'image/' + (extn || '').toLowerCase()
	}
}

/**
 * Resolve the OPC content type for an embedded audio/video part by file extension,
 * matching what PowerPoint authors (e.g. `mp3` → `audio/mpeg`, not `audio/mp3`).
 * The `mtype` disambiguates extensions Office maps differently per kind and seeds
 * the `mtype/extn` fallback for anything unlisted.
 * @param {string} extn - media file extension (no dot), case-insensitive
 * @param {'audio' | 'video'} mtype - whether the item is audio or video
 */
export function avContentType(extn: string, mtype: 'audio' | 'video'): string {
	switch ((extn || '').toLowerCase()) {
		// video
		case 'mp4':
			return mtype === 'audio' ? 'audio/mp4' : 'video/mp4'
		case 'm4v':
			return 'video/mp4'
		case 'mov':
			return 'video/quicktime'
		case 'avi':
			return 'video/avi'
		case 'wmv':
			return 'video/x-ms-wmv'
		case 'mpg':
		case 'mpeg':
			return mtype === 'audio' ? 'audio/mpeg' : 'video/mpeg'
		case 'ogv':
			return 'video/ogg'
		case 'webm':
			return 'video/webm'
		// audio
		case 'mp3':
			return 'audio/mpeg'
		case 'm4a':
			return 'audio/mp4'
		case 'wav':
			return 'audio/x-wav' // PowerPoint authors the x- form (e.g. embedded transition sounds)
		case 'wma':
			return 'audio/x-ms-wma'
		case 'aac':
			return 'audio/aac'
		case 'oga':
		case 'ogg':
			return 'audio/ogg'
		case 'flac':
			return 'audio/flac'
		default:
			return mtype + '/' + (extn || '').toLowerCase()
	}
}
