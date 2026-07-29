/**
 * OPC content types for embedded media parts, keyed by file extension.
 *
 * Both mappings reproduce what PowerPoint itself authors rather than the IANA-preferred
 * spelling — a package that disagrees (e.g. `audio/mp3` for an `.mp3`) is the kind of thing
 * PowerPoint offers to repair.
 */

/**
 * Map an image file extension to its OOXML content type.
 * Inverse of the read-side `IMAGE_EXTENSION_BY_CONTENT_TYPE` (src/read/api/shapes/picture.ts):
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
 * Resolve the file extension an embedded image part should carry, from the caller's source.
 *
 * A `data:` URI states its mime type outright, so it wins over `path`: the bytes are the thing
 * being embedded, and a caller may supply them with no path at all (or with a path whose
 * extension disagrees with the payload). The extension chosen here names the media part *and*,
 * via {@link imageContentType}, decides the `<Default>` the package declares for it — so
 * getting it from a placeholder path is how a deck ends up declaring `image/png` over SVG
 * bytes, which is exactly the mismatch PowerPoint offers to "repair".
 *
 * `image/svg+xml` needs its own test because the `\w+` mime capture stops at the `+`. The
 * result is always lower-cased, so a `data:image/PNG;` source names its part the same way a
 * `photo.PNG` path does.
 * @param {string} path - caller-supplied path/URL (may be empty); query and fragment are stripped
 * @param {string} data - caller-supplied `data:`/base64 payload (may be empty)
 * @returns {string} file extension, no dot (e.g. `png`, `jpeg`, `svg`)
 */
export function imageExtensionForSource(path: string, data: string): string {
	// NOTE: Split to address URLs with params (eg: `path/brent.jpg?someParam=true`)
	const strPath = path || ''
	const strData = data || ''
	const pathFile = strPath.slice(strPath.lastIndexOf('/') + 1).split('?')[0] || ''
	const pathExtn = ((pathFile.split('.').pop() || 'png').split('#')[0] || 'png').toLowerCase()

	// Pre-encoded images can be whatever mime-type they want (and good for them!)
	const mimeMatch = /image\/(\w+);/.exec(strData)
	if (strData && mimeMatch?.[1]) return mimeMatch[1].toLowerCase()
	if (strData.toLowerCase().includes('image/svg+xml')) return 'svg'
	return pathExtn
}

/**
 * Map an audio content type's subtype back to the file extension a media part should use.
 *
 * Needed because a caller supplying sound bytes as a `data:` URI states a *content type*
 * (`data:audio/x-wav;base64,…`), while the media part needs a *filename*. Taking the subtype
 * verbatim produces `audio1.x-wav` — legal, since the package would then declare a matching
 * `<Default Extension="x-wav"/>`, but a file type that does not exist and that no other tool
 * recognises. `audio/x-wav` in particular is not an edge case: it is exactly what PowerPoint
 * authors for an embedded transition sound, so it arrives on any deck read back in.
 *
 * Only the spellings that are *not* already an extension are listed; anything else (`ogg`,
 * `flac`, `aac`, …) is its own extension and falls through. Read against {@link avContentType},
 * which maps the same pairs in the other direction.
 */
export function audioExtensionForSubtype(subtype: string): string {
	switch ((subtype || '').toLowerCase()) {
		case 'x-wav':
		case 'wave':
		case 'vnd.wave':
			return 'wav'
		case 'mpeg':
			return 'mp3'
		case 'x-ms-wma':
			return 'wma'
		case 'mp4':
			return 'm4a'
		default:
			return (subtype || '').toLowerCase()
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
