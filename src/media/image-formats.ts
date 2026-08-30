/**
 * The image formats this library recognises, one row each.
 *
 * Four tables used to answer overlapping questions about the same nine formats in three
 * directions — extension → content type on the write path, content type → extension on the read
 * and script paths, magic bytes → both in the sniffer — and a fifth partial re-derived the same
 * signatures to read pixel dimensions. They had already drifted: `image/jpeg` resolved to `jpeg`
 * in one and `jpg` in another, `image/tiff` to `tiff` and `tif`, and the GIF signature was four
 * bytes in one place and three in another, so a file could be measured but not sniffed.
 *
 * A row states every spelling once, so the directions cannot disagree. Where two spellings are
 * both wanted they are separate *columns* with separate reasons ({@link ImageFormat.ext} vs
 * {@link ImageFormat.filenameExt}), never separate tables.
 *
 * Content types are what PowerPoint authors, not what IANA prefers: EMF and WMF keep the `x-`
 * prefix and SVG is `image/svg+xml`. A package that disagrees is the kind of thing PowerPoint
 * offers to repair.
 */

/** One image format, in every spelling the library needs it in. */
export interface ImageFormat {
	/**
	 * The extension a *media part* carries (`/ppt/media/image1.<ext>`), and the canonical answer
	 * when a content type is turned back into an extension. Changing one of these changes emitted
	 * package bytes.
	 */
	readonly ext: string
	/** The OPC content type, as PowerPoint authors it. */
	readonly contentType: string
	/** Other extensions a caller may spell the format with; each resolves to {@link contentType}. */
	readonly altExts: readonly string[]
	/**
	 * The extension for a *file on disk* — the asset a printed script writes beside itself. It
	 * differs from {@link ext} for two formats (`jpg`, `tif`) because those are the spellings
	 * desktop viewers and file managers recognise, and an asset directory is read by people. It
	 * names no part and is in no package, so it is free to differ.
	 */
	readonly filenameExt: string
	/**
	 * Whether these leading bytes are this format's signature, or `null` for a format with no
	 * usable one (SVG is text; EMF and WMF are not sniffed).
	 */
	readonly magic: ((b: Uint8Array) => boolean) | null
}

/**
 * Every recognised format, keyed by its canonical extension. Iteration order is sniffing order;
 * no two signatures overlap, so it only has to be stable, not carefully chosen.
 */
export const IMAGE_FORMATS = Object.freeze({
	png: {
		ext: 'png',
		contentType: 'image/png',
		altExts: [],
		filenameExt: 'png',
		magic: (b) => b.length >= 8 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47,
	},
	jpeg: {
		ext: 'jpeg',
		contentType: 'image/jpeg',
		altExts: ['jpg'],
		filenameExt: 'jpg',
		magic: (b) => b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff,
	},
	gif: {
		ext: 'gif',
		contentType: 'image/gif',
		altExts: [],
		filenameExt: 'gif',
		// `GIF8`, not `GIF`: both released versions are GIF87a and GIF89a, and the fourth byte is
		// what tells a header apart from a file that merely starts with those three letters.
		magic: (b) => b.length >= 4 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38,
	},
	bmp: {
		ext: 'bmp',
		contentType: 'image/bmp',
		altExts: [],
		filenameExt: 'bmp',
		magic: (b) => b.length >= 2 && b[0] === 0x42 && b[1] === 0x4d,
	},
	tiff: {
		ext: 'tiff',
		contentType: 'image/tiff',
		altExts: ['tif'],
		filenameExt: 'tif',
		// Both byte orders: `II*\0` little-endian, `MM\0*` big-endian.
		magic: (b) =>
			b.length >= 4 &&
			((b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a && b[3] === 0x00) ||
				(b[0] === 0x4d && b[1] === 0x4d && b[2] === 0x00 && b[3] === 0x2a)),
	},
	webp: {
		ext: 'webp',
		contentType: 'image/webp',
		altExts: [],
		filenameExt: 'webp',
		// `RIFF` then `WEBP` at offset 8; the chunk fourCC after it is the codec, not the format.
		magic: (b) =>
			b.length >= 12 &&
			b[0] === 0x52 &&
			b[1] === 0x49 &&
			b[2] === 0x46 &&
			b[3] === 0x46 &&
			b[8] === 0x57 &&
			b[9] === 0x45 &&
			b[10] === 0x42 &&
			b[11] === 0x50,
	},
	svg: { ext: 'svg', contentType: 'image/svg+xml', altExts: [], filenameExt: 'svg', magic: null },
	emf: { ext: 'emf', contentType: 'image/x-emf', altExts: [], filenameExt: 'emf', magic: null },
	wmf: { ext: 'wmf', contentType: 'image/x-wmf', altExts: [], filenameExt: 'wmf', magic: null },
}) satisfies Readonly<Record<string, ImageFormat>>

const BY_EXTENSION: ReadonlyMap<string, ImageFormat> = new Map(
	Object.values(IMAGE_FORMATS).flatMap((f) => [f.ext, ...f.altExts].map((e) => [e, f] as const))
)

const BY_CONTENT_TYPE: ReadonlyMap<string, ImageFormat> = new Map(
	Object.values(IMAGE_FORMATS).map((f) => [f.contentType, f] as const)
)

/** The format a caller's file extension names, or `null` for one the registry does not carry. */
export function imageFormatForExtension(extn: string): ImageFormat | null {
	return BY_EXTENSION.get((extn || '').toLowerCase()) ?? null
}

/** The format a content type names, or `null` for one the registry does not carry. */
export function imageFormatForContentType(contentType: string): ImageFormat | null {
	return BY_CONTENT_TYPE.get((contentType || '').toLowerCase()) ?? null
}

/** The format these leading bytes are the header of, or `null` when nothing matches. */
export function imageFormatForBytes(bytes: Uint8Array): ImageFormat | null {
	if (!bytes) return null
	for (const format of Object.values(IMAGE_FORMATS)) {
		if (format.magic?.(bytes)) return format
	}
	return null
}
