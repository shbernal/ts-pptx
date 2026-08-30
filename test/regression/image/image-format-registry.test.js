// The image-format registry is internally consistent in all three directions.
//
// It replaced four tables that answered overlapping questions about the same nine formats, and
// two of the four had already drifted (`image/jpeg` → `jpeg` in one, `jpg` in another;
// `image/tiff` → `tiff` and `tif`). One table cannot drift against itself, but it can still be
// written down wrong — a row whose content type does not map back to its own extension, or whose
// signature is another row's — so each direction is asserted against the rows themselves rather
// than against a second copy of the answers.
//
// Imported from `src/` rather than `dist/`: the registry is internal plumbing with no entry-point
// export, and the behaviour it drives is covered through the package by the suites around it.

import { describe, test } from 'vitest'
import {
	IMAGE_FORMATS,
	imageFormatForBytes,
	imageFormatForContentType,
	imageFormatForExtension,
} from '../../../src/media/image-formats.ts'
import { assetFilenameExtension, imageContentType } from '../../../src/media/content-type.ts'
import { assert, assertEqual } from '../../helpers.js'

/** Leading bytes that satisfy each row's signature, built independently of the matchers. */
const HEADERS = {
	png: [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a],
	jpeg: [0xff, 0xd8, 0xff, 0xe0],
	gif: [0x47, 0x49, 0x46, 0x38, 0x39, 0x61],
	bmp: [0x42, 0x4d, 0x00, 0x00],
	tiff: [0x49, 0x49, 0x2a, 0x00],
	webp: [0x52, 0x49, 0x46, 0x46, 0, 0, 0, 0, 0x57, 0x45, 0x42, 0x50],
}

const rows = Object.entries(IMAGE_FORMATS)

describe('image format registry', () => {
	test('every row is keyed by its own canonical extension', () => {
		for (const [key, format] of rows) assertEqual(format.ext, key, `${key} is keyed by its ext`)
	})

	test('every row round-trips ext → contentType → ext, alternate spellings included', () => {
		for (const [key, format] of rows) {
			assertEqual(imageContentType(format.ext), format.contentType, `${key}: ext → contentType`)
			assertEqual(imageFormatForContentType(format.contentType)?.ext, format.ext, `${key}: contentType → ext`)
			for (const alt of format.altExts) {
				assertEqual(imageContentType(alt), format.contentType, `${key}: ${alt} → contentType`)
				assertEqual(imageFormatForExtension(alt)?.ext, format.ext, `${key}: ${alt} resolves to the row`)
			}
		}
	})

	test('extensions and content types are each claimed by exactly one row', () => {
		const exts = rows.flatMap(([, f]) => [f.ext, ...f.altExts])
		assertEqual(new Set(exts).size, exts.length, `no extension is claimed twice (${exts.join(', ')})`)
		const types = rows.map(([, f]) => f.contentType)
		assertEqual(new Set(types).size, types.length, `no content type is claimed twice (${types.join(', ')})`)
	})

	test("each row's header sniffs back to that same row, and to no other", () => {
		for (const [key, header] of Object.entries(HEADERS)) {
			const bytes = new Uint8Array(header)
			assertEqual(imageFormatForBytes(bytes)?.ext, key, `${key} header sniffs back to ${key}`)
			for (const [other, format] of rows) {
				if (other === key) continue
				assert(!format.magic?.(bytes), `${key}'s header does not also match ${other}`)
			}
		}
	})

	test('every row with a signature has a header case above it', () => {
		for (const [key, format] of rows) {
			assertEqual(
				format.magic !== null,
				key in HEADERS,
				`${key} has a signature iff this suite exercises one — add a header when a row gains a matcher`
			)
		}
	})

	test('bytes matching nothing sniff to null', () => {
		assertEqual(imageFormatForBytes(new Uint8Array([0x00, 0x01, 0x02, 0x03])), null, 'unknown bytes')
		assertEqual(imageFormatForBytes(new Uint8Array()), null, 'empty bytes')
	})

	test('the on-disk asset spelling is the filenameExt column, not the part spelling', () => {
		for (const [key, format] of rows) {
			assertEqual(assetFilenameExtension(format.contentType), format.filenameExt, `${key}: asset filename`)
		}
		// The two rows where the two columns deliberately differ, spelled out so a change to either
		// has to be made on purpose.
		assertEqual(IMAGE_FORMATS.jpeg.ext, 'jpeg', 'jpeg media parts are .jpeg')
		assertEqual(IMAGE_FORMATS.jpeg.filenameExt, 'jpg', 'jpeg assets on disk are .jpg')
		assertEqual(IMAGE_FORMATS.tiff.ext, 'tiff', 'tiff media parts are .tiff')
		assertEqual(IMAGE_FORMATS.tiff.filenameExt, 'tif', 'tiff assets on disk are .tif')
	})

	test('content types PowerPoint authors are kept over the IANA-preferred spellings', () => {
		assertEqual(imageContentType('emf'), 'image/x-emf', 'emf')
		assertEqual(imageContentType('wmf'), 'image/x-wmf', 'wmf')
		assertEqual(imageContentType('svg'), 'image/svg+xml', 'svg')
	})

	test('an extension the registry does not carry still gets image/<extn>', () => {
		assertEqual(imageContentType('avif'), 'image/avif', 'unlisted extension')
		assertEqual(imageContentType('PNG'), 'image/png', 'case-insensitive')
		assertEqual(imageFormatForExtension('avif'), null, 'and resolves to no row')
	})

	test('audio content types resolve to their own filename spelling', () => {
		assertEqual(assetFilenameExtension('audio/x-wav'), 'wav', 'audio/x-wav')
		assertEqual(assetFilenameExtension('audio/wav'), 'wav', 'audio/wav')
		assertEqual(assetFilenameExtension('audio/mpeg'), 'mp3', 'audio/mpeg')
		assertEqual(assetFilenameExtension('video/mp4'), null, 'video is left to the source partname')
	})
})
