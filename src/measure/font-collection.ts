/**
 * Font **collection** (`.ttc` / `.otc`) support for the write-side metrics.
 *
 * A collection is a single file holding several fonts that share their heavy tables.
 * `msgothic.ttc` is three fonts (MS Gothic, MS UI Gothic, MS PGothic) over one 5.5 MB
 * `glyf`; `cambria.ttc` is Cambria plus Cambria Math over one outline set. On Windows
 * that is how most of the East Asian faces ship, and Cambria too, so a measurement API
 * that only reads a bare sfnt cannot measure them at all.
 *
 * `opentype.js` does not read the `ttcf` wrapper (it throws `Unsupported OpenType
 * signature ttcf` from `parseBuffer`, still true in 2.0.0), so this module unwraps one
 * member into a standalone sfnt before handing the bytes over. It parses only the two
 * structures that requires: the table directories and the `name` table.
 *
 * ## Layout
 *
 * ```text
 * 0    'ttcf'  uint16 major, uint16 minor, uint32 numFonts, uint32 dirOffset[numFonts]
 * ...  per member: uint32 sfntVersion, uint16 numTables, uint16 searchRange,
 *      uint16 entrySelector, uint16 rangeShift, then numTables x 16-byte records
 *      { char tag[4], uint32 checksum, uint32 offset, uint32 length }
 * ```
 *
 * The record `offset` is **absolute from the start of the file**, which is the whole
 * trick: two members name the same `glyf` bytes by pointing at them. So a member does
 * not have to be *copied out* to be read on its own; it is enough to give the reader a
 * buffer whose first bytes are that member's directory, leaving every table where it
 * already sits. See {@link extractFontFace}.
 */

import { InvalidOptionError, MediaError } from '../errors.js'

/** `'ttcf'`, the tag that marks a font collection. */
const TTC_TAG = 'ttcf'

/** Bytes in one sfnt table record: tag, checksum, offset, length. */
const RECORD_SIZE = 16

/** Bytes in an sfnt table directory header, before the first record. */
const DIRECTORY_HEADER_SIZE = 12

/** `name` IDs read for face identity: family, subfamily, full name, PostScript name. */
const NAME_FAMILY = 1
const NAME_SUBFAMILY = 2
const NAME_FULL = 4
const NAME_POSTSCRIPT = 6

/** Identity of one font inside a font file, as read from its `name` table. */
export interface FontFaceInfo {
	/** Position in the file: `0` for a plain TTF/OTF, `0..n-1` inside a collection. */
	readonly index: number
	/** `name` ID 1, e.g. `'MS PGothic'`. Absent when the table has no readable record. */
	readonly family?: string
	/** `name` ID 2, e.g. `'Regular'`. */
	readonly subfamily?: string
	/** `name` ID 4, e.g. `'MS PGothic'`. */
	readonly fullName?: string
	/** `name` ID 6, e.g. `'MS-PGothic'`. */
	readonly postScriptName?: string
}

/** A parsed sfnt table directory: the member's format tag plus its table records. */
interface Directory {
	readonly sfntVersion: number
	readonly tables: readonly TableRecord[]
}

interface TableRecord {
	/** Four ASCII characters, e.g. `'glyf'`. */
	readonly tag: string
	readonly checksum: number
	/** Absolute from the start of the file, shared between members of a collection. */
	readonly offset: number
	readonly length: number
}

function view(data: Uint8Array): DataView {
	return new DataView(data.buffer, data.byteOffset, data.byteLength)
}

/** Four ASCII bytes at `off` as a string, or `''` when they run past the end. */
function tagAt(data: Uint8Array, off: number): string {
	if (off < 0 || off + 4 > data.byteLength) return ''
	let s = ''
	for (let i = 0; i < 4; i++) s += String.fromCharCode(data[off + i] as number)
	return s
}

/** True when `data` starts with the `ttcf` tag, i.e. holds more than one font. */
export function isFontCollection(data: Uint8Array): boolean {
	return tagAt(data, 0) === TTC_TAG
}

/**
 * Offsets of every member's table directory, in file order. A plain sfnt has exactly
 * one, at 0.
 */
function directoryOffsets(data: Uint8Array): number[] {
	if (!isFontCollection(data)) return [0]
	const dv = view(data)
	if (data.byteLength < 12) throw new MediaError('font/parse-failed', 'Font collection header is truncated')
	const numFonts = dv.getUint32(8)
	// A collection with no members, or one whose offset table cannot fit, is malformed
	// rather than merely unsupported: there is nothing to select from.
	if (numFonts === 0) throw new MediaError('font/parse-failed', 'Font collection declares 0 fonts')
	if (12 + numFonts * 4 > data.byteLength)
		throw new MediaError(
			'font/parse-failed',
			`Font collection declares ${numFonts} fonts but the offset table is truncated`
		)
	const offsets: number[] = []
	for (let i = 0; i < numFonts; i++) offsets.push(dv.getUint32(12 + i * 4))
	return offsets
}

/** Read the table directory at `dirOff`, materializing every record before any write. */
function readDirectory(data: Uint8Array, dirOff: number): Directory {
	const dv = view(data)
	if (dirOff + DIRECTORY_HEADER_SIZE > data.byteLength)
		throw new MediaError('font/parse-failed', `Font table directory at ${dirOff} is truncated`)
	const sfntVersion = dv.getUint32(dirOff)
	const numTables = dv.getUint16(dirOff + 4)
	if (numTables === 0) throw new MediaError('font/parse-failed', `Font at ${dirOff} declares 0 tables`)
	const end = dirOff + DIRECTORY_HEADER_SIZE + numTables * RECORD_SIZE
	if (end > data.byteLength)
		throw new MediaError(
			'font/parse-failed',
			`Font at ${dirOff} declares ${numTables} tables but the directory is truncated`
		)
	const tables: TableRecord[] = []
	for (let i = 0; i < numTables; i++) {
		const rec = dirOff + DIRECTORY_HEADER_SIZE + i * RECORD_SIZE
		tables.push({
			tag: tagAt(data, rec),
			checksum: dv.getUint32(rec + 4),
			offset: dv.getUint32(rec + 8),
			length: dv.getUint32(rec + 12),
		})
	}
	return { sfntVersion, tables }
}

/**
 * Decode one `name` record's string. Windows (platform 3) and Unicode (platform 0)
 * records are UTF-16BE; Macintosh (platform 1) Roman is byte-per-character and
 * coincides with Latin-1 over the ASCII range these names use.
 */
function decodeNameString(data: Uint8Array, off: number, length: number, platformID: number): string {
	if (off < 0 || off + length > data.byteLength) return ''
	if (platformID === 1) {
		let s = ''
		for (let i = 0; i < length; i++) s += String.fromCharCode(data[off + i] as number)
		return s
	}
	let s = ''
	for (let i = 0; i + 1 < length; i += 2)
		s += String.fromCharCode(((data[off + i] as number) << 8) | (data[off + i + 1] as number))
	return s
}

/**
 * How much a `name` record is preferred when several carry the same ID. Windows English
 * wins, then any Windows or Unicode record, then Macintosh. The ranking matters because
 * `msgothic.ttc` carries both `'MS Gothic'` (0x409) and the Japanese localized name for
 * the same ID, and a caller selecting by name types the former.
 */
function nameRecordRank(platformID: number, languageID: number): number {
	if (platformID === 3) return languageID === 0x409 ? 3 : 2
	if (platformID === 0) return 2
	if (platformID === 1) return languageID === 0 ? 1 : 0
	return 0
}

/** Read the wanted IDs out of a `name` table, keeping the best-ranked record for each. */
function readNames(data: Uint8Array, base: number, length: number): Map<number, string> {
	const out = new Map<number, string>()
	const ranks = new Map<number, number>()
	if (base + 6 > data.byteLength || base + length > data.byteLength) return out
	const dv = view(data)
	const count = dv.getUint16(base + 2)
	const storage = base + dv.getUint16(base + 4)
	for (let i = 0; i < count; i++) {
		const rec = base + 6 + i * 12
		if (rec + 12 > data.byteLength) break
		const nameID = dv.getUint16(rec + 6)
		if (nameID !== NAME_FAMILY && nameID !== NAME_SUBFAMILY && nameID !== NAME_FULL && nameID !== NAME_POSTSCRIPT)
			continue
		const platformID = dv.getUint16(rec)
		const languageID = dv.getUint16(rec + 4)
		const rank = nameRecordRank(platformID, languageID)
		if (rank <= (ranks.get(nameID) ?? -1)) continue
		const text = decodeNameString(data, storage + dv.getUint16(rec + 10), dv.getUint16(rec + 8), platformID)
		if (!text) continue
		out.set(nameID, text)
		ranks.set(nameID, rank)
	}
	return out
}

/**
 * The fonts inside `data`, in file order. A plain TTF/OTF yields exactly one entry, so a
 * caller never has to branch on the container format.
 *
 * Identity comes from each member's own `name` table, which is one of the tables a
 * collection does *not* share. A field is absent when the table carries no readable
 * record for it; a font whose `name` table is missing or malformed still yields an entry
 * (with only `index`), because the metrics do not depend on it.
 * @param {Uint8Array} data - the font file's bytes
 * @returns {FontFaceInfo[]} one entry per font, index-ordered
 * @example listFontFaces(readFileSync('C:/Windows/Fonts/msgothic.ttc')).map((f) => f.family)
 */
export function listFontFaces(data: Uint8Array): FontFaceInfo[] {
	return directoryOffsets(data).map((dirOff, index) => {
		const { tables } = readDirectory(data, dirOff)
		const name = tables.find((t) => t.tag === 'name')
		const names = name ? readNames(data, name.offset, name.length) : new Map<number, string>()
		const family = names.get(NAME_FAMILY)
		const subfamily = names.get(NAME_SUBFAMILY)
		const fullName = names.get(NAME_FULL)
		const postScriptName = names.get(NAME_POSTSCRIPT)
		return {
			index,
			...(family ? { family } : {}),
			...(subfamily ? { subfamily } : {}),
			...(fullName ? { fullName } : {}),
			...(postScriptName ? { postScriptName } : {}),
		}
	})
}

/** A face's names, lowercased, for case-insensitive selection. */
function selectableNames(face: FontFaceInfo): string[] {
	return [face.family, face.fullName, face.postScriptName]
		.filter((n): n is string => typeof n === 'string' && n.length > 0)
		.map((n) => n.toLowerCase())
}

/** Render the available faces for an error message, e.g. ``0 `MS Gothic`, 1 `MS PGothic` ``. */
function describeFaces(faces: readonly FontFaceInfo[]): string {
	return faces.map((f) => `${f.index} \`${f.family ?? f.fullName ?? '(unnamed)'}\``).join(', ')
}

/**
 * Resolve a `font` selector against the faces in `data`.
 *
 * The selector means the same thing whatever the container: a plain TTF/OTF is a
 * one-entry list, so `{ font: 1 }` or a name that does not match is an error there too
 * rather than being quietly ignored. Silently falling back to the only/first font is the
 * failure mode worth refusing: it measures one face's advances while the caller believes
 * it registered another's, and nothing downstream can tell.
 * @param {Uint8Array} data - the font file's bytes
 * @param {number | string} [font] - 0-based index, or a name matched case-insensitively
 *   against the family, full, and PostScript names. Absent selects index 0.
 * @returns {number} the index of the selected font
 */
export function resolveFontFace(data: Uint8Array, font?: number | string): number {
	if (font === undefined) return 0
	const faces = listFontFaces(data)
	if (typeof font === 'number') {
		if (!Number.isInteger(font) || font < 0 || font >= faces.length)
			throw new InvalidOptionError(
				'font/collection-index-out-of-range',
				`Font index ${font} is out of range: the file holds ${faces.length} font(s) (${describeFaces(faces)})`
			)
		return font
	}
	const wanted = font.toLowerCase()
	const hit = faces.find((f) => selectableNames(f).includes(wanted))
	if (!hit)
		throw new InvalidOptionError(
			'font/collection-face-not-found',
			`No font named "${font}" in this file. It holds: ${describeFaces(faces)}`
		)
	return hit.index
}

/**
 * Bytes for one font of `data` as a standalone sfnt, ready for a parser that does not
 * read the `ttcf` wrapper.
 *
 * A plain TTF/OTF at index 0 is returned as-is. For a collection the member's tables are
 * **not copied out**: the file is copied once and the member's directory is written over
 * the prologue at offset 0, so every record keeps the absolute offset it already had and
 * the shared `glyf` is neither duplicated nor moved. The prologue is always large enough
 * to hold that directory, because a well-formed collection places the first table after
 * the `ttcf` header *and* every member's directory, which already includes this one.
 * That is asserted rather than assumed: a file that violates it would otherwise have its
 * first table silently overwritten.
 * @param {Uint8Array} data - the font file's bytes
 * @param {number} index - which font, as resolved by {@link resolveFontFace}
 * @returns {Uint8Array} a buffer whose first bytes are a plain sfnt table directory
 */
export function extractFontFace(data: Uint8Array, index: number): Uint8Array {
	const offsets = directoryOffsets(data)
	const dirOff = offsets[index]
	if (dirOff === undefined)
		throw new InvalidOptionError(
			'font/collection-index-out-of-range',
			`Font index ${index} is out of range: the file holds ${offsets.length} font(s)`
		)
	if (!isFontCollection(data)) return data

	const { sfntVersion, tables } = readDirectory(data, dirOff)
	const prologue = DIRECTORY_HEADER_SIZE + tables.length * RECORD_SIZE
	const firstTable = Math.min(...tables.map((t) => t.offset))
	if (prologue > firstTable)
		throw new MediaError(
			'font/parse-failed',
			`Font ${index} in this collection cannot be unwrapped: its ${tables.length}-table directory needs ${prologue} bytes but the first table starts at ${firstTable}`
		)

	// Copy first, then overwrite: `dirOff` frequently sits *inside* the region about to be
	// written (member 0 of msgothic.ttc has its directory at 36, and the prologue is 380),
	// which is safe only because `readDirectory` already materialized every record.
	const out = new Uint8Array(data)
	const dv = new DataView(out.buffer, out.byteOffset, out.byteLength)
	const entrySelector = Math.floor(Math.log2(tables.length))
	const searchRange = RECORD_SIZE * 2 ** entrySelector
	dv.setUint32(0, sfntVersion)
	dv.setUint16(4, tables.length)
	dv.setUint16(6, searchRange)
	dv.setUint16(8, entrySelector)
	dv.setUint16(10, tables.length * RECORD_SIZE - searchRange)
	let p = DIRECTORY_HEADER_SIZE
	for (const t of tables) {
		for (let i = 0; i < 4; i++) out[p + i] = t.tag.charCodeAt(i) & 0xff
		dv.setUint32(p + 4, t.checksum)
		dv.setUint32(p + 8, t.offset)
		dv.setUint32(p + 12, t.length)
		p += RECORD_SIZE
	}
	// `head.checkSumAdjustment` now describes the collection rather than this buffer. No
	// consumer here validates it, and recomputing it would mean check-summing the whole
	// file to produce a number only a validator reads.
	return out
}
