// Font collections (`.ttc`/`.otc`) through the public `ts-pptx/measure` surface.
//
// A collection is one file holding several fonts over shared tables, and it is how most
// of the East Asian faces ship on Windows (MS Gothic, Yu Gothic, SimSun, Microsoft YaHei)
// plus Cambria. opentype.js does not read the `ttcf` wrapper - `parseBuffer` throws
// `Unsupported OpenType signature ttcf`, still true in 2.0.0 - so src/measure/font-collection.ts
// unwraps the selected member into a standalone sfnt first.
//
// Two suites, because the two claims have different reachability:
//
//   1. A collection SYNTHESIZED here from the repo's own Silkscreen .ttf files. This runs
//      everywhere including CI, and it is the assertion with teeth: the metrics read back
//      out of a member must equal, exactly, the metrics of that same font parsed as a
//      plain .ttf. The plain path hands opentype.js the untouched file, so equality means
//      the directory rewrite preserved everything the parser reads. `regular != bold` is
//      asserted alongside it so the equality check is known to be able to fail.
//   2. The GENUINE collections in the Windows font directory, checked against
//      `windows-collections.oracle.json` - advances read by WPF's GlyphTypeface, an
//      implementation that shares no code with this repo (see the authoring script). This
//      is what rules out the failure the synthetic case cannot see: a builder and a reader
//      agreeing on a format that real files do not follow. Skips off Windows.
import { readFileSync, existsSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, test, expect } from 'vitest'
import TsPptx from '../../../dist/node.js'
import { parseFontMetrics, listFontFaces, isFontCollection } from '../../../dist/measure.js'
import { InvalidOptionError, MediaError } from '../../../dist/measure.js'

const fixture = (name) => fileURLToPath(new URL(`../../read/fixtures/fonts/${name}`, import.meta.url))
const REG_BYTES = new Uint8Array(readFileSync(fixture('Silkscreen-Regular.ttf')))
const BOLD_BYTES = new Uint8Array(readFileSync(fixture('Silkscreen-Bold.ttf')))

/** Characters Silkscreen covers, wide enough that a wrong table would show up. */
const SWEEP = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,;:!?()[]{}/@#$%&*-+='

/** Every advance in `SWEEP`, one per character, at a size that keeps design units visible. */
function advanceProfile(metrics) {
	return [...SWEEP].map((ch) => metrics.advanceWidthPt(ch, 1000))
}

// --- a faithful .ttc builder -------------------------------------------------------
//
// Mirrors the layout the real files use: the `ttcf` header, then every member's table
// directory, then the table data, with each record carrying an offset ABSOLUTE to the
// start of the file. Identical tables are stored once and named by both members, which is
// the whole point of the container and what makes those offsets shared rather than
// merely absolute.

const RECORD_SIZE = 16

/** Read a plain sfnt's table directory into records. */
function readTables(data) {
	const dv = new DataView(data.buffer, data.byteOffset, data.byteLength)
	const numTables = dv.getUint16(4)
	const tables = []
	for (let i = 0; i < numTables; i++) {
		const rec = 12 + i * RECORD_SIZE
		tables.push({
			tag: String.fromCharCode(...data.subarray(rec, rec + 4)),
			checksum: dv.getUint32(rec + 4),
			offset: dv.getUint32(rec + 8),
			length: dv.getUint32(rec + 12),
		})
	}
	return { sfntVersion: dv.getUint32(0), tables }
}

const align4 = (n) => (n + 3) & ~3

/**
 * Pack plain sfnt fonts into one `ttcf` collection, in the order given.
 * @param {Uint8Array[]} fonts
 * @returns {Uint8Array}
 */
function buildCollection(fonts) {
	const parsed = fonts.map(readTables)
	let cursor = align4(12 + fonts.length * 4 + parsed.reduce((n, p) => n + 12 + p.tables.length * RECORD_SIZE, 0))

	// Place table data once per distinct (tag, bytes), so a table two members share is
	// stored once and pointed at twice.
	const placed = new Map()
	const blocks = []
	const placements = parsed.map((p, f) =>
		p.tables.map((t) => {
			const bytes = fonts[f].subarray(t.offset, t.offset + t.length)
			const key = `${t.tag}\u0000${Buffer.from(bytes).toString('base64')}`
			let offset = placed.get(key)
			if (offset === undefined) {
				offset = cursor
				placed.set(key, offset)
				blocks.push({ offset, bytes })
				cursor = align4(cursor + t.length)
			}
			return { ...t, offset }
		})
	)

	const out = new Uint8Array(cursor)
	const dv = new DataView(out.buffer)
	out.set([0x74, 0x74, 0x63, 0x66], 0) // 'ttcf'
	dv.setUint16(4, 2) // major version
	dv.setUint16(6, 0) // minor version
	dv.setUint32(8, fonts.length)

	let dirCursor = 12 + fonts.length * 4
	placements.forEach((tables, f) => {
		dv.setUint32(12 + f * 4, dirCursor)
		const entrySelector = Math.floor(Math.log2(tables.length))
		const searchRange = RECORD_SIZE * 2 ** entrySelector
		dv.setUint32(dirCursor, parsed[f].sfntVersion)
		dv.setUint16(dirCursor + 4, tables.length)
		dv.setUint16(dirCursor + 6, searchRange)
		dv.setUint16(dirCursor + 8, entrySelector)
		dv.setUint16(dirCursor + 10, tables.length * RECORD_SIZE - searchRange)
		let rec = dirCursor + 12
		for (const t of tables) {
			for (let i = 0; i < 4; i++) out[rec + i] = t.tag.charCodeAt(i)
			dv.setUint32(rec + 4, t.checksum)
			dv.setUint32(rec + 8, t.offset)
			dv.setUint32(rec + 12, t.length)
			rec += RECORD_SIZE
		}
		dirCursor = rec
	})
	for (const b of blocks) out.set(b.bytes, b.offset)
	return out
}

const TTC = buildCollection([REG_BYTES, BOLD_BYTES])

describe('a synthesized collection: the unwrapped member equals the standalone font', () => {
	test('the fixture really is a collection, and holds both fonts', () => {
		expect(String.fromCharCode(...TTC.subarray(0, 4))).toBe('ttcf')
		expect(isFontCollection(TTC)).toBe(true)
		expect(isFontCollection(REG_BYTES)).toBe(false)
		// Sharing is what a collection is for: storing each table once makes the container
		// smaller than the two fonts concatenated.
		expect(TTC.byteLength).toBeLessThan(REG_BYTES.byteLength + BOLD_BYTES.byteLength)
	})

	test('listFontFaces reads each member name out of its own name table', () => {
		const faces = listFontFaces(TTC)
		expect(faces.map((f) => f.index)).toEqual([0, 1])
		expect(faces[0].family).toBe('Silkscreen')
		expect(faces[0].subfamily).toBe('Regular')
		expect(faces[0].postScriptName).toBe('Silkscreen-Regular')
		expect(faces[1].subfamily).toBe('Bold')
		expect(faces[1].postScriptName).toBe('Silkscreen-Bold')
	})

	test('a plain TTF is a one-entry list, so a caller never branches on the container', () => {
		const faces = listFontFaces(REG_BYTES)
		expect(faces).toHaveLength(1)
		expect(faces[0]).toMatchObject({ index: 0, family: 'Silkscreen', subfamily: 'Regular' })
	})

	test('member 0 advances are IDENTICAL to the standalone regular .ttf', async () => {
		const fromCollection = await parseFontMetrics(TTC, { font: 0 })
		const standalone = await parseFontMetrics(REG_BYTES)
		expect(fromCollection.unitsPerEm).toBe(standalone.unitsPerEm)
		expect(advanceProfile(fromCollection)).toEqual(advanceProfile(standalone))
	})

	test('member 1 advances are IDENTICAL to the standalone bold .ttf', async () => {
		const fromCollection = await parseFontMetrics(TTC, { font: 1 })
		const standalone = await parseFontMetrics(BOLD_BYTES)
		expect(advanceProfile(fromCollection)).toEqual(advanceProfile(standalone))
	})

	test('sensitivity: the two members differ, so the equality above can fail', async () => {
		const regular = await parseFontMetrics(TTC, { font: 0 })
		const bold = await parseFontMetrics(TTC, { font: 1 })
		expect(advanceProfile(regular)).not.toEqual(advanceProfile(bold))
	})

	test('cmap coverage survives the unwrap', async () => {
		const fromCollection = await parseFontMetrics(TTC, { font: 0 })
		const standalone = await parseFontMetrics(REG_BYTES)
		for (const cp of [0x41, 0x7a, 0x30, 0x2011, 0x65e5, 0x1f600]) {
			expect(fromCollection.hasCodepoint(cp), `U+${cp.toString(16)}`).toBe(standalone.hasCodepoint(cp))
		}
		expect(standalone.hasCodepoint(0x41)).toBe(true) // the sweep above is not vacuous
		expect(standalone.hasCodepoint(0x65e5)).toBe(false)
	})
})

describe('selecting a font by name', () => {
	test('family, full and PostScript names all select, case-insensitively', async () => {
		const byIndex = advanceProfile(await parseFontMetrics(TTC, { font: 1 }))
		for (const name of ['Silkscreen-Bold', 'silkscreen-bold', 'SILKSCREEN-BOLD']) {
			expect(advanceProfile(await parseFontMetrics(TTC, { font: name })), name).toEqual(byIndex)
		}
	})

	test('an ambiguous family name selects the first member carrying it', async () => {
		// Both Silkscreen members share family 'Silkscreen'; only the PostScript name
		// separates them. First-wins is the documented rule, and it is index order.
		const byName = advanceProfile(await parseFontMetrics(TTC, { font: 'Silkscreen' }))
		expect(byName).toEqual(advanceProfile(await parseFontMetrics(TTC, { font: 0 })))
	})
})

describe('a wrong selector is refused rather than quietly falling back', () => {
	// Falling back to the first font is the failure worth refusing: it measures one face's
	// advances while the caller believes it registered another's, and nothing downstream
	// can tell the difference.
	test('an out-of-range index throws InvalidOptionError with a stable code', async () => {
		await expect(parseFontMetrics(TTC, { font: 2 })).rejects.toMatchObject({
			constructor: InvalidOptionError,
			code: 'font/collection-index-out-of-range',
		})
	})

	test('a name that matches no member throws, and the message lists what is there', async () => {
		const err = await parseFontMetrics(TTC, { font: 'MS PGothic' }).catch((e) => e)
		expect(err).toBeInstanceOf(InvalidOptionError)
		expect(err.code).toBe('font/collection-face-not-found')
		expect(err.message).toContain('Silkscreen')
	})

	test('the selector means the same thing for a plain TTF, which is a one-entry list', async () => {
		await expect(parseFontMetrics(REG_BYTES, { font: 1 })).rejects.toMatchObject({
			code: 'font/collection-index-out-of-range',
		})
		await expect(parseFontMetrics(REG_BYTES, { font: 'Arial' })).rejects.toMatchObject({
			code: 'font/collection-face-not-found',
		})
		// ...and selecting the one font it holds works, by index or by name.
		expect((await parseFontMetrics(REG_BYTES, { font: 0 })).unitsPerEm).toBe(1000)
		expect((await parseFontMetrics(REG_BYTES, { font: 'Silkscreen' })).unitsPerEm).toBe(1000)
	})

	test('bytes that are not a font throw a MediaError, not a bare Error', async () => {
		const err = await parseFontMetrics(new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7])).catch((e) => e)
		expect(err).toBeInstanceOf(MediaError)
		expect(err.code).toBe('font/parse-failed')
	})

	test('a truncated collection header throws a MediaError rather than reading past the end', () => {
		const truncated = TTC.slice(0, 10)
		expect(() => listFontFaces(truncated)).toThrow(MediaError)
	})
})

describe('a malformed collection is refused, not read into garbage', () => {
	// These guards are the difference between a legible error and advances taken from the
	// wrong bytes: every one of them sits on a path where the file still *parses*.
	/** A copy of the good collection with one field rewritten. */
	function corrupt(mutate) {
		const bytes = TTC.slice()
		mutate(new DataView(bytes.buffer), bytes)
		return bytes
	}

	test('a collection declaring no fonts', () => {
		const bytes = corrupt((dv) => dv.setUint32(8, 0))
		expect(() => listFontFaces(bytes)).toThrow(/0 fonts/)
	})

	test('a collection declaring more fonts than its offset table can hold', () => {
		const bytes = corrupt((dv) => dv.setUint32(8, 0xffff))
		expect(() => listFontFaces(bytes)).toThrow(/offset table is truncated/)
	})

	test('a member whose directory runs past the end of the file', () => {
		// Member 0's directory sits right after the ttcf header; claiming 4096 tables makes
		// the records overrun without touching anything else.
		const dirOff = new DataView(TTC.buffer).getUint32(12)
		const bytes = corrupt((dv) => dv.setUint16(dirOff + 4, 4096))
		expect(() => listFontFaces(bytes)).toThrow(/directory is truncated/)
	})

	test('a member declaring no tables', () => {
		const dirOff = new DataView(TTC.buffer).getUint32(12)
		const bytes = corrupt((dv) => dv.setUint16(dirOff + 4, 0))
		expect(() => listFontFaces(bytes)).toThrow(/0 tables/)
	})

	test('a member whose first table would be overwritten by its own directory', async () => {
		// The unwrap writes the directory over the prologue, which is safe because a
		// well-formed collection always puts the first table after every member's
		// directory. A file that breaks that assumption is refused rather than having its
		// table data silently clobbered.
		const dirOff = new DataView(TTC.buffer).getUint32(12)
		const bytes = corrupt((dv) => {
			// Point the first table record at byte 16, inside the region about to be written.
			dv.setUint32(dirOff + 12 + 8, 16)
		})
		await expect(parseFontMetrics(bytes, { font: 0 })).rejects.toMatchObject({
			code: 'font/parse-failed',
		})
	})

	test('a member with no readable name table still lists, since metrics do not need one', async () => {
		// Point the name table's record past the end of the file so `readNames` bails out.
		// The entry survives with only its index, and a name selector then reports it as
		// `(unnamed)` rather than crashing while building the message.
		const dirOff = new DataView(TTC.buffer).getUint32(12)
		const bytes = corrupt((dv, raw) => {
			const numTables = dv.getUint16(dirOff + 4)
			for (let i = 0; i < numTables; i++) {
				const rec = dirOff + 12 + i * 16
				if (String.fromCharCode(...raw.subarray(rec, rec + 4)) === 'name') dv.setUint32(rec + 8, 0xfffffff0)
			}
		})
		expect(listFontFaces(bytes)[0]).toEqual({ index: 0 })
		const err = await parseFontMetrics(bytes, { font: 'Nothing Here' }).catch((e) => e)
		expect(err.code).toBe('font/collection-face-not-found')
		expect(err.message).toContain('(unnamed)')
	})

	test('a member whose directory offset points past the end of the file', () => {
		// The offset table is well-formed and its entry is read; the directory it names is
		// not there. Distinct from the two cases above, which corrupt a directory that the
		// file does at least contain.
		const bytes = corrupt((dv, raw) => dv.setUint32(12, raw.byteLength - 4))
		expect(() => listFontFaces(bytes)).toThrow(/Font table directory at \d+ is truncated/)
	})

	test('a buffer too short to hold even a tag is not a collection', () => {
		// `isFontCollection` is the first thing a caller reaches, on bytes from anywhere.
		// Reading four bytes out of three must answer false, not read past the end.
		expect(isFontCollection(new Uint8Array([0x74, 0x74, 0x63]))).toBe(false)
		expect(isFontCollection(new Uint8Array(0))).toBe(false)
	})

	// --- one bad `name` record, the rest of the table still good ---------------------
	//
	// The cases above break the whole table. These break a single record, which is the
	// harder shape: `readNames` must drop that one field and keep the others, because a
	// face is still measurable — and still selectable by its other names — without it.

	/** Offset of member 0's `name` table inside `TTC`, and its declared length. */
	function nameTable(bytes) {
		const dv = new DataView(bytes.buffer)
		const dirOff = dv.getUint32(12)
		const numTables = dv.getUint16(dirOff + 4)
		for (let i = 0; i < numTables; i++) {
			const rec = dirOff + 12 + i * 16
			if (String.fromCharCode(...bytes.subarray(rec, rec + 4)) === 'name')
				return { base: dv.getUint32(rec + 8), length: dv.getUint32(rec + 12) }
		}
		throw new Error('member 0 has no name table')
	}

	/** Offset of the record for `nameID` within member 0's `name` table. */
	function nameRecord(bytes, nameID) {
		const { base } = nameTable(bytes)
		const dv = new DataView(bytes.buffer)
		const count = dv.getUint16(base + 2)
		for (let i = 0; i < count; i++) {
			const rec = base + 6 + i * 12
			if (dv.getUint16(rec + 6) === nameID) return rec
		}
		throw new Error(`member 0 has no name record ${nameID}`)
	}

	test('a record whose string runs past the end of the file drops that field only', () => {
		// Silkscreen carries exactly one record per name ID, so pointing the family's
		// string outside the file is enough to lose it. The face keeps its index and its
		// other names, and stays selectable by the PostScript name.
		const FAMILY = 1
		const bytes = corrupt((dv, raw) => dv.setUint16(nameRecord(raw, FAMILY) + 10, 0xffff))
		const face = listFontFaces(bytes)[0]
		expect(face.family).toBeUndefined()
		expect(face.postScriptName).toBe('Silkscreen-Regular')
		expect(listFontFaces(TTC)[0].family).toBe('Silkscreen') // the field is there to lose
	})

	test('a record on an unrecognized platform is still read', async () => {
		// Platform 2 (ISO) is deprecated and ranks below every platform this reader knows,
		// but it is the only record for its ID here, so it is what the family name comes
		// from. Ignoring an unknown platform outright would silently unname the face.
		const FAMILY = 1
		const bytes = corrupt((dv, raw) => dv.setUint16(nameRecord(raw, FAMILY), 2))
		expect(listFontFaces(bytes)[0].family).toBe('Silkscreen')
		expect((await parseFontMetrics(bytes, { font: 'Silkscreen' })).unitsPerEm).toBe(1000)
	})

	test('a record count larger than the file stops at the end, not past it', () => {
		// The declared table length still fits, so the early bail-out does not catch this:
		// only the per-record bounds check stands between the walk and the end of the file.
		const bytes = corrupt((dv, raw) => dv.setUint16(nameTable(raw).base + 2, 0xffff))
		const faces = listFontFaces(bytes)
		expect(faces).toHaveLength(2)
		// Member 1's own table is untouched, so it reads exactly as before.
		expect(faces[1]).toMatchObject({ family: 'Silkscreen', postScriptName: 'Silkscreen-Bold' })
	})
})

describe('pptx.registerFontMetrics accepts a collection', () => {
	test('with no font option the deck-side face name selects the member', async () => {
		const pptx = new TsPptx()
		await pptx.registerFontMetrics('Silkscreen', TTC)
		const measured = pptx.measureText('WWWWWWWWWW', { wIn: 1, fontSize: 18, fontFace: 'Silkscreen' })
		expect(measured.measurable).toBe(true)
	})

	test('a face naming no member throws instead of silently measuring the first', async () => {
		const pptx = new TsPptx()
		const err = await pptx.registerFontMetrics('MS PGothic', TTC).catch((e) => e)
		expect(err).toBeInstanceOf(InvalidOptionError)
		expect(err.code).toBe('font/collection-face-not-found')
	})

	test('an explicit font option overrides the face name', async () => {
		const pptx = new TsPptx()
		// The deck calls it 'Pixel'; the file calls it Silkscreen-Bold.
		await pptx.registerFontMetrics('Pixel', TTC, { font: 'Silkscreen-Bold' })
		expect(pptx.measureText('W', { wIn: 4, fontSize: 18, fontFace: 'Pixel' }).measurable).toBe(true)
	})

	test('a plain TTF still registers under any face name (no name matching there)', async () => {
		const pptx = new TsPptx()
		await pptx.registerFontMetrics('Anything At All', REG_BYTES)
		expect(pptx.measureText('W', { wIn: 4, fontSize: 18, fontFace: 'Anything At All' }).measurable).toBe(true)
	})

	test('a raw ArrayBuffer registers like the Uint8Array view over it', async () => {
		// The documented third source shape, and the one a browser caller has after
		// `await response.arrayBuffer()`. Same bytes, so the same advances.
		const buffer = new ArrayBuffer(TTC.byteLength)
		new Uint8Array(buffer).set(TTC)
		expect(buffer).toBeInstanceOf(ArrayBuffer)
		const fromBuffer = new TsPptx()
		await fromBuffer.registerFontMetrics('Silkscreen', buffer)
		const fromView = new TsPptx()
		await fromView.registerFontMetrics('Silkscreen', TTC)
		const opts = { wIn: 4, fontSize: 18, fontFace: 'Silkscreen' }
		expect(fromBuffer.measureText('WWWWWWWWWW', opts).widestLineIn).toBe(
			fromView.measureText('WWWWWWWWWW', opts).widestLineIn
		)
		expect(fromBuffer.measureText('WWWWWWWWWW', opts).measurable).toBe(true)
	})
})

// --- the genuine Windows collections ------------------------------------------------

const ORACLE_PATH = fixture('windows-collections.oracle.json')
const FONT_DIR = `${process.env.SystemRoot ?? 'C:\\Windows'}\\Fonts`
const oracle = existsSync(ORACLE_PATH) ? JSON.parse(readFileSync(ORACLE_PATH, 'utf8')) : null
const oracleFaces = (oracle?.faces ?? []).filter((f) => existsSync(`${FONT_DIR}\\${f.file}`))

describe.skipIf(process.platform !== 'win32' || oracleFaces.length === 0)(
	'genuine Windows collections agree with WPF, an independent reader',
	() => {
		const bytesFor = new Map()
		const load = (file) => {
			if (!bytesFor.has(file)) bytesFor.set(file, new Uint8Array(readFileSync(`${FONT_DIR}\\${file}`)))
			return bytesFor.get(file)
		}

		test('the oracle still describes the fonts on this machine', () => {
			// A Windows font update that adds or drops a collection member should fail here
			// rather than silently shrinking what the next test compares.
			expect(oracleFaces.length).toBeGreaterThan(0)
			for (const face of oracleFaces) {
				expect(listFontFaces(load(face.file))[face.index], `${face.file}#${face.index}`).toMatchObject({
					index: face.index,
					family: face.family,
				})
			}
		})

		test('every advance and every coverage decision matches GlyphTypeface exactly', async () => {
			let compared = 0
			for (const face of oracleFaces) {
				const metrics = await parseFontMetrics(load(face.file), { font: face.index })
				for (const [label, em] of Object.entries(face.advances)) {
					const cp = Number.parseInt(label.slice(2), 16)
					const where = `${face.family} (${face.file}#${face.index}) ${label}`
					// The oracle omits a code point the member's cmap lacks, so everything
					// listed must be covered - which fails if the wrong member's cmap came back.
					expect(metrics.hasCodepoint(cp), `${where} coverage`).toBe(true)
					// 1000pt makes an em 1000pt, so this compares design units, not rounding.
					expect(metrics.advanceWidthPt(String.fromCodePoint(cp), 1000) / 1000, where).toBeCloseTo(em, 5)
					compared++
				}
			}
			expect(compared).toBeGreaterThan(100)
		})

		test('members of one collection are genuinely distinct, not the same font read twice', async () => {
			// msgothic.ttc is the sharp case: an unwrapper that ignored the selector would
			// return equal profiles here and still pass every other test in this file.
			// Both scripts are needed. Latin separates monospaced MS Gothic (A = 0.5 em)
			// from the two proportional members (0.6328), but MS UI Gothic and MS PGothic
			// have the SAME Latin advances - only Kana tells those apart (の advances
			// 0.8164 em in MS UI Gothic against a full em in MS PGothic).
			const msgothic = `${FONT_DIR}\\msgothic.ttc`
			if (!existsSync(msgothic)) return
			const bytes = new Uint8Array(readFileSync(msgothic))
			const profiles = await Promise.all(
				listFontFaces(bytes).map(async (f) => {
					const m = await parseFontMetrics(bytes, { font: f.index })
					return [...'AMiのテ'].map((ch) => m.advanceWidthPt(ch, 1000))
				})
			)
			expect(new Set(profiles.map((p) => p.join(','))).size).toBe(profiles.length)
		})
	}
)
