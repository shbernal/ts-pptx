// Write→read round-trip for the document-properties accessors: build a deck with
// the write API's metadata setters (pptx.title/subject/author/revision) and
// pptx.setCustomProperty across every value type, serialize it, reload it via
// Presentation, and assert coreProperties / customProperties decode back to the
// values (and JS types) that went in. Because both parts have writers, this needs
// no hand-crafted fixture — the writer is the oracle. Two authored fixtures cover
// the missing-part and real-PowerPoint edges.

import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'
import TsPptx from '../../dist/node.js'
import { Presentation } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { openFixture } from './corpus.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

/** Build a deck with the given setup, serialize, and reload as a Presentation. */
async function roundTrip(setup) {
	const pptx = new TsPptx()
	setup(pptx)
	pptx.addSlide().addText('hello', { x: 1, y: 1, w: 4, h: 1 })
	const buf = /** @type {Uint8Array} */ (await pptx.stream())
	return Presentation.load(buf)
}

describe('Presentation.coreProperties', () => {
	test('round-trips the write-side title/subject/creator/revision', async () => {
		const pres = await roundTrip((pptx) => {
			pptx.title = 'Quarterly Review'
			pptx.subject = 'FY25 numbers'
			pptx.author = 'Ada Lovelace'
			pptx.revision = '7'
		})
		const core = pres.coreProperties
		assertEqual(core.title, 'Quarterly Review', 'title')
		assertEqual(core.subject, 'FY25 numbers', 'subject')
		// pptx.author writes dc:creator (and cp:lastModifiedBy).
		assertEqual(core.creator, 'Ada Lovelace', 'creator')
		assertEqual(core.lastModifiedBy, 'Ada Lovelace', 'lastModifiedBy')
		assertEqual(core.revision, '7', 'revision')
	})

	test('created/modified decode as raw W3CDTF strings (writer stamps now)', async () => {
		const pres = await roundTrip((pptx) => {
			pptx.title = 'Timestamps'
		})
		const core = pres.coreProperties
		// The writer stamps these to "now", so assert shape not value: an unparsed
		// W3CDTF string, deliberately not a Date object.
		assert(typeof core.created === 'string', `created is a string, got ${typeof core.created}`)
		assert(typeof core.modified === 'string', `modified is a string, got ${typeof core.modified}`)
		assert(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/.test(core.created), `created is W3CDTF, got ${core.created}`)
	})

	test('a PowerPoint-authored fixture decodes without throwing', async () => {
		const pres = await openFixture('read-stress')
		const core = pres.coreProperties
		// Real authored deck: core.xml is present, so this is a non-empty object; the
		// exact values are PowerPoint's, we only guarantee typed, throw-free decode.
		assert(typeof core === 'object' && core !== null, 'coreProperties is an object')
		for (const v of Object.values(core))
			assert(typeof v === 'string', `every present core field is a string, got ${typeof v}`)
	})
})

describe('Presentation.customProperties', () => {
	test('round-trips each value type back to its correct JS type', async () => {
		const when = new Date('2020-01-15T10:30:00Z')
		const pres = await roundTrip((pptx) => {
			pptx.setCustomProperty('Department', 'Finance')
			pptx.setCustomProperty('FiscalYear', 2025)
			pptx.setCustomProperty('Margin', 0.375)
			pptx.setCustomProperty('Approved', true)
			pptx.setCustomProperty('Rejected', false)
			pptx.setCustomProperty('ReviewedOn', when)
		})
		const byName = new Map(pres.customProperties.map((p) => [p.name, p.value]))

		assertEqual(byName.get('Department'), 'Finance', 'string stays string')
		assertEqual(byName.get('FiscalYear'), 2025, 'integer stays number')
		assert(byName.get('FiscalYear') === 2025 && typeof byName.get('FiscalYear') === 'number', 'integer is a number')
		assertEqual(byName.get('Margin'), 0.375, 'float stays number')
		assertEqual(byName.get('Approved'), true, 'true stays boolean true')
		assertEqual(byName.get('Rejected'), false, 'false stays boolean false')
		assert(typeof byName.get('Approved') === 'boolean', 'boolean is a boolean')
		// Date is written as vt:filetime and decoded as the raw W3CDTF string (no Date parse).
		assertEqual(byName.get('ReviewedOn'), '2020-01-15T10:30:00Z', 'Date → raw filetime string')
	})

	test('preserves property order and count', async () => {
		const pres = await roundTrip((pptx) => {
			pptx.setCustomProperty('First', 'a')
			pptx.setCustomProperty('Second', 'b')
			pptx.setCustomProperty('Third', 'c')
		})
		assertEqual(pres.customProperties.length, 3, 'three custom properties')
		assertEqual(pres.customProperties.map((p) => p.name).join(','), 'First,Second,Third', 'names in authored order')
	})

	test('a deck with no custom.xml → []', async () => {
		const pres = await roundTrip(() => {})
		assertEqual(pres.customProperties.length, 0, 'no custom properties')
		assert(Array.isArray(pres.customProperties), 'customProperties is an array')
	})

	test('an authored fixture with no custom properties → []', async () => {
		const pres = await openFixture('read-stress')
		assertEqual(pres.customProperties.length, 0, 'authored deck has no custom props')
	})
})
