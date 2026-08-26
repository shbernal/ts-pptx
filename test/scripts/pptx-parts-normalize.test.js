// The byte-identity harness's normalizers.
//
// The gate's whole value is that any byte change is a real regression, so a normalizer
// that matches one character too widely silently excuses a diff it was never meant to
// cover — and the gate stays green while saying nothing. Every entry in `NORMALIZERS`
// therefore needs both halves pinned: what it erases, and the neighbouring construct it
// must leave alone.

import { describe, expect, test } from 'vitest'
import { normalize } from '../../scripts/pptx-parts.mjs'

describe('scatter data-label field ids', () => {
	// `gen/chart/plot-scatter.ts` mints a fresh GUID per field, so a deck with
	// `dataLabelFormatScatter: 'custom'`/`'customXY'` emits different bytes on every build.
	// A field id has to be unique, so the id is right and the *comparison* is what has to
	// give — the same call the `c16:uniqueId` two lines above it gets.
	test('an XVALUE field id is erased', () => {
		expect(normalize('<a:fld id="{650c7360-4b87-0998-da56-aae7e835e166}" type="XVALUE">')).toBe(
			'<a:fld id="{NORMALIZED-SCATTERFLD}" type="XVALUE">'
		)
	})

	test('a YVALUE field id is erased', () => {
		expect(normalize('<a:fld id="{fc5e702d-9156-1769-e9bf-a3b613076579}" type="YVALUE">')).toBe(
			'<a:fld id="{NORMALIZED-SCATTERFLD}" type="YVALUE">'
		)
	})

	test('both fields of one label are erased', () => {
		const label =
			'<a:fld id="{032f8e74-0352-3d73-c60b-2945be1a91a0}" type="XVALUE">' +
			'<a:fld id="{28a4cf09-4f70-5c9c-2397-ba6f1f82246b}" type="YVALUE">'
		expect(normalize(label)).toBe(
			'<a:fld id="{NORMALIZED-SCATTERFLD}" type="XVALUE"><a:fld id="{NORMALIZED-SCATTERFLD}" type="YVALUE">'
		)
	})

	// The other three `a:fld` emitters carry a constant id. Changing one of those IS a real
	// diff, which is exactly what a normalizer keyed on `a:fld` alone would have hidden.
	test('a fixed slidenum field id is left alone', () => {
		const fixed = '<a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum">'
		expect(normalize(fixed)).toBe(fixed)
	})

	test('a fixed datetime field id is left alone', () => {
		const fixed = '<a:fld id="{5282F153-3F37-0F45-9E97-73ACFA13230C}" type="datetimeFigureOut">'
		expect(normalize(fixed)).toBe(fixed)
	})
})

describe('the other three normalizers still bound', () => {
	test('a core.xml timestamp is erased, and only its own element', () => {
		const core =
			'<dcterms:created xsi:type="dcterms:W3CDTF">2026-08-26T05:00:00Z</dcterms:created>' +
			'<dc:title>2026-08-26T05:00:00Z</dc:title>'
		expect(normalize(core)).toBe(
			'<dcterms:created xsi:type="dcterms:W3CDTF">NORMALIZED-TIMESTAMP</dcterms:created>' +
				'<dc:title>2026-08-26T05:00:00Z</dc:title>'
		)
	})

	test('a section id is erased but the section name is not', () => {
		expect(normalize('<p14:section name="Intro" id="{4F1C2A3B-0000-0000-0000-000000000001}">')).toBe(
			'<p14:section name="Intro" id="{NORMALIZED-SECTION}">'
		)
	})

	test('a c16:uniqueId is erased', () => {
		expect(normalize('<c16:uniqueId val="{00000001-1234-5678-9abc-def012345678}"/>')).toBe(
			'<c16:uniqueId val="{NORMALIZED-UNIQUEID}"/>'
		)
	})
})
