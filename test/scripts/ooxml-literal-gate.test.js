// The OOXML literal gate: what its scan counts, and how it reads the allowlist.
//
// A gate like this fails silent. A pattern that stops matching, or a comparison that stops
// flagging, reports a clean tree while literals creep back outside `src/ooxml/`. These pin both.

import { describe, expect, test } from 'vitest'
import { OOXML_LITERAL, compareToAllowlist } from '../../scripts/ooxml-literal-gate.mjs'
import { scanSource } from '../../scripts/raw-xml-ratchet.mjs'

/**
 * The literals the gate's scan reports for a snippet.
 * @param {string} source
 * @returns {string[]}
 */
const scan = (source) => scanSource(source, 'input.ts', OOXML_LITERAL).map((finding) => finding.text)

describe('what counts', () => {
	test('a schema URI in a string literal, whole', () => {
		expect(scan(`const NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'`)).toEqual([
			'http://schemas.openxmlformats.org/drawingml/2006/main',
		])
	})

	test('a vendor content type, whole or as a prefix', () => {
		const source = `const A = 'application/vnd.ms-office.chartex+xml'\nconst B = "application/vnd.openxmlformats-package."`
		expect(scan(source)).toEqual(['application/vnd.ms-office.chartex+xml', 'application/vnd.openxmlformats-package.'])
	})

	test('one inside a template of XML stops at its closing quote', () => {
		expect(scan('const x = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`')).toEqual([
			'http://schemas.openxmlformats.org/package/2006/content-types',
		])
	})

	test('a non-OOXML namespace or media type is not a finding', () => {
		expect(scan(`const a = 'http://purl.org/dc/terms/'\nconst b = 'image/png'`)).toEqual([])
	})
})

describe('what does not, for the ratchet’s reasons', () => {
	test('a doc comment', () => {
		expect(scan(`/** Reads the http://schemas.openxmlformats.org/drawingml/2006/main root. */\nconst a = 1`)).toEqual(
			[]
		)
	})

	test('a literal handed to a message sink', () => {
		expect(scan(`warn('code', 'expected application/vnd.ms-office.chartex+xml')`)).toEqual([])
	})

	test('a declaration marked as XML captured from Office', () => {
		const source = `/** @raw-xml-asset */\nconst T = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'`
		expect(scan(source)).toEqual([])
	})
})

describe('the allowlist', () => {
	const finding = { file: 'src/a.ts', line: 3, literal: 'http://schemas.example.org/one' }
	const entry = { file: 'src/a.ts', literal: 'http://schemas.example.org/one', reason: 'only this module reads it' }

	test('a listed literal in its own file passes', () => {
		expect(compareToAllowlist([finding], [entry])).toEqual({ unlisted: [], stale: [], reasonless: [] })
	})

	test('the same literal in another file is unlisted there', () => {
		const elsewhere = { ...finding, file: 'src/b.ts' }
		const result = compareToAllowlist([elsewhere], [entry])
		expect(result.unlisted).toEqual([elsewhere])
		expect(result.stale).toEqual([entry])
	})

	test('a new literal in an allowlisted file is unlisted', () => {
		const added = { ...finding, literal: 'http://schemas.example.org/two' }
		expect(compareToAllowlist([finding, added], [entry]).unlisted).toEqual([added])
	})

	test('an entry whose literal is gone is stale', () => {
		expect(compareToAllowlist([], [entry]).stale).toEqual([entry])
	})

	test('an entry without a reason is reported', () => {
		const bare = { ...entry, reason: '  ' }
		expect(compareToAllowlist([finding], [bare]).reasonless).toEqual([bare])
	})
})
