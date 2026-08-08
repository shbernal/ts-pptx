// The raw-XML ratchet's exemption rules.
//
// The gate's budget is only meaningful if the scan keeps counting what it counted
// yesterday. Both carve-outs — the message-sink one and `@raw-xml-asset` — are
// judgement calls implemented as AST walks, and either over-matching would quietly
// drop occurrences from the count while the check stayed green. These pin the edges.

import { describe, expect, test } from 'vitest'
import { scanSource } from '../../scripts/raw-xml-ratchet.mjs'

/**
 * The matched delimiters.
 *
 * `TAG_DELIMITER` stops one character past the colon, so `<a:bodyPr` reports as `<a:b`.
 * That is deliberate — the gate counts occurrences, it does not extract element names —
 * and the expectations below are written against the truncated form on purpose.
 * @param {string} source
 * @returns {string[]}
 */
const scan = (source) => scanSource(source).map((finding) => finding.text)

describe('what counts', () => {
	test('a namespaced delimiter in a plain string literal', () => {
		expect(scan(`const x = '<a:bodyPr/>'`)).toEqual(['<a:b'])
	})

	test('a closing delimiter counts separately', () => {
		expect(scan('const x = `<a:t>hi</a:t>`')).toEqual(['<a:t', '</a:t'])
	})

	test('each span of a multi-part template is visited', () => {
		expect(scan('const x = `<a:t>${value}</a:t>`')).toEqual(['<a:t', '</a:t'])
	})

	test('an un-namespaced tag is not a finding — the gate is about OOXML markup', () => {
		expect(scan(`const x = '<div>plain html</div>'`)).toEqual([])
	})

	test('a doc comment is prose, not emitted bytes', () => {
		expect(scan(`/** Emits <a:bodyPr/> for the shape. */\nconst x = 1`)).toEqual([])
	})
})

describe('message-sink exemption', () => {
	test('a literal handed straight to warn() or note() is prose', () => {
		expect(scan(`warn('code', 'missing <a:bodyPr>')`)).toEqual([])
		expect(scan(`notes.note('unsupported <p:sp>')`)).toEqual([])
	})

	test('a literal handed to any *Error constructor is prose', () => {
		expect(scan(`throw new ShapeError('bad <a:xfrm>')`)).toEqual([])
		expect(scan(`throw new Error('bad <a:xfrm>')`)).toEqual([])
	})

	// "Nearest enclosing call" is the rule the header states. A literal nested inside a
	// builder that merely happens to sit in a warn() argument is still built markup.
	test('only the literal handed directly to the sink is exempt', () => {
		expect(scan('warn(`<a:t/>`, buildXml(`<a:bodyPr/>`))')).toEqual(['<a:b'])
	})

	test('an unrelated callee is not a sink', () => {
		expect(scan(`emit('<a:bodyPr/>')`)).toEqual(['<a:b'])
	})
})

describe('@raw-xml-asset exemption', () => {
	test('marks a captured constant as shipped-verbatim, not built', () => {
		expect(scan(`/** @raw-xml-asset */\nconst STYLE = '<cs:chartStyle/>'`)).toEqual([])
	})

	test('applies to a marked property declaration', () => {
		expect(scan(`class C {\n/** @raw-xml-asset */\nstatic X = '<cs:chartStyle/>'\n}`)).toEqual([])
	})

	// The narrowness the header promises: a template with substitutions is something
	// being *built*, so the marker must not reach it however it is labelled.
	test('does not exempt a template that interpolates', () => {
		expect(scan('/** @raw-xml-asset */\nconst X = `<cs:chartStyle val="${v}"/>`')).toEqual(['<cs:c'])
	})

	test('an unmarked neighbour is still counted', () => {
		const source = [`/** @raw-xml-asset */`, `const A = '<cs:aaa/>'`, `const B = '<cs:bbb/>'`].join('\n')
		expect(scan(source)).toEqual(['<cs:b'])
	})
})

describe('line numbers', () => {
	test('report the literal, so a finding is navigable', () => {
		const source = [`const a = 1`, `const b = 2`, `const c = '<a:t/>'`].join('\n')
		expect(scanSource(source)).toEqual([{ line: 3, text: '<a:t' }])
	})
})
