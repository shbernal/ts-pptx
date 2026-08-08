// Unit tests for the write-side XML element builder (`src/gen/oxml/el.ts`).
//
// Imports the source directly (precedent: font-heuristic.test.mjs) because the
// builder is an internal substrate not re-exported from any dist entry point.
// It still earns dist coverage indirectly via the emitters built on it.
//
// The behaviors pinned here are the ones a byte-identity migration depends on:
// paired-vs-self-closing is decided by ARITY, attributes escape and drop
// nullish, and whitespace is emitted exactly where `fmt` says.

import { describe, test, expect } from 'vitest'
import { el, voidEl, raw } from '../../../src/gen/oxml/el.ts'

describe('el() / voidEl() XML builder', () => {
	test('emits a paired tag with escaped text', () => {
		expect(el('dc:title', null, 'Q1 & Q2')).toBe('<dc:title>Q1 &amp; Q2</dc:title>')
	})

	test('emits a paired tag even when the child is empty or undefined', () => {
		// REGRESSION: encodeXmlEntities(undefined) === '', so a builder that
		// decided self-closing from the child VALUE would emit `<dc:title/>` and
		// silently change the bytes of every optional-text element.
		expect(el('dc:title', null, undefined)).toBe('<dc:title></dc:title>')
		expect(el('dc:title', null, '')).toBe('<dc:title></dc:title>')
		expect(el('dc:title')).toBe('<dc:title></dc:title>')
	})

	test('voidEl() self-closes', () => {
		expect(voidEl('a:noFill')).toBe('<a:noFill/>')
		expect(voidEl('Relationship', { Id: 'rId1' })).toBe('<Relationship Id="rId1"/>')
	})

	test('voidEl() closePrefix goes before the self-closing slash', () => {
		// Several DrawingML emitters write `<a:avLst />` with a space; that space is
		// byte-significant, so it has to be expressible rather than normalized away.
		expect(voidEl('a:avLst', null, { closePrefix: ' ' })).toBe('<a:avLst />')
		expect(voidEl('a:rect', { l: 'l', b: 'b' }, { closePrefix: ' ' })).toBe('<a:rect l="l" b="b" />')
	})

	test('escapes attribute values and preserves insertion order', () => {
		expect(voidEl('p:tag', { name: 'A&B', val: '"<x>"' })).toBe('<p:tag name="A&amp;B" val="&quot;&lt;x&gt;&quot;"/>')
	})

	test('attribute values escape tab/CR/LF as character references', () => {
		// REGRESSION (dn-xml-attr-whitespace): XML 1.0 section 3.3.3 has a parser normalise a
		// LITERAL tab/CR/LF inside an attribute value to a single space before any consumer sees
		// it, so emitting one raw silently destroys a caller's line break. Text children keep the
		// literal character, where it is content — the two paths use different escapers.
		expect(voidEl('p14:section', { name: 'Abschnitts-\nuberschrift' })).toBe(
			'<p14:section name="Abschnitts-&#10;uberschrift"/>'
		)
		expect(voidEl('p:tag', { val: 'a\tb\r\nc' })).toBe('<p:tag val="a&#9;b&#13;&#10;c"/>')
		expect(el('a:t', null, 'line1\nline2')).toBe('<a:t>line1\nline2</a:t>')
	})

	test('omits nullish attributes but keeps empty string and zero', () => {
		expect(voidEl('a:t', { a: undefined, b: null, c: '', d: 0 })).toBe('<a:t c="" d="0"/>')
	})

	test('raw() children are interpolated verbatim, text children are escaped', () => {
		expect(el('p:sp', null, [raw('<a:off x="1"/>'), '5 > 4'])).toBe('<p:sp><a:off x="1"/>5 &gt; 4</p:sp>')
	})

	test('nullish children are skipped, so conditionals inline cleanly', () => {
		const optional = false
		expect(el('p:x', null, [raw('<a/>'), optional ? raw('<b/>') : null, undefined, raw('<c/>')])).toBe(
			'<p:x><a/><c/></p:x>'
		)
	})

	test('strips XML-illegal control characters', () => {
		// \v and \x07 are illegal in XML 1.0 and trigger a PowerPoint repair dialog.
		expect(el('a:t', null, 'a\u000bb\u0007c')).toBe('<a:t>abc</a:t>')
	})

	test('fmt places whitespace exactly, including irregular indentation', () => {
		// The misindented closing tag is real: makeXmlRootRels closes with
		// `\n\t\t</Relationships>` despite children at the same depth.
		const xml = el(
			'Relationships',
			{ xmlns: 'ns' },
			[raw(voidEl('Relationship', { Id: 'rId1' }, { openPrefix: '\n\t\t' }))],
			{
				closePrefix: '\n\t\t',
			}
		)
		expect(xml).toBe('<Relationships xmlns="ns">\n\t\t<Relationship Id="rId1"/>\n\t\t</Relationships>')
	})

	test('childPrefix applies to every child', () => {
		expect(el('r', null, [raw('<a/>'), raw('<b/>')], { childPrefix: '\n\t' })).toBe('<r>\n\t<a/>\n\t<b/></r>')
	})
})
