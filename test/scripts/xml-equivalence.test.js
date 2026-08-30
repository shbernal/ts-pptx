// The whitespace-equivalence prover.
//
// This module's job is to discharge an AGENTS.md STOP: a whitespace-only byte diff is
// refused unless a program can prove that is all it is. That makes a false PASS the worst
// thing it can do — it would wave through the exact content change the STOP exists to
// catch, while looking like evidence. So the tests here are weighted toward the red cases:
// every kind of difference that is NOT inert whitespace gets one, and each is written so
// that a prover which merely stripped whitespace and compared would fail it.
//
// AGENTS.md, on the render oracle: "Make any new case fail on purpose before trusting it."
// That is what the `rejects` blocks are.

import { describe, expect, test } from 'vitest'
import {
	XmlSyntaxError,
	isTextFrozen,
	proveWhitespaceOnly,
	tokenizeXml,
	buildTree,
} from '../../scripts/xml-equivalence.mjs'

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

// The three helpers take the proof result unannotated on purpose. `ProofPass`/`ProofFail`
// discriminate on a literal `ok`, and this project runs with `strictNullChecks: false`
// (see the reasoning in `tsconfig.test.json`), under which TypeScript does not narrow a
// union by its discriminant at all. Annotating them would therefore not buy the checking
// it looks like it buys; it would just make every `.reason` read an error.

/** Prove `current` against `base`, both wrapped in the declaration every real part carries. */
const prove = (base, current) => proveWhitespaceOnly(DECL + base, DECL + current)

/**
 * The failure text, or a marker that reads wrong in an assertion if the proof unexpectedly
 * passed -- so a red case that silently turns green fails on the message rather than passing.
 */
const reasonOf = (result) => (result.ok ? '(the proof PASSED; there is no reason)' : result.reason)

/**
 * The relaxed positions, or a throw naming why the proof failed. The throw is the point: a
 * green case that goes red should say what the prover objected to, not just miss a count.
 */
const relaxationsOf = (result) => {
	if (!result.ok) throw new Error('expected a proof, got a failure at ' + result.path + ': ' + result.reason)
	return result.relaxations
}

describe('accepts inert whitespace', () => {
	test('identical input proves with nothing relaxed', () => {
		const xml = '<c:ser><c:idx val="0"/></c:ser>'
		const result = prove(xml, xml)
		expect(result.ok).toBe(true)
		expect(relaxationsOf(result)).toEqual([])
	})

	// The actual shape of the change this was built for: `<c:ser>  <c:idx .../>` losing the
	// two-space run that `plot-*.ts` passes as `openPrefix`.
	test('a space run between element children is relaxed', () => {
		const result = prove(
			'<c:ser>  <c:idx val="0"/>  <c:order val="0"/></c:ser>',
			'<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>'
		)
		expect(result.ok).toBe(true)
		expect(relaxationsOf(result)).toHaveLength(1)
	})

	test('a newline-and-indent run is relaxed', () => {
		const result = prove('<c:title>\n      <c:tx/>\n    </c:title>', '<c:title><c:tx/></c:title>')
		expect(result.ok).toBe(true)
	})

	// `<a:p>` holds runs, and PowerPoint takes its text from `<a:t>` alone — so whitespace
	// between an `<a:pPr>` and an `<a:r>` is layout. This is the one relaxation that touches
	// the text tree at all, so it is pinned deliberately rather than left implied.
	test('whitespace between runs inside a paragraph is relaxed', () => {
		const result = prove(
			'<a:p>\n  <a:pPr/>\n  <a:r><a:t>Q1</a:t></a:r>\n</a:p>',
			'<a:p><a:pPr/><a:r><a:t>Q1</a:t></a:r></a:p>'
		)
		expect(result.ok).toBe(true)
	})

	test('whitespace outside the root element is relaxed', () => {
		const result = prove('<a:x/>\n', '<a:x/>')
		expect(result.ok).toBe(true)
	})
})

describe('rejects character-data changes', () => {
	test('text content changed', () => {
		const result = prove('<c:v>12</c:v>', '<c:v>13</c:v>')
		expect(result.ok).toBe(false)
		expect(reasonOf(result)).toContain('character data changed')
	})

	// The case the structural rule exists for. `<a:t>` has no element children, so its text
	// is frozen whatever it contains — including when what it contains is only spaces.
	test('interior spacing inside a text leaf', () => {
		const result = prove('<a:t>a b</a:t>', '<a:t>a  b</a:t>')
		expect(result.ok).toBe(false)
	})

	test('a text leaf whose entire content is whitespace', () => {
		const result = prove('<a:t>  </a:t>', '<a:t></a:t>')
		expect(result.ok).toBe(false)
	})

	test('leading whitespace on text content', () => {
		const result = prove('<c:formatCode> #,##0</c:formatCode>', '<c:formatCode>#,##0</c:formatCode>')
		expect(result.ok).toBe(false)
	})

	// Mixed content: the spaces flanking the text are part of the text, and the presence of
	// an element child must not unfreeze them.
	test('whitespace beside text in mixed content', () => {
		const result = prove('<x:mixed> hello <x:b/> </x:mixed>', '<x:mixed>hello<x:b/></x:mixed>')
		expect(result.ok).toBe(false)
	})

	// The deny-list case: `<si>` has element children and only whitespace between them, so
	// the structural rule alone would relax it. The xlsx string table is not in scope.
	test('whitespace between children of a text-bearing container', () => {
		const result = prove('<si>  <t>North</t></si>', '<si><t>North</t></si>')
		expect(result.ok).toBe(false)
	})
})

describe('rejects escaping changes', () => {
	// A DOM comparison cannot see either of these: both sides decode to the same character.
	// `gen/oxml/el.ts` splits attribute and text escaping precisely because they differ, so
	// a prover blind to the difference would be blind to a regression in the split.
	test('an entity reference replaced by its character in text', () => {
		const result = prove('<c:v>a&amp;b</c:v>', '<c:v>a&#38;b</c:v>')
		expect(result.ok).toBe(false)
	})

	test('a character reference replaced by its character in an attribute', () => {
		const result = prove('<a:t val="a&#10;b"/>', '<a:t val="a\nb"/>')
		expect(result.ok).toBe(false)
	})
})

describe('rejects structural changes', () => {
	test('an attribute value changed', () => {
		const result = prove('<c:idx val="0"/>', '<c:idx val="1"/>')
		expect(result.ok).toBe(false)
		expect(reasonOf(result)).toContain('@val changed')
	})

	// Inert per the XML spec, and deliberately still a failure: this proves a whitespace-only
	// change, and unifying the two `<a:defRPr>` orderings in `chart-parts.ts` is a separate
	// change that must not ride along under this gate.
	test('attributes reordered', () => {
		const result = prove('<a:defRPr b="0" sz="1200"/>', '<a:defRPr sz="1200" b="0"/>')
		expect(result.ok).toBe(false)
		expect(reasonOf(result)).toContain('attribute 0')
	})

	test('an attribute added', () => {
		const result = prove('<c:idx val="0"/>', '<c:idx val="0" extra="1"/>')
		expect(result.ok).toBe(false)
	})

	test('self-closing form changed', () => {
		const result = prove('<a:effectLst/>', '<a:effectLst></a:effectLst>')
		expect(result.ok).toBe(false)
		expect(reasonOf(result)).toContain('became')
	})

	test('an element added', () => {
		const result = prove('<c:ser><c:idx val="0"/></c:ser>', '<c:ser><c:idx val="0"/><c:order val="0"/></c:ser>')
		expect(result.ok).toBe(false)
	})

	test('an element removed', () => {
		const result = prove('<c:ser>  <c:idx val="0"/>  <c:order val="0"/></c:ser>', '<c:ser><c:idx val="0"/></c:ser>')
		expect(result.ok).toBe(false)
	})

	// Sibling order is semantic in OOXML — the schemas are sequences, and PowerPoint reports
	// a repair on a misordered child. Stripping whitespace and comparing the rest as a set
	// would miss this; comparing as a list is what catches it.
	test('siblings reordered', () => {
		const result = prove(
			'<c:ser>  <c:idx val="0"/>  <c:order val="1"/></c:ser>',
			'<c:ser><c:order val="1"/><c:idx val="0"/></c:ser>'
		)
		expect(result.ok).toBe(false)
	})

	test('an element renamed', () => {
		const result = prove('<c:ser><c:idx val="0"/></c:ser>', '<c:ser><c:order val="0"/></c:ser>')
		expect(result.ok).toBe(false)
	})

	test('the XML declaration changed', () => {
		const result = proveWhitespaceOnly(DECL + '<a:x/>', '<?xml version="1.0" encoding="UTF-8"?><a:x/>')
		expect(result.ok).toBe(false)
		expect(reasonOf(result)).toContain('declaration')
	})
})

describe('the tokenizer refuses what it does not understand', () => {
	// A scanner that skipped these would be agreeing that two parts match on the strength of
	// bytes it never looked at. None appear in the emitted corpus; if one starts to, the
	// gate should stop rather than shrug.
	test('a comment', () => {
		expect(() => tokenizeXml('<a:x><!-- note --></a:x>')).toThrow(XmlSyntaxError)
	})

	test('a CDATA section', () => {
		expect(() => tokenizeXml('<a:x><![CDATA[raw]]></a:x>')).toThrow(XmlSyntaxError)
	})

	test('a DOCTYPE', () => {
		expect(() => tokenizeXml('<!DOCTYPE x><a:x/>')).toThrow(XmlSyntaxError)
	})

	test('a foreign processing instruction', () => {
		expect(() => tokenizeXml('<?mso-application progid="Word.Document"?><a:x/>')).toThrow(XmlSyntaxError)
	})

	test('an unquoted attribute value', () => {
		expect(() => tokenizeXml('<a:x val=1/>')).toThrow(XmlSyntaxError)
	})

	test('a valueless attribute', () => {
		expect(() => tokenizeXml('<a:x checked/>')).toThrow(XmlSyntaxError)
	})

	test('a mismatched end tag', () => {
		expect(() => buildTree(tokenizeXml('<a:x></a:y>'))).toThrow(XmlSyntaxError)
	})

	test('an unclosed element', () => {
		expect(() => buildTree(tokenizeXml('<a:x><a:y></a:x>'))).toThrow(XmlSyntaxError)
	})
})

describe('the tokenizer preserves what it must', () => {
	test('attribute order, raw value and quote character survive', () => {
		const [, open] = tokenizeXml(DECL + '<a:x b="1" a="&amp;"/>')
		if (open?.kind !== 'open') throw new Error('expected an open tag, got ' + open?.kind)
		expect(open).toEqual({
			kind: 'open',
			name: 'a:x',
			selfClosing: true,
			endGap: '',
			attrs: [
				{ name: 'b', gap: ' ', value: '1', quote: '"' },
				{ name: 'a', gap: ' ', value: '&amp;', quote: '"' },
			],
		})
	})

	test('a single-quoted attribute value is read as such', () => {
		const [open] = tokenizeXml("<a:x a='1'/>")
		if (open?.kind !== 'open') throw new Error('expected an open tag, got ' + open?.kind)
		expect(open.attrs[0]?.quote).toBe("'")
	})
})

describe('isTextFrozen', () => {
	/** @param {string} xml @returns {boolean} */
	const frozen = (xml) => {
		const first = buildTree(tokenizeXml(xml)).children[0]
		if (first?.type !== 'element') throw new Error('expected an element, got ' + first?.type)
		return isTextFrozen(first)
	}

	test('a leaf with text is frozen', () => {
		expect(frozen('<c:v>12</c:v>')).toBe(true)
	})

	test('a leaf with only whitespace is frozen', () => {
		expect(frozen('<x:leaf>  </x:leaf>')).toBe(true)
	})

	test('an empty element is frozen', () => {
		expect(frozen('<x:leaf></x:leaf>')).toBe(true)
	})

	test('a container of elements separated by whitespace is not frozen', () => {
		expect(frozen('<x:box>  <x:a/>  <x:b/></x:box>')).toBe(false)
	})

	test('mixed content is frozen', () => {
		expect(frozen('<x:box>text<x:a/></x:box>')).toBe(true)
	})

	test('a name on the deny-list is frozen even when it looks like a container', () => {
		expect(frozen('<si>  <t>a</t></si>')).toBe(true)
	})
})

describe('whitespace inside a start tag is frozen too', () => {
	// Also inert per the XML spec, and deliberately not relaxed. `chart-parts.ts` emits a
	// double space before `b=` when the caller set no font size, and that is a real emitted
	// byte this gate has no business absorbing: it proves a claim about whitespace BETWEEN
	// elements, and quietly covering a second category would make the claim untrue.
	test('a doubled space between attributes', () => {
		const result = prove('<a:defRPr  b="0"/>', '<a:defRPr b="0"/>')
		expect(result.ok).toBe(false)
		expect(reasonOf(result)).toContain('whitespace before @b')
	})

	test('a space before the self-closing delimiter', () => {
		const result = prove('<a:avLst />', '<a:avLst/>')
		expect(result.ok).toBe(false)
		expect(reasonOf(result)).toContain('closing delimiter')
	})

	test('a newline between attributes', () => {
		const result = prove('<a:x a="1"\n b="2"/>', '<a:x a="1" b="2"/>')
		expect(result.ok).toBe(false)
	})
})
