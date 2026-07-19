/**
 * Characterization tests for the OMML equation wrappers.
 *
 * These pin the emitted bytes of `genXmlMathParagraph` / `genXmlInlineMath`. The demo deck
 * used by the byte-identity harness contains no equations, so no `<a14:m>` markup reaches
 * that gate — these assertions are the only thing standing between a refactor of
 * `drawingml/math.ts` and a silent change to the equation envelope.
 *
 * Imports `src/` directly (precedent: `xml-el-builder.test.mjs`); contributes no dist coverage.
 */
import { describe, expect, test } from 'vitest'
import { genXmlMathParagraph, genXmlInlineMath } from '../../src/gen/drawingml/math.ts'

const NS =
	'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
const PARA_PR = '<m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>'
const OMATH = '<m:oMath><m:r><m:t>a+b</m:t></m:r></m:oMath>'

describe('genXmlMathParagraph()', () => {
	test('wraps a bare <m:oMath> in a display paragraph', () => {
		expect(genXmlMathParagraph(OMATH)).toBe(
			`<a:p><a14:m ${NS}><m:oMathPara>${PARA_PR}${OMATH}</m:oMathPara></a14:m><a:endParaRPr lang="en-US"/></a:p>`
		)
	})

	test('wraps inner-only OMML in <m:oMath> first', () => {
		expect(genXmlMathParagraph('<m:r><m:t>x</m:t></m:r>')).toBe(
			`<a:p><a14:m ${NS}><m:oMathPara>${PARA_PR}<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US"/></a:p>`
		)
	})

	test('passes an authored <m:oMathPara> through untouched — no second wrapper, no paraPr', () => {
		const para = `<m:oMathPara>${OMATH}</m:oMathPara>`
		expect(genXmlMathParagraph(para)).toBe(`<a:p><a14:m ${NS}>${para}</a14:m><a:endParaRPr lang="en-US"/></a:p>`)
	})

	test('trims surrounding whitespace before classifying the input', () => {
		expect(genXmlMathParagraph(`   ${OMATH}   `)).toBe(genXmlMathParagraph(OMATH))
	})

	test('empty input still yields a well-formed (empty) equation paragraph', () => {
		expect(genXmlMathParagraph('')).toBe(
			`<a:p><a14:m ${NS}><m:oMathPara>${PARA_PR}<m:oMath></m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US"/></a:p>`
		)
	})
})

describe('genXmlInlineMath()', () => {
	test('emits a bare <a14:m> run — no <m:oMathPara>, no paragraph, no endParaRPr', () => {
		expect(genXmlInlineMath(OMATH)).toBe(`<a14:m ${NS}>${OMATH}</a14:m>`)
	})

	test('unwraps a display <m:oMathPara>, since a paragraph block cannot flow inline', () => {
		expect(genXmlInlineMath(`<m:oMathPara>${PARA_PR}${OMATH}</m:oMathPara>`)).toBe(`<a14:m ${NS}>${OMATH}</a14:m>`)
	})

	test('wraps inner-only OMML in <m:oMath>', () => {
		expect(genXmlInlineMath('<m:r><m:t>x</m:t></m:r>')).toBe(
			`<a14:m ${NS}><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></a14:m>`
		)
	})

	test('a self-closed <m:oMath/> does not match the unwrap regex and is re-wrapped', () => {
		// Documents today's behavior rather than endorsing it: the regex requires a closing
		// tag, so `<m:oMath/>` falls through to the inner-OMML branch and nests.
		expect(genXmlInlineMath('<m:oMath/>')).toBe(`<a14:m ${NS}><m:oMath><m:oMath/></m:oMath></a14:m>`)
	})
})
