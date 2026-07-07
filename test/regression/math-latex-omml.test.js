import { describe, test, expect } from 'vitest'
import { DOMParser, onErrorStopParsing } from '@xmldom/xmldom'
import { latexToOmml, mathmlToOmml } from '../../dist/math.js'

// `@shbernal/pptxgenjs/math` converts LaTeX/MathML to OMML for the `math:` option on
// addText (upstream-issue-1456). Pipeline: LaTeX --temml--> MathML --mathml2omml--> OMML.
// These tests pin the canonical output shape, well-formedness, the display/inline
// distinction, and the throw-on-invalid-input policy. Schema validity of the emitted
// deck is covered separately by test/schema.test.js.

const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

// Parse OMML by declaring the `m` prefix (the `math:` envelope supplies it at runtime);
// @xmldom's onErrorStopParsing turns any well-formedness fault into a thrown error.
function assertWellFormed(omml, label) {
	let threw = null
	const doc = new DOMParser({ onError: onErrorStopParsing }).parseFromString(
		`<root xmlns:m="${M_NS}">${omml}</root>`,
		'text/xml'
	)
	try {
		// Touching the tree surfaces a parse error node if xmldom recovered instead of throwing.
		if (doc.getElementsByTagName('parsererror').length > 0) threw = 'parsererror node'
	} catch (e) {
		threw = e.message
	}
	expect(threw, `${label} should be well-formed XML`).toBeNull()
	return doc
}

// One formula per requested corpus family (plan Step 3).
const CORPUS = {
	fraction: 'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}',
	nestedRadical: '\\sqrt{1+\\sqrt{1+x}}',
	sumLimits: '\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2}',
	integral: '\\int_0^\\infty e^{-x}\\,dx = 1',
	matrix: '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}',
	cases: 'f(x) = \\begin{cases} 1 & x>0 \\\\ 0 & x\\le 0 \\end{cases}',
	greek: '\\alpha + \\beta = \\gamma',
	accents: '\\hat{a} + \\bar{b}',
	fences: '\\left( \\frac{a}{b} \\right)',
}

describe('math/latexToOmml', () => {
	for (const [name, latex] of Object.entries(CORPUS)) {
		test(`converts ${name} to well-formed display OMML`, () => {
			const omml = latexToOmml(latex)
			// Canonical display form: a full <m:oMathPara> centered paragraph, no ns decls.
			expect(omml.startsWith('<m:oMathPara>'), `${name} is an oMathPara`).toBe(true)
			expect(omml.endsWith('</m:oMathPara>'), `${name} closes oMathPara`).toBe(true)
			expect(/xmlns/.test(omml), `${name} carries no namespace declarations`).toBe(false)
			expect(omml.includes('<m:oMath>'), `${name} wraps an m:oMath`).toBe(true)
			assertWellFormed(omml, name)
		})
	}

	test('display flag toggles oMathPara wrapping', () => {
		const latex = 'a^2 + b^2 = c^2'
		const display = latexToOmml(latex, { display: true })
		const inline = latexToOmml(latex, { display: false })
		expect(display.startsWith('<m:oMathPara>')).toBe(true)
		// Inline is a bare <m:oMath> — no display paragraph wrapper.
		expect(inline.startsWith('<m:oMath>')).toBe(true)
		expect(inline.includes('oMathPara')).toBe(false)
		expect(inline.endsWith('</m:oMath>')).toBe(true)
	})

	test('default is display math', () => {
		const latex = 'x + 1'
		expect(latexToOmml(latex)).toBe(latexToOmml(latex, { display: true }))
	})

	test('throws on invalid LaTeX with parse position', () => {
		expect(() => latexToOmml('\\frac{')).toThrowError(/Invalid LaTeX.*position/s)
		expect(() => latexToOmml('\\unknowncommand{x}')).toThrowError(/Invalid LaTeX/)
	})
})

describe('math/mathmlToOmml', () => {
	test('converts MathML to a bare, namespace-free m:oMath', () => {
		const omml = mathmlToOmml('<math><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow></math>')
		expect(omml.startsWith('<m:oMath>')).toBe(true)
		expect(omml.endsWith('</m:oMath>')).toBe(true)
		expect(/xmlns/.test(omml)).toBe(false)
		assertWellFormed(omml, 'mathml a+b')
	})
})
