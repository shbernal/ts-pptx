import { describe, test, expect } from 'vitest'
import { DOMParser, onErrorStopParsing } from '@xmldom/xmldom'
import { latexToOmml, mathmlToOmml } from '../../../dist/math.js'

// `@shbernal/ts-pptx/math` converts LaTeX/MathML to OMML for the `math:` option on
// addText (upstream-issue-1456). Pipeline: LaTeX --temml--> MathML --mathml2omml--> OMML.
// These tests pin the canonical output shape, well-formedness, the display/inline
// distinction, and the throw-on-invalid-input policy. Schema validity of the emitted
// deck is covered separately by test/schema-cases.js.

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
		expect(() => latexToOmml('\\frac{')).toThrow(/Invalid LaTeX.*position/s)
		expect(() => latexToOmml('\\unknowncommand{x}')).toThrow(/Invalid LaTeX/)
	})
})

describe('math/latexToOmml — accents become m:acc, not m:limUpp', () => {
	// The pipeline's two halves disagree about how an accent is signalled: temml renders one
	// as a bare `<mover>` (correct for a browser, which derives accent positioning from the
	// MathML operator dictionary), while mathml2omml has no dictionary and keys strictly off
	// `accent="true"` — so every accent landed as `<m:limUpp>`, an over-*limit*, with limit
	// spacing and semantics. `markAccentedMovers` carries the dictionary subset that closes
	// it. These pin the whole table, because a missing entry is silent: the deck still opens.

	/** `[element kinds, m:chr code points]` for one inline conversion. */
	function convert(latex) {
		const omml = latexToOmml(latex, { display: false })
		assertWellFormed(omml, latex)
		return {
			kinds: [...omml.matchAll(/<m:(acc|limUpp|limLow|groupChr|borderBox)\b/g)].map((m) => m[1]),
			chars: [...omml.matchAll(/<m:chr m:val="([^"]*)"/g)].map((m) => m[1].codePointAt(0)),
		}
	}

	// Every accent command temml documents, with the combining mark ECMA-376 §22.1.2.20 says
	// an `accPr` character should be (U+0300–U+036F or U+20D0–U+20EF). temml emits the
	// *spacing* modifiers instead (U+02C6, U+2192, …) and mathml2omml passes them straight
	// through, so without the swap `\vec{v}` hangs a full-size arrow over the base.
	const ACCENTS = [
		['\\hat{a}', 0x0302],
		['\\^{a}', 0x0302],
		['\\tilde{n}', 0x0303],
		['\\~{n}', 0x0303],
		['\\acute{e}', 0x0301],
		["\\'{e}", 0x0301],
		['\\grave{e}', 0x0300],
		['\\`{e}', 0x0300],
		['\\ddot{y}', 0x0308],
		['\\"{y}', 0x0308],
		['\\dot{y}', 0x0307],
		['\\.{y}', 0x0307],
		['\\bar{x}', 0x0304],
		['\\={x}', 0x0304],
		['\\breve{u}', 0x0306],
		['\\u{u}', 0x0306],
		['\\check{s}', 0x030c],
		['\\v{s}', 0x030c],
		['\\mathring{A}', 0x030a],
		['\\r{A}', 0x030a],
		['\\H{o}', 0x030b],
		['\\vec{v}', 0x20d7],
		['\\dddot{y}', 0x20db],
	]

	for (const [latex, codePoint] of ACCENTS) {
		test(`${latex} is an accent carrying U+${codePoint.toString(16).toUpperCase().padStart(4, '0')}`, () => {
			const { kinds, chars } = convert(latex)
			expect(kinds, `${latex} is an m:acc, not an over-limit`).toEqual(['acc'])
			expect(chars).toEqual([codePoint])
		})
	}

	test('nested accents each carry their own mark', () => {
		// The rewrite pairs `<mover>` with `</mover>` on a stack, so an accent inside an accent
		// has to come out as two m:acc with different characters — a naive innermost-only or
		// outermost-only match would drop one of them back to m:limUpp.
		const { kinds, chars } = convert('\\vec{\\hat{n}}')
		expect(kinds).toEqual(['acc', 'acc'])
		expect(chars).toEqual([0x20d7, 0x0302])
	})

	test('what deliberately stays as it was', () => {
		// Each of these is already a better rendering than m:acc would be, so the rewrite is
		// scoped off them rather than sweeping every <mover> into an accent.
		expect(convert('\\widehat{abc}').kinds, 'a wide accent stretches — m:groupChr').toEqual(['groupChr'])
		expect(convert('\\overbrace{a+b}').kinds, 'a brace is a group character').toEqual(['groupChr'])
		expect(convert('\\overline{AB}').kinds, 'temml emits menclose, which maps to a rule').toEqual(['borderBox'])
		expect(convert('\\underline{y}').kinds, 'likewise below').toEqual(['borderBox'])
		expect(convert('\\stackrel{?}{=}').kinds, 'an over-relation really is a limit').toEqual(['limUpp'])
		// OMML has no under-accent object: `accentunder="true"` makes mathml2omml emit m:acc,
		// which would move the tilde ABOVE the base. m:limLow keeps it below.
		expect(convert('\\utilde{y}').kinds, 'an under-accent stays a low limit').toEqual(['limLow'])
		// "…." is two characters and m:chr takes one, so this keeps the limit form rather than
		// producing an m:chr no consumer can read.
		expect(convert('\\ddddot{y}').kinds, 'a four-dot accent has no single m:chr').toEqual(['limUpp'])
	})

	test('an extensible arrow is left alone — temml states accent="false" on it', () => {
		// The one place temml *does* answer the question the rewrite exists to answer, and it
		// answers "no": `\xrightarrow` is `<mover accent="false">` over an inner `<mover>` whose
		// top is an `<mspace>` rather than an `<mo>`. Both of the rewrite's skip conditions —
		// `accent` already stated, and an over-script that is not an operator — are load-bearing
		// here, so this is the case that keeps them honest.
		expect(convert('\\xrightarrow{f}').kinds, 'both movers stay limits').toEqual(['limUpp', 'limUpp'])
	})

	test('mathmlToOmml does not rewrite a caller-supplied <mover>', () => {
		// The shim compensates for something *temml* does. MathML gives a caller the `accent`
		// attribute to state this themselves, so second-guessing it on hand-written input would
		// take away the only way to ask for an over-limit.
		const bare = mathmlToOmml('<math><mover><mi>a</mi><mo stretchy="false">ˆ</mo></mover></math>')
		expect(bare.includes('<m:limUpp>'), 'a bare mover stays a limit').toBe(true)
		const stated = mathmlToOmml('<math><mover accent="true"><mi>a</mi><mo stretchy="false">ˆ</mo></mover></math>')
		expect(stated.includes('<m:acc>'), 'and a stated accent is honoured').toBe(true)
	})

	test('an accent already stating `accent` is left alone by latexToOmml too', () => {
		// Belt and braces on the guard in `markAccentedMovers`: a double rewrite would insert a
		// second attribute and produce malformed XML.
		const omml = latexToOmml('\\hat{a}', { display: false })
		expect((omml.match(/<m:acc>/g) || []).length).toBe(1)
		assertWellFormed(omml, 'single accent')
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
