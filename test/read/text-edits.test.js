// Read-model coverage for the Run / Paragraph / TextFrame edit setters and the
// getter edges in src/read/api/text.ts that the fixture + edit suites don't hit:
// the underline / fontName / colour clear paths, the bullet buChar / buAutoNum
// branches, paragraph text with a:br / a:fld, TextFrame.text collapse / carry /
// whitespace, resolvedAnchor, and the element_ escape hatches.
//
// A synthetic TextFrame over hand-authored p:txBody XML drives all of these off
// -fixture; a stub part with a no-op markDirty lets the setters run without a
// real package. Setting a value then reading it back exercises each setter with
// its own getter — the read-model's own edit contract, not a writer round-trip.

import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { TextFrame } from '../../dist/read.js'
import { assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

const stubPart = () => ({ markDirty() {} })

/** A TextFrame over hand-authored p:txBody inner XML (a stub part absorbs markDirty). */
function frame(inner) {
	const xml = `<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr/>${inner}</p:txBody>`
	const txBody = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	return new TextFrame(txBody, /** @type {any} */ (stubPart()))
}

/** The first run of a single-paragraph frame. */
function run(runInner) {
	return frame(`<a:p>${runInner}</a:p>`).paragraphs[0].runs[0]
}

describe('Run character-property setters', () => {
	test('underline: get / set a token / clear', () => {
		const r = run(`<a:r><a:rPr u="sng"/><a:t>x</a:t></a:r>`)
		assertEqual(r.underline, 'sng', 'reads the u token')
		r.underline = 'dbl'
		assertEqual(r.underline, 'dbl', 'sets a new u token')
		r.underline = null
		assertEqual(r.underline, null, 'clearing removes @u')
	})

	test('fontName clear removes the a:latin child', () => {
		const r = run(`<a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>x</a:t></a:r>`)
		assertEqual(r.fontName, 'Arial', 'reads the latin typeface')
		r.fontName = null
		assertEqual(r.fontName, null, 'clearing removes a:latin')
	})

	test('colour: clear an explicit fill; set a scheme colour', () => {
		const r = run(`<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>x</a:t></a:r>`)
		assertEqual(r.color, 'FF0000', 'reads the srgb fill')
		r.color = null
		assertEqual(r.color, null, 'clearing removes the solidFill')
		r.schemeColor = 'accent1'
		assertEqual(r.schemeColor, 'accent1', 'sets a scheme-colour fill')
	})

	test('setting a bool attr null on a run with no rPr is a no-op', () => {
		const r = run(`<a:r><a:t>x</a:t></a:r>`)
		r.bold = null // #removeRPrAttr with no rPr → early return
		assertEqual(r.bold, null, 'stays null')
		r.color = null // #setSolidFill(null) with no rPr → early return
		assertEqual(r.color, null, 'stays null')
	})

	test('italic round-trips through the shared bool setter', () => {
		const r = run(`<a:r><a:t>x</a:t></a:r>`)
		r.italic = true
		assertEqual(r.italic, true, 'italic true')
		r.italic = false
		assertEqual(r.italic, false, 'italic false is written explicitly')
	})
})

describe('Run getter edges without a theme context', () => {
	test('text is empty when the run has no a:t', () => {
		assertEqual(run(`<a:r><a:rPr/></a:r>`).text, '', 'no a:t → empty text')
	})

	test('resolvedColor is null and resolvedFontFace falls back to the literal own face', () => {
		assertEqual(run(`<a:r><a:t>x</a:t></a:r>`).resolvedColor, null, 'no themeContext → null resolvedColor')
		const r = run(`<a:r><a:rPr><a:latin typeface="Georgia"/></a:rPr><a:t>x</a:t></a:r>`)
		assertEqual(r.resolvedFontFace, 'Georgia', 'own literal face resolves without a fontScheme')
	})

	test('element_ exposes the a:r element', () => {
		assertEqual(run(`<a:r><a:t>x</a:t></a:r>`).element_.localName, 'r', 'element_ is the a:r')
	})
})

describe('Paragraph getter edges', () => {
	/**
	 * The bullet of a single-paragraph frame, narrowed to the expected kind.
	 * `assert` is not an assertion signature, so a plain check would leave the
	 * union unnarrowed and every field access a type error.
	 * @template {import('../../dist/read.js').BulletDetail['kind']} K
	 * @param {string} pPr - the `a:pPr` inner XML
	 * @param {K} kind - the kind the bullet is expected to have
	 * @returns {Extract<import('../../dist/read.js').BulletDetail, { kind: K }>}
	 */
	function bulletOf(pPr, kind) {
		const bullet = frame(`<a:p><a:pPr>${pPr}</a:pPr><a:r><a:t>x</a:t></a:r></a:p>`).paragraphs[0].bulletDetail
		if (bullet?.kind !== kind) throw new Error(`expected a ${kind} bullet, got ${JSON.stringify(bullet)}`)
		return /** @type {Extract<import('../../dist/read.js').BulletDetail, { kind: K }>} */ (bullet)
	}

	test('bulletDetail distinguishes buNone, buChar, buAutoNum and buBlip', () => {
		bulletOf(`<a:buNone/>`, 'none')

		// The bare glyph, not `char:•` — a colon glyph is unambiguous in this form.
		assertEqual(bulletOf(`<a:buChar char="•"/>`, 'char').char, '•', 'a:buChar/@char verbatim')

		const buAuto = bulletOf(`<a:buAutoNum type="arabicPeriod"/>`, 'autoNum')
		assertEqual(buAuto.scheme, 'arabicPeriod', 'a:buAutoNum/@type')
		assertEqual(buAuto.startAt, null, 'no @startAt reads null, not a defaulted 1')

		// No relationships were threaded to this synthetic frame, so the part is unresolvable.
		const buBlip = bulletOf(`<a:buBlip><a:blip/></a:buBlip>`, 'picture')
		assertEqual(buBlip.imagePartName, null, 'unresolvable embed reads null')

		const bare = frame(`<a:p><a:r><a:t>x</a:t></a:r></a:p>`).paragraphs[0]
		assertEqual(bare.bulletDetail, null, 'no a:pPr at all → inherited, reported as null')
	})

	test('bulletDetail reads startAt and the bullet own font / size / colour', () => {
		const bullet = bulletOf(
			`<a:buClr><a:srgbClr val="C00000"/></a:buClr>` +
				`<a:buSzPct val="80000"/>` +
				`<a:buFont typeface="Wingdings"/>` +
				`<a:buAutoNum type="arabicPeriod" startAt="5"/>`,
			'autoNum'
		)
		assertEqual(bullet.startAt, 5, 'a:buAutoNum/@startAt')
		assertEqual(bullet.font, 'Wingdings', 'a:buFont/@typeface')
		// @val is thousandths of a percent: 80000 → 80%.
		assertEqual(bullet.sizePct, 80, 'a:buSzPct/@val as a percentage')
		assertEqual(bullet.sizePt, null, 'no a:buSzPts')
		assertEqual(bullet.color, 'C00000', 'a:buClr/a:srgbClr/@val')
		assertEqual(bullet.schemeColor, null, 'not a scheme colour')
		// No theme context on a synthetic frame, so the resolved form is unavailable.
		assertEqual(bullet.resolvedColor, null, 'no theme context → null resolvedColor')
	})

	test('bulletDetail reads a scheme bullet colour and an absolute buSzPts', () => {
		const bullet = bulletOf(
			`<a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buSzPts val="1400"/><a:buChar char="▪"/>`,
			'char'
		)
		assertEqual(bullet.schemeColor, 'accent2', 'a:buClr/a:schemeClr/@val')
		assertEqual(bullet.color, null, 'not an srgb colour')
		// @val is hundredths of a point: 1400 → 14pt.
		assertEqual(bullet.sizePt, 14, 'a:buSzPts/@val as points')
		assertEqual(bullet.sizePct, null, 'no a:buSzPct')
	})

	test('text concatenates runs and fields and renders a:br as a newline', () => {
		const p = frame(`<a:p><a:r><a:t>A</a:t></a:r><a:br/><a:fld><a:t>B</a:t></a:fld></a:p>`).paragraphs[0]
		assertEqual(p.text, 'A\nB', 'run + break + field text in order')
	})

	test('element_ exposes the a:p element', () => {
		assertEqual(
			frame(`<a:p><a:r><a:t>x</a:t></a:r></a:p>`).paragraphs[0].element_.localName,
			'p',
			'element_ is the a:p'
		)
	})
})

describe('TextFrame.text setter + resolvedAnchor + element_', () => {
	test('setting text collapses multiple paragraphs and carries the first run rPr', () => {
		const f = frame(`<a:p><a:r><a:rPr b="1"/><a:t>first</a:t></a:r></a:p><a:p><a:r><a:t>second</a:t></a:r></a:p>`)
		f.text = 'merged'
		assertEqual(f.paragraphs.length, 1, 'collapsed to a single paragraph')
		assertEqual(f.paragraphs[0].text, 'merged', 'new text is written')
		assertEqual(f.paragraphs[0].runs[0].bold, true, 'the first run rPr (bold) is carried onto the new run')
	})

	test('setting text on an empty body creates the paragraph and preserves whitespace', () => {
		const f = frame(``) // just <a:bodyPr/>, no a:p
		f.text = '  spaced  '
		assertEqual(f.paragraphs[0].runs[0].text, '  spaced  ', 'a fresh paragraph carries the padded text verbatim')
	})

	test('resolvedAnchor returns the own bodyPr anchor, else null with no placeholder', () => {
		const withAnchor = new DOMParser().parseFromString(
			`<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr anchor="ctr"/><a:p/></p:txBody>`,
			'text/xml'
		).documentElement
		const f = new TextFrame(withAnchor, /** @type {any} */ (stubPart()))
		assertEqual(f.resolvedAnchor, 'ctr', 'own bodyPr @anchor wins')
		assertEqual(frame(`<a:p/>`).resolvedAnchor, null, 'no own anchor and no placeholder → null')
	})

	test('element_ exposes the p:txBody element', () => {
		assertEqual(frame(`<a:p/>`).element_.localName, 'txBody', 'element_ is the p:txBody')
	})
})
