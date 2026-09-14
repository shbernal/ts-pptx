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
import { Relationships, TextFrame } from '../../dist/read.js'
import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

const stubPart = () => ({ markDirty() {} })

/** A theme with no colours, fonts or roots: a scheme token resolves to nothing, a literal to itself. */
const emptyTheme = () => ({ clrMap: new Map(), clrScheme: new Map(), fmtScheme: null })

/**
 * The context a hand-built frame is read against: a part (a stub absorbs markDirty), the empty
 * theme, the given relationships, and no inheritance.
 * @param {any} [part]
 * @param {any} [rels]
 */
const bareText = (part = stubPart(), rels = null) => ({ part, ctx: emptyTheme(), rels, inherit: null })

/** A TextFrame over hand-authored p:txBody inner XML. */
function frame(inner) {
	const xml = `<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr/>${inner}</p:txBody>`
	const txBody = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	return new TextFrame(txBody, bareText())
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

	test('underline and schemeColor refuse a token outside the schema, leaving the run as it was', () => {
		const r = run(
			`<a:r><a:rPr u="sng"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>x</a:t></a:r>`
		)
		const codeOf = (/** @type {() => void} */ fn) => {
			try {
				fn()
				return null
			} catch (err) {
				return /** @type {any} */ (err).code
			}
		}
		assertEqual(
			codeOf(() => {
				r.underline = 'wavy-nonsense'
			}),
			'text/invalid-underline',
			'an unknown underline token'
		)
		assertEqual(
			codeOf(() => {
				r.schemeColor = 'bogus'
			}),
			'color/invalid-scheme-token',
			'an unknown theme colour token'
		)
		assertEqual(r.underline, 'sng', 'the underline is unchanged')
		assertEqual(r.schemeColor, 'accent1', 'the colour is unchanged')
	})

	test('clearing a colour or font the run does not have leaves the part clean', () => {
		// The run's clears marked the part dirty whenever it had an `a:rPr`, where the shape and table
		// cell clears marked it only when something was removed.
		let marks = 0
		const part = {
			markDirty() {
				marks++
			},
		}
		const xml = `<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr/><a:p><a:r><a:rPr b="1"/><a:t>x</a:t></a:r></a:p></p:txBody>`
		const txBody = new DOMParser().parseFromString(xml, 'text/xml').documentElement
		const r = new TextFrame(txBody, bareText(part)).paragraphs[0].runs[0]
		r.color = null
		r.schemeColor = null
		r.fontName = null
		assertEqual(marks, 0, 'nothing changed, so nothing is marked')
		r.color = 'FF0000'
		assertEqual(marks, 1, 'a colour that is written marks the part')
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

describe('Run.resolvedColor with a fill of its own that is not solid', () => {
	test('reports no colour rather than the inherited one', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addText('run', { x: 1, y: 1, w: 3, h: 1, color: '336699' })
		})
		const run = presentation.slides[0].shapes[0].textFrame.paragraphs[0].runs[0]
		const rPr = run.element_.getElementsByTagNameNS(A_NS, 'rPr')[0]
		for (const child of Array.from(rPr.childNodes)) {
			if (child.nodeName === 'a:solidFill') rPr.removeChild(child)
		}
		assert(run.resolvedColor !== null, 'with no fill of its own the run inherits a colour')
		rPr.insertBefore(rPr.ownerDocument.createElementNS(A_NS, 'a:noFill'), rPr.firstChild)
		assertEqual(run.resolvedColor, null, 'with a:noFill it reports none')
	})
})

describe('Run getter edges on text that inherits through nothing', () => {
	test('text is empty when the run has no a:t', () => {
		assertEqual(run(`<a:r><a:rPr/></a:r>`).text, '', 'no a:t → empty text')
	})

	test('resolvedColor is null and resolvedFontFace falls back to the literal own face', () => {
		assertEqual(run(`<a:r><a:t>x</a:t></a:r>`).resolvedColor, null, 'no own fill and no inheritance → null')
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

	test('a picture bullet whose embed names no internal part reads null rather than throwing', () => {
		// A picture already reports a linked or missing image as `null`. The bullet resolved its embed
		// straight through `resolveTarget`, so the same embed threw out of `bulletDetail`. No
		// PowerPoint fixture carries a linked bullet image, so the relationships are built here.
		const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
		const IMAGE_REL = `${R_NS}/image`
		const rels = Relationships.empty('/ppt/slides/slide1.xml')
		rels.addWithId('rId9', IMAGE_REL, 'https://example.invalid/bullet.png', 'External')
		rels.addWithId('rId10', IMAGE_REL, '../media/image1.png')

		/** The picture bullet of a paragraph whose `a:buBlip` embeds `relId`, read against `rels`. */
		const bulletImage = (relId) => {
			const xml =
				`<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"><a:bodyPr/>` +
				`<a:p><a:pPr><a:buBlip><a:blip r:embed="${relId}"/></a:buBlip></a:pPr><a:r><a:t>x</a:t></a:r></a:p></p:txBody>`
			const txBody = /** @type {any} */ (new DOMParser().parseFromString(xml, 'text/xml').documentElement)
			const bullet = new TextFrame(txBody, bareText(stubPart(), rels)).paragraphs[0].bulletDetail
			if (bullet?.kind !== 'picture') throw new Error(`expected a picture bullet, got ${JSON.stringify(bullet)}`)
			return bullet.imagePartName
		}

		assertEqual(bulletImage('rId8'), null, 'an embed naming no relationship reads null')
		assertEqual(bulletImage('rId9'), null, 'a linked (External) image reads null')
		assertEqual(bulletImage('rId10'), '/ppt/media/image1.png', 'an embedded image still resolves to its part')
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
		// A literal colour resolves to itself, even against a theme that maps nothing.
		assertEqual(bullet.resolvedColor?.effectiveHex, 'C00000', 'a:buClr literal resolved')
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
		const f = new TextFrame(withAnchor, bareText())
		assertEqual(f.resolvedAnchor, 'ctr', 'own bodyPr @anchor wins')
		assertEqual(frame(`<a:p/>`).resolvedAnchor, null, 'no own anchor and no placeholder → null')
	})

	test('element_ exposes the p:txBody element', () => {
		assertEqual(frame(`<a:p/>`).element_.localName, 'txBody', 'element_ is the p:txBody')
	})
})
