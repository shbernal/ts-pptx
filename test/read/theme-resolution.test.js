// Read-model coverage for src/read/oxml/theme.ts colour / theme-font / style-
// matrix resolution branches the fixture decks don't all carry:
//   - resolveColor across sysClr / schemeClr (clrMap + direct-slot) / phClr,
//   - resolveThemeFont across +mj/+mn and lt/ea/cs script slots,
//   - styleRefFill/styleRefLine with phClr substitution and the bgFillStyleLst
//     (idx >= 1000) branch.
// resolveColorElement is exported and tested directly; the theme-font and style-
// matrix legs go through the public Run.resolvedFontFace / Shape.resolvedFill /
// Shape.resolvedLine getters over a synthetic TextFrame / AutoShape carrying a
// hand-authored fontScheme / fmtScheme in its theme context.

import { DOMParser } from '@xmldom/xmldom'
import { describe, test } from 'vitest'
import { TextFrame, AutoShape, resolveColorElement } from '../../dist/read.js'
import { assertEqual } from '../helpers.js'

const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

function ctx(overrides = {}) {
	return {
		clrMap: new Map(),
		clrScheme: new Map(),
		fmtScheme: null,
		fontScheme: null,
		layoutRoot: null,
		masterRoot: null,
		...overrides,
	}
}

/** Parse a single DrawingML element (`<a:srgbClr/>`, `<a:fontScheme>…`, `<a:fmtScheme>…`). */
/** @returns {import('@xmldom/xmldom').Element} the wrapper's sole child — callers pass exactly one element. */
function drawingEl(xml) {
	return /** @type {import('@xmldom/xmldom').Element} */ (
		new DOMParser().parseFromString(`<a:w xmlns:a="${A_NS}">${xml}</a:w>`, 'text/xml').documentElement.firstChild
	)
}

const stubPart = () => ({ markDirty() {} })

/** The first run of a synthetic single-run TextFrame resolving against `flatten`. */
function runWith(rPrInner, flatten) {
	const xml = `<p:txBody xmlns:p="${P_NS}" xmlns:a="${A_NS}"><a:bodyPr/><a:p><a:r>${rPrInner}<a:t>x</a:t></a:r></a:p></p:txBody>`
	const txBody = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	return new TextFrame(txBody, /** @type {any} */ (stubPart()), flatten).paragraphs[0].runs[0]
}

/** An AutoShape over a hand-authored p:sp, resolving against `flatten`. */
function autoShape(spXml, flatten) {
	const xml = `<p:spTree xmlns:p="${P_NS}" xmlns:a="${A_NS}">${spXml}</p:spTree>`
	const spTree = new DOMParser().parseFromString(xml, 'text/xml').documentElement
	const el = spTree.getElementsByTagNameNS(P_NS, 'sp')[0]
	return new AutoShape(el, /** @type {any} */ ({ themeContext: () => flatten }))
}

describe('resolveColor — colour models', () => {
	test('sysClr resolves via @lastClr, falling back to @val', () => {
		assertEqual(
			resolveColorElement(drawingEl(`<a:sysClr val="windowText" lastClr="AABBCC"/>`), ctx()).hex,
			'AABBCC',
			'lastClr wins'
		)
		assertEqual(
			resolveColorElement(drawingEl(`<a:sysClr val="808080"/>`), ctx()).hex,
			'808080',
			'val is the fallback when no lastClr'
		)
	})

	test('schemeClr resolves through the clrMap → clrScheme indirection', () => {
		const c = ctx({ clrMap: new Map([['tx1', 'dk1']]), clrScheme: new Map([['dk1', '111111']]) })
		assertEqual(resolveColorElement(drawingEl(`<a:schemeClr val="tx1"/>`), c).hex, '111111', 'tx1 → dk1 → hex')
	})

	test('a direct-slot scheme token (dk1/lt1/dk2/lt2) bypasses the clrMap', () => {
		const c = ctx({ clrScheme: new Map([['lt2', 'EEEEEE']]) })
		assertEqual(resolveColorElement(drawingEl(`<a:schemeClr val="lt2"/>`), c).hex, 'EEEEEE', 'lt2 is a direct slot')
	})

	test('a phClr placeholder token resolves to null (it is filled in by a styleRef)', () => {
		assertEqual(
			resolveColorElement(drawingEl(`<a:schemeClr val="phClr"/>`), ctx()),
			null,
			'phClr is never a literal on its own'
		)
	})
})

describe('resolveThemeFont — +mj/+mn tokens across script slots', () => {
	const fontScheme = drawingEl(
		`<a:fontScheme>` +
			`<a:majorFont><a:latin typeface="Georgia"/><a:cs typeface="Major CS"/></a:majorFont>` +
			`<a:minorFont><a:latin typeface="Verdana"/><a:ea typeface="MS Mincho"/></a:minorFont>` +
			`</a:fontScheme>`
	)
	const face = (typeface, flatten) =>
		runWith(`<a:rPr><a:latin typeface="${typeface}"/></a:rPr>`, flatten).resolvedFontFace

	test('minor/ea and major/cs and minor/lt tokens resolve to their scheme faces', () => {
		const f = ctx({ fontScheme })
		assertEqual(face('+mn-ea', f), 'MS Mincho', '+mn-ea → minorFont/ea')
		assertEqual(face('+mj-cs', f), 'Major CS', '+mj-cs → majorFont/cs')
		assertEqual(face('+mn-lt', f), 'Verdana', '+mn-lt → minorFont/latin')
	})

	test('a literal (non-token) face is returned verbatim', () => {
		assertEqual(face('Rockwell', ctx({ fontScheme })), 'Rockwell', 'a plain face resolves to itself')
	})

	test('a token resolves to null with no fontScheme, or when the scheme slot is empty', () => {
		assertEqual(face('+mn-lt', ctx()), null, 'no fontScheme → null')
		// majorFont has no <a:ea>, so +mj-ea finds no face.
		assertEqual(face('+mj-ea', ctx({ fontScheme })), null, 'an empty script slot → null')
	})
})

describe('styleRef fill/line — phClr substitution + bgFillStyleLst', () => {
	test('a fillRef solid entry has its phClr replaced by the ref colour', () => {
		const fmtScheme = drawingEl(
			`<a:fmtScheme><a:fillStyleLst>` +
				`<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
				`</a:fillStyleLst><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>`
		)
		const shape = autoShape(
			`<p:sp><p:style><a:fillRef idx="1"><a:srgbClr val="FF0000"/></a:fillRef></p:style><p:spPr/></p:sp>`,
			ctx({ fmtScheme })
		)
		assertEqual(shape.resolvedFill.hex, 'FF0000', 'the ref colour fills in the phClr slot')
	})

	test('a fillRef idx >= 1000 selects the bgFillStyleLst', () => {
		const fmtScheme = drawingEl(
			`<a:fmtScheme><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/>` +
				`<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>` +
				`</a:fmtScheme>`
		)
		const shape = autoShape(
			`<p:sp><p:style><a:fillRef idx="1001"><a:srgbClr val="00FF00"/></a:fillRef></p:style><p:spPr/></p:sp>`,
			ctx({ fmtScheme })
		)
		assertEqual(shape.resolvedFill.hex, '00FF00', 'idx 1001 → bgFillStyleLst entry 1, phClr filled')
	})

	test('a lnRef resolves the style-matrix line colour', () => {
		const fmtScheme = drawingEl(
			`<a:fmtScheme><a:fillStyleLst/>` +
				`<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
				`<a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>`
		)
		const shape = autoShape(
			`<p:sp><p:style><a:lnRef idx="1"><a:srgbClr val="0000FF"/></a:lnRef></p:style><p:spPr/></p:sp>`,
			ctx({ fmtScheme })
		)
		assertEqual(shape.resolvedLine.hex, '0000FF', 'the lnRef colour fills the phClr in the lnStyleLst entry')
	})
})
