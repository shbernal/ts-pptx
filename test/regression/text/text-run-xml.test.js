import { describe, expect, test } from 'vitest'
import {
	genXmlNormAutofit,
	genXmlParagraphProperties,
	genXmlTextRunProperties,
} from '../../../src/gen/drawingml/text-run.ts'

// Characterization tests for text-run XML that the byte-identity harness CANNOT see: the demo
// deck emits zero parts containing `rtl="1"`, `<a:buClr>`, `<a:buBlip>` or `altLang`, so a green
// gate says nothing about them. These pin the exact bytes so the next refactor is not blind.
//
// Pinning is not endorsement — where the current output is quirky it is called out as such below.

const pPr = (options, isDefault = false) => genXmlParagraphProperties({ options }, isDefault)
const rPr = (options, isDefault = false) => genXmlTextRunProperties(options, isDefault)

describe('paragraph properties: right-to-left', () => {
	test('rtlMode alone emits a stray space before the closing bracket', () => {
		// QUIRK, pinned deliberately: the `rtl="1" ` fragment carries a trailing space.
		expect(pPr({ rtlMode: true, bullet: true })).toContain('<a:pPr rtl="1" ')
	})

	test('rtlMode followed by another attribute emits a DOUBLE space between them', () => {
		// QUIRK, pinned deliberately: `rtl="1" ` ends with a space and ` algn=` starts with one.
		// This is why the `<a:pPr>` open tag is not built with the element builder — `openTag`
		// joins attributes with exactly one space and cannot reproduce this.
		const xml = pPr({ rtlMode: true, align: 'center', bullet: true })
		expect(xml).toContain('rtl="1"  algn="ctr"')
	})
})

describe('paragraph properties: bullets', () => {
	test('bullet color emits buClr before the bullet glyph', () => {
		expect(pPr({ bullet: { color: 'FF00FF' } })).toBe(
			'<a:pPr marL="342900" indent="-342900">' +
				'<a:buClr><a:srgbClr val="FF00FF"/></a:buClr>' +
				'<a:buChar char="&#x2022;"/>' +
				'</a:pPr>'
		)
	})

	test('bullet char is a numeric character reference, NOT an escaped literal', () => {
		// `&#x2022;` must reach the file verbatim; `&amp;#x2022;` would render the text "&#x2022;".
		const xml = pPr({ bullet: true })
		expect(xml).toContain('<a:buChar char="&#x2022;"/>')
		expect(xml).not.toContain('&amp;#x2022;')
	})

	test('a picture bullet references the media rel', () => {
		expect(pPr({ bullet: { image: { data: 'x' }, _rId: 9 } })).toBe(
			'<a:pPr marL="342900" indent="-342900">' + '<a:buBlip><a:blip r:embed="rId9"/></a:buBlip>' + '</a:pPr>'
		)
	})

	test('an SVG picture bullet carries the svgBlip extension alongside the PNG preview', () => {
		expect(pPr({ bullet: { image: { data: 'x' }, _rId: 9, _rIdSvg: 10 } })).toBe(
			'<a:pPr marL="342900" indent="-342900">' +
				'<a:buBlip><a:blip r:embed="rId9">' +
				'<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
				'<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId10"/>' +
				'</a:ext></a:extLst></a:blip></a:buBlip>' +
				'</a:pPr>'
		)
	})

	test('a numbered bullet falls back to the major-latin theme font', () => {
		expect(pPr({ bullet: { type: 'number' } })).toContain(
			'<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod" startAt="1"/>'
		)
	})

	test('bullet fontFace is escaped', () => {
		expect(pPr({ bullet: { fontFace: 'A&B "C"' } })).toContain('<a:buFont typeface="A&amp;B &quot;C&quot;"/>')
	})
})

describe('run properties', () => {
	test('an explicit lang adds altLang', () => {
		expect(rPr({ lang: 'de-DE' })).toBe('<a:rPr lang="de-DE" altLang="en-US" dirty="0"></a:rPr>')
	})

	test('no lang means no altLang', () => {
		expect(rPr({})).toBe('<a:rPr lang="en-US" dirty="0"></a:rPr>')
	})

	test('isDefault switches the tag to a:defRPr', () => {
		expect(rPr({}, true)).toBe('<a:defRPr lang="en-US" dirty="0"></a:defRPr>')
	})

	test('charSpacing also disables kerning', () => {
		expect(rPr({ charSpacing: 3 })).toContain('spc="300" kern="0"')
	})

	test('font names are escaped in every slot', () => {
		const xml = rPr({ fontFace: 'A&B', fontFaceEA: '<EA>' })
		expect(xml).toContain('<a:latin typeface="A&amp;B" pitchFamily="34" charset="0"/>')
		expect(xml).toContain('<a:ea typeface="&lt;EA&gt;"/>')
		expect(xml).toContain('<a:cs typeface="A&amp;B"/>')
	})

	test('a plain hyperlink self-closes; tooltip is present but empty', () => {
		expect(rPr({ hyperlink: { url: 'https://x.com', _rId: 4 } })).toBe(
			'<a:rPr lang="en-US" u="sng" dirty="0">' +
				'<a:hlinkClick r:id="rId4" invalidUrl="" action="" tgtFrame="" tooltip="" history="1" highlightClick="0" endSnd="0"/>' +
				'</a:rPr>'
		)
	})

	test('a colored hyperlink becomes a paired element carrying the follow-text-color extension', () => {
		// The leading spaces inside the extension are byte-significant historical indentation.
		const xml = rPr({ hyperlink: { url: 'https://x.com', _rId: 4 }, color: 'FF0000' })
		expect(xml).toContain(
			'> <a:extLst>  <a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}">' +
				'   <ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/>' +
				'  </a:ext> </a:extLst></a:hlinkClick>'
		)
	})

	test('hyperlink tooltips are escaped', () => {
		expect(rPr({ hyperlink: { slide: 2, _rId: 7, tooltip: 'A & B' } })).toContain('tooltip="A &amp; B"')
	})
})

describe('shrink autofit', () => {
	test('omits attributes that were not supplied', () => {
		expect(genXmlNormAutofit({ type: 'shrink' })).toBe('<a:normAutofit/>')
	})

	test('converts percents to thousandths of a percent', () => {
		expect(genXmlNormAutofit({ type: 'shrink', fontScale: 92.5, lnSpcReduction: 10 })).toBe(
			'<a:normAutofit fontScale="92500" lnSpcReduction="10000"/>'
		)
	})

	test('a zero percent is emitted, not treated as absent', () => {
		expect(genXmlNormAutofit({ type: 'shrink', fontScale: 0 })).toBe('<a:normAutofit fontScale="0"/>')
	})

	test('an out-of-range value is clamped to the nearest bound, not dropped', () => {
		// The shared percentage policy: a finite value has a nearest legal neighbour, so it
		// moves there and warns. Dropping the attribute left the shrink un-parameterised,
		// which is a discarded request reported as a warning.
		expect(genXmlNormAutofit({ type: 'shrink', fontScale: 500, lnSpcReduction: 20 })).toBe(
			'<a:normAutofit fontScale="100000" lnSpcReduction="20000"/>'
		)
	})

	test('a value that is not a number throws rather than emitting val="NaN"', () => {
		expect(() => genXmlNormAutofit({ type: 'shrink', fontScale: Number.NaN })).toThrow(/fit.fontScale/)
	})
})
