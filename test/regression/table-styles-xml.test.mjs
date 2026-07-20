import { describe, expect, test } from 'vitest'
import { makeXmlTableStyles } from '../../src/gen/pres/table-styles.ts'

// Characterization tests for tableStyles.xml that the byte-identity harness CANNOT see — the demo
// deck defines no custom table styles, so `.tmp/byte-identity/baseline/ppt/tableStyles.xml` is
// always the empty self-closing `<a:tblStyleLst .../>` stub; `defineTableStyle()`'s whole payload
// (regions, borders, text/fill styling) carries ZERO baseline parts. These pin the byte-level
// details the migration to el()/voidEl() must preserve: schema-required region/attribute/side
// order, the tcTxStyle-before-tcStyle sequence, and styleName escaping.

const style = (guid, def) => ({ guid, def })

describe('makeXmlTableStyles', () => {
	test('no custom styles emits a self-closing tblStyleLst', () => {
		const xml = makeXmlTableStyles([])
		expect(xml).toContain(
			'<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'
		)
	})

	test('styleId/styleName attribute order, styleName is escaped', () => {
		const xml = makeXmlTableStyles([style('{GUID-1}', { name: 'Q&A <Brand>' })])
		expect(xml).toContain('<a:tblStyle styleId="{GUID-1}" styleName="Q&amp;A &lt;Brand&gt;">')
	})

	test('regions are emitted in CT_TableStyle schema order regardless of def property order', () => {
		const xml = makeXmlTableStyles([
			style('{GUID-1}', {
				name: 'Brand',
				firstRow: { fill: '111111' },
				wholeTbl: { fill: '222222' },
				band1H: { fill: '333333' },
			}),
		])
		const iWhole = xml.indexOf('<a:wholeTbl>')
		const iBand1H = xml.indexOf('<a:band1H>')
		const iFirstRow = xml.indexOf('<a:firstRow>')
		expect(iWhole).toBeGreaterThan(-1)
		expect(iWhole).toBeLessThan(iBand1H)
		expect(iBand1H).toBeLessThan(iFirstRow)
	})

	test('tcTxStyle (bold/italic/color) precedes tcStyle (border/fill)', () => {
		const xml = makeXmlTableStyles([
			style('{GUID-1}', {
				name: 'Brand',
				wholeTbl: { bold: true, italic: true, color: 'FFFFFF', fill: '1A2B3C' },
			}),
		])
		expect(xml).toContain(
			'<a:wholeTbl><a:tcTxStyle b="on" i="on"><a:fontRef idx="minor"/><a:srgbClr val="FFFFFF"/></a:tcTxStyle><a:tcStyle><a:fill>'
		)
	})

	test('bold/italic omitted when unset; tcTxStyle still emitted for color alone', () => {
		const xml = makeXmlTableStyles([style('{GUID-1}', { name: 'Brand', wholeTbl: { color: 'FFFFFF' } })])
		expect(xml).toContain('<a:tcTxStyle><a:fontRef idx="minor"/>')
	})

	test('no tcTxStyle/tcStyle when a region has neither text nor border/fill styling', () => {
		const xml = makeXmlTableStyles([style('{GUID-1}', { name: 'Brand', wholeTbl: {} })])
		expect(xml).toContain('<a:wholeTbl></a:wholeTbl>')
	})

	test('single BorderProps applies to all six tcBdr sides in schema order', () => {
		const xml = makeXmlTableStyles([
			style('{GUID-1}', { name: 'Brand', wholeTbl: { border: { type: 'solid', color: 'D9D9D9', width: 0.5 } } }),
		])
		const bdr = xml.slice(xml.indexOf('<a:tcBdr>'), xml.indexOf('</a:tcBdr>'))
		expect(
			['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map((s) => `<a:${s}>`).every((tag) => bdr.includes(tag))
		).toBe(true)
		const order = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV'].map((s) => bdr.indexOf(`<a:${s}>`))
		expect(order).toEqual([...order].sort((a, b) => a - b))
	})

	test('TRBL border array styles only the four outer sides, remapped left/right/top/bottom', () => {
		const top = { type: 'solid', color: '111111', width: 1 }
		const right = { type: 'solid', color: '222222', width: 1 }
		const bottom = { type: 'solid', color: '333333', width: 1 }
		const left = { type: 'solid', color: '444444', width: 1 }
		const xml = makeXmlTableStyles([
			style('{GUID-1}', { name: 'Brand', wholeTbl: { border: [top, right, bottom, left] } }),
		])
		const bdr = xml.slice(xml.indexOf('<a:tcBdr>'), xml.indexOf('</a:tcBdr>'))
		expect(bdr).not.toContain('insideH')
		expect(bdr).not.toContain('insideV')
		expect(bdr.indexOf('444444')).toBeLessThan(bdr.indexOf('222222'))
		expect(bdr.indexOf('222222')).toBeLessThan(bdr.indexOf('111111'))
		expect(bdr.indexOf('111111')).toBeLessThan(bdr.indexOf('333333'))
	})

	test('border type "none" emits noFill instead of a colored line', () => {
		const xml = makeXmlTableStyles([style('{GUID-1}', { name: 'Brand', wholeTbl: { border: { type: 'none' } } })])
		expect(xml).toContain('<a:left><a:ln><a:noFill/></a:ln></a:left>')
	})

	test('border line width/color/dash attributes', () => {
		const xml = makeXmlTableStyles([
			style('{GUID-1}', { name: 'Brand', wholeTbl: { border: { type: 'dash', color: 'ABCDEF', width: 1 } } }),
		])
		expect(xml).toContain('<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">')
		expect(xml).toContain('<a:srgbClr val="ABCDEF"/>')
		expect(xml).toContain('<a:prstDash val="sysDash"/>')
	})

	test('border color defaults to 666666 when unset', () => {
		const xml = makeXmlTableStyles([style('{GUID-1}', { name: 'Brand', wholeTbl: { border: { type: 'solid' } } })])
		expect(xml).toContain('<a:srgbClr val="666666"/>')
	})

	test('multiple styles are concatenated with no separator', () => {
		const xml = makeXmlTableStyles([style('{GUID-1}', { name: 'A' }), style('{GUID-2}', { name: 'B' })])
		expect(xml).toContain(
			'<a:tblStyle styleId="{GUID-1}" styleName="A"></a:tblStyle><a:tblStyle styleId="{GUID-2}" styleName="B"></a:tblStyle>'
		)
	})
})
