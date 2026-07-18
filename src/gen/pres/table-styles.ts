/**
 * PptxGenJS: `ppt/tableStyles.xml`
 *
 * Emit the table-styles part: the default style id plus any custom table styles,
 * each built from its `CT_TablePartStyle` regions (text style, cell borders/fill)
 * in the schema-required order.
 */

import { CRLF, TableStyle, XML_DECL } from '../../core-enums.js'
import type { BorderProps, TableStyleInternal, TableStyleRegionProps } from '../../core-interfaces.js'
import {
	createColorElement,
	encodeXmlEntities,
	genXmlColorSelection,
	lineWidthToEmu,
	resolveBorderWidth,
} from '../../gen-utils.js'

/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
export function makeXmlTableStyles(tableStyles: TableStyleInternal[] = []): string {
	const open = `${XML_DECL}${CRLF}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${TableStyle.MEDIUM_STYLE_2_ACCENT_1}"`
	if (!tableStyles || tableStyles.length === 0) return `${open}/>`

	let strXml = `${open}>`
	tableStyles.forEach(({ guid, def }) => {
		strXml += `<a:tblStyle styleId="${guid}" styleName="${encodeXmlEntities(def.name)}">`
		// NOTE: regions MUST be emitted in CT_TableStyle schema order or PowerPoint reports the file as corrupt
		;(
			[
				['wholeTbl', def.wholeTbl],
				['band1H', def.band1H],
				['band2H', def.band2H],
				['band1V', def.band1V],
				['band2V', def.band2V],
				['lastCol', def.lastCol],
				['firstCol', def.firstCol],
				['lastRow', def.lastRow],
				['firstRow', def.firstRow],
			] as const
		).forEach(([name, region]) => {
			if (region) strXml += genXmlTableStyleRegion(name, region)
		})
		strXml += '</a:tblStyle>'
	})
	strXml += '</a:tblStyleLst>'
	return strXml
}

/**
 * Build one `CT_TablePartStyle` region (e.g. `firstRow`, `band1H`) for a custom table style.
 * Emits `tcTxStyle` (text) before `tcStyle` (cell fill/borders) per the schema sequence.
 * @param {string} name - region element name
 * @param {TableStyleRegionProps} region - region styling
 * @return {string} XML
 */
function genXmlTableStyleRegion(name: string, region: TableStyleRegionProps): string {
	let xml = `<a:${name}>`

	// A: tcTxStyle — text style (only when text formatting is requested)
	if (region.bold !== undefined || region.italic !== undefined || region.color) {
		const b = region.bold ? ' b="on"' : ''
		const i = region.italic ? ' i="on"' : ''
		xml += `<a:tcTxStyle${b}${i}><a:fontRef idx="minor"/>`
		xml += region.color ? createColorElement(region.color) : ''
		xml += '</a:tcTxStyle>'
	}

	// B: tcStyle — cell style: tcBdr (borders) then fill, in schema order
	if (region.border !== undefined || region.fill !== undefined) {
		xml += '<a:tcStyle>'
		if (region.border !== undefined) xml += genXmlTableStyleBorders(region.border)
		if (region.fill !== undefined) xml += `<a:fill>${genXmlColorSelection(region.fill)}</a:fill>`
		xml += '</a:tcStyle>'
	}

	xml += `</a:${name}>`
	return xml
}

/**
 * Build the `tcBdr` border block for a custom table style region.
 * A single `BorderProps` styles all four sides plus the interior grid lines; a
 * TRBL array styles only the four outer sides. Sides are emitted in schema order.
 * @param {BorderProps | BorderProps[]} border - border definition
 * @return {string} XML
 */
function genXmlTableStyleBorders(border: BorderProps | BorderProps[]): string {
	// NOTE: order MUST be left,right,top,bottom,insideH,insideV (CT_TableCellBorderStyle sequence)
	let sides: Array<[string, BorderProps | undefined]>
	if (Array.isArray(border)) {
		const [top, right, bottom, left] = border // TRBL input order
		sides = [
			['left', left],
			['right', right],
			['top', top],
			['bottom', bottom],
		]
	} else {
		sides = [
			['left', border],
			['right', border],
			['top', border],
			['bottom', border],
			['insideH', border],
			['insideV', border],
		]
	}

	let xml = '<a:tcBdr>'
	sides.forEach(([side, b]) => {
		if (!b) return
		xml += `<a:${side}>`
		if (b.type === 'none') {
			xml += '<a:ln><a:noFill/></a:ln>'
		} else {
			xml += `<a:ln w="${lineWidthToEmu(resolveBorderWidth(b, 1))}" cap="flat" cmpd="sng" algn="ctr">`
			xml += genXmlColorSelection({ color: b.color ?? '666666', transparency: b.transparency })
			xml += `<a:prstDash val="${b.type === 'dash' ? 'sysDash' : 'solid'}"/>`
			xml += '</a:ln>'
		}
		xml += `</a:${side}>`
	})
	xml += '</a:tcBdr>'
	return xml
}
