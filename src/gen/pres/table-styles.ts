/**
 * PptxGenJS: `ppt/tableStyles.xml`
 *
 * Emit the table-styles part: the default style id plus any custom table styles,
 * each built from its `CT_TablePartStyle` regions (text style, cell borders/fill)
 * in the schema-required order.
 */

import { CRLF, TableStyle, XML_DECL } from '../../core-enums.js'
import type { BorderProps, TableStyleInternal, TableStyleRegionProps } from '../../core-interfaces.js'
import { createColorElement } from '../drawingml/color.js'
import { genXmlColorSelection } from '../drawingml/fill.js'
import { resolveBorderWidth } from '../drawingml/line.js'
import { lineWidthToEmu } from '../../units-internal.js'
import { el, raw, voidEl } from '../oxml/el.js'

/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
export function makeXmlTableStyles(tableStyles: TableStyleInternal[] = []): string {
	const attrs = {
		'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
		def: TableStyle.MEDIUM_STYLE_2_ACCENT_1,
	}
	if (!tableStyles || tableStyles.length === 0) return XML_DECL + CRLF + voidEl('a:tblStyleLst', attrs)

	return (
		XML_DECL +
		CRLF +
		el(
			'a:tblStyleLst',
			attrs,
			tableStyles.map(({ guid, def }) =>
				raw(
					el(
						'a:tblStyle',
						{ styleId: guid, styleName: def.name },
						// NOTE: regions MUST be emitted in CT_TableStyle schema order or PowerPoint reports the file as corrupt
						(
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
						).map(([name, region]) => (region ? raw(genXmlTableStyleRegion(name, region)) : null))
					)
				)
			)
		)
	)
}

/**
 * Build one `CT_TablePartStyle` region (e.g. `firstRow`, `band1H`) for a custom table style.
 * Emits `tcTxStyle` (text) before `tcStyle` (cell fill/borders) per the schema sequence.
 * @param {string} name - region element name
 * @param {TableStyleRegionProps} region - region styling
 * @return {string} XML
 */
function genXmlTableStyleRegion(name: string, region: TableStyleRegionProps): string {
	const children = [
		// A: tcTxStyle — text style (only when text formatting is requested)
		region.bold !== undefined || region.italic !== undefined || region.color
			? raw(
					el('a:tcTxStyle', { b: region.bold ? 'on' : null, i: region.italic ? 'on' : null }, [
						raw(voidEl('a:fontRef', { idx: 'minor' })),
						region.color ? raw(createColorElement(region.color)) : null,
					])
				)
			: null,
		// B: tcStyle — cell style: tcBdr (borders) then fill, in schema order
		region.border !== undefined || region.fill !== undefined
			? raw(
					el('a:tcStyle', null, [
						region.border !== undefined ? raw(genXmlTableStyleBorders(region.border)) : null,
						region.fill !== undefined ? raw(el('a:fill', null, raw(genXmlColorSelection(region.fill)))) : null,
					])
				)
			: null,
	]

	return el(`a:${name}`, null, children)
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

	return el(
		'a:tcBdr',
		null,
		sides.map(([side, b]) => {
			if (!b) return null
			if (b.type === 'none') return raw(el(`a:${side}`, null, raw(el('a:ln', null, raw(voidEl('a:noFill'))))))
			return raw(
				el(
					`a:${side}`,
					null,
					raw(
						el('a:ln', { w: lineWidthToEmu(resolveBorderWidth(b, 1)), cap: 'flat', cmpd: 'sng', algn: 'ctr' }, [
							raw(genXmlColorSelection({ color: b.color ?? '666666', transparency: b.transparency })),
							raw(voidEl('a:prstDash', { val: b.type === 'dash' ? 'sysDash' : 'solid' })),
						])
					)
				)
			)
		})
	)
}
