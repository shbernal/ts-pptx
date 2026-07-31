/**
 * ts-pptx: DrawingML table-cell borders
 *
 * Emit the `<a:lnL>/<a:lnR>/<a:lnT>/<a:lnB>` border children of a table cell's
 * `<a:tcPr>`, in the LRTB document order PowerPoint requires, plus the two
 * optional diagonals (`<a:lnTlToBr>`/`<a:lnBlToTr>`) that follow them.
 */

import type { BorderProps, TableCellDiagonals } from '../../types/index.js'
import { genXmlColorSelection } from './fill.js'
import { createLineCap, resolveBorderDash, resolveBorderWidth } from './line.js'
import { valToPts } from '../../units-internal.js'
import { el, raw, voidEl } from '../oxml/el.js'

/**
 * One `<a:lnX>` border element. Every side of a table cell — the four edges and the two
 * diagonals — is a `CT_LineProperties`, so they differ only in element name.
 * @param {string} name - the element's local name (`lnL`, `lnTlToBr`, …)
 * @param {BorderProps} border - the resolved border for that side
 * @return {string} the border element XML
 */
function genBorderLine(name: string, border: BorderProps): string {
	const cap = createLineCap(border.cap)
	if (border.type === 'none') return el(`a:${name}`, { w: 0, cap, cmpd: 'sng', algn: 'ctr' }, raw(voidEl('a:noFill')))
	return el(
		`a:${name}`,
		{ w: valToPts(resolveBorderWidth(border, 1)), cap, cmpd: 'sng', algn: 'ctr' },
		[
			genXmlColorSelection({ color: border.color ?? '363636', transparency: border.transparency }),
			voidEl('a:prstDash', { val: resolveBorderDash(border) }),
			voidEl('a:round'),
			voidEl('a:headEnd', { type: 'none', w: 'med', len: 'med' }),
			voidEl('a:tailEnd', { type: 'none', w: 'med', len: 'med' }),
		].map(raw)
	)
}

/**
 * Emit the border children of an `<a:tcPr>` for a table cell: the four edges, then the two
 * diagonals when the cell asks for them.
 *
 * The edges are shared by normal cells and the dummy span (`_hmerge`/`_vmerge`) cells so a
 * merged region's outer edges render with the same border as its origin cell. The diagonals
 * are **not**: a diagonal across a merged region is one stroke corner-to-corner, and
 * repeating it on each covered cell would draw a sawtooth instead. Only the span origin
 * carries them, which is also where PowerPoint puts them.
 *
 * @param {BorderProps[]} cellBorder - 4-tuple of border props in [top, right, bottom, left] order
 * @param {TableCellDiagonals} [diagonal] - the cell's optional corner-to-corner rules
 * @return {string} concatenated border element XML, in the document order PowerPoint expects
 */
export function genTableCellBorderXml(cellBorder: BorderProps[], diagonal?: TableCellDiagonals): string {
	let strXml = ''
	// NOTE: *** IMPORTANT! *** LRTB order matters! (Reorder a line below to watch the borders go wonky in MS-PPT-2013!!)
	;(
		[
			{ idx: 3, name: 'lnL' },
			{ idx: 1, name: 'lnR' },
			{ idx: 0, name: 'lnT' },
			{ idx: 2, name: 'lnB' },
		] as const
	).forEach((obj) => {
		const border = cellBorder[obj.idx]
		if (!border) return
		strXml += genBorderLine(obj.name, border)
	})
	// CT_TableCellProperties declares the diagonals immediately after `lnB` and before
	// `cell3D`/the fill, so they append here rather than being emitted by the caller.
	if (diagonal?.tlToBr) strXml += genBorderLine('lnTlToBr', diagonal.tlToBr)
	if (diagonal?.blToTr) strXml += genBorderLine('lnBlToTr', diagonal.blToTr)
	return strXml
}
