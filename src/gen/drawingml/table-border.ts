/**
 * PptxGenJS: DrawingML table-cell borders
 *
 * Emit the `<a:lnL>/<a:lnR>/<a:lnT>/<a:lnB>` border children of a table cell's
 * `<a:tcPr>`, in the LRTB document order PowerPoint requires.
 */

import type { BorderProps } from '../../core-interfaces.js'
import { createLineCap, genXmlColorSelection, resolveBorderWidth, valToPts } from '../../gen-utils.js'

/**
 * Emit the `<a:lnL>/<a:lnR>/<a:lnT>/<a:lnB>` border children of an `<a:tcPr>` for a table cell.
 * Shared by normal cells and the dummy span (`_hmerge`/`_vmerge`) cells so a merged region's
 * outer edges render with the same border as its origin cell.
 * @param {BorderProps[]} cellBorder - 4-tuple of border props in [top, right, bottom, left] order
 * @return {string} concatenated border element XML, in the LRTB document order PowerPoint expects
 */
export function genTableCellBorderXml(cellBorder: BorderProps[]): string {
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
		const cap = createLineCap(border.cap)
		if (border.type !== 'none') {
			strXml += `<a:${obj.name} w="${valToPts(resolveBorderWidth(border, 1))}" cap="${cap}" cmpd="sng" algn="ctr">`
			strXml += genXmlColorSelection({ color: border.color ?? '363636', transparency: border.transparency })
			strXml += `<a:prstDash val="${
				border.type === 'dash' ? 'sysDash' : 'solid'
			}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/>`
			strXml += `</a:${obj.name}>`
		} else {
			strXml += `<a:${obj.name} w="0" cap="${cap}" cmpd="sng" algn="ctr"><a:noFill/></a:${obj.name}>`
		}
	})
	return strXml
}
