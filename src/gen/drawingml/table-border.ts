/**
 * ts-pptx: DrawingML table-cell borders
 *
 * Emit the `<a:lnL>/<a:lnR>/<a:lnT>/<a:lnB>` border children of a table cell's
 * `<a:tcPr>`, in the LRTB document order PowerPoint requires.
 */

import type { BorderProps } from '../../core-interfaces.js'
import { genXmlColorSelection } from './fill.js'
import { createLineCap, resolveBorderWidth } from './line.js'
import { valToPts } from '../../units-internal.js'
import { el, raw, voidEl } from '../oxml/el.js'

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
			strXml += el(
				`a:${obj.name}`,
				{ w: valToPts(resolveBorderWidth(border, 1)), cap, cmpd: 'sng', algn: 'ctr' },
				[
					genXmlColorSelection({ color: border.color ?? '363636', transparency: border.transparency }),
					voidEl('a:prstDash', { val: border.type === 'dash' ? 'sysDash' : 'solid' }),
					voidEl('a:round'),
					voidEl('a:headEnd', { type: 'none', w: 'med', len: 'med' }),
					voidEl('a:tailEnd', { type: 'none', w: 'med', len: 'med' }),
				].map(raw)
			)
		} else {
			strXml += el(`a:${obj.name}`, { w: 0, cap, cmpd: 'sng', algn: 'ctr' }, raw(voidEl('a:noFill')))
		}
	})
	return strXml
}
