/**
 * ts-pptx: the `<a:lvlNpPr>` list-style level and the theme-font `<a:defRPr>` inside it.
 *
 * A deck states its text defaults in four places — the presentation's `p:defaultTextStyle`, the
 * notesMaster's `p:notesStyle`, and the slide master's built-in and configured `p:txStyles` —
 * and all four write the same level element around the same run-property block: a `solidFill`
 * over one colour, then the `latin`/`ea`/`cs` triple pointing at one theme font family. They
 * differed in the size, in whether `b`/`i` were carried, and in whether the colour and the
 * Latin face were overridable, which is what the parameters here are.
 */

import { LEVEL_PPR_TAIL } from '../../constants-internal.js'
import { el, raw, voidEl, type XmlAttrs, type XmlChild } from '../oxml/el.js'

/**
 * The `<a:defRPr>` a list-style level carries: `attrs` (the size, and `b`/`i`/`kern` where the
 * caller states them), then the fill and the three typeface slots.
 *
 * @param font - the theme font family the three slots point at (`mj` = major/heading, `mn` = minor/body)
 * @param attrs - the run-property attributes, in emission order; `null` values are omitted
 * @param colorXml - the colour element inside `<a:solidFill>`; defaults to `<a:schemeClr val="tx1"/>`
 * @param latinXml - the `<a:latin>` element; defaults to the `font` family's own `+xx-lt` token
 */
export function themeFontDefRPr(font: 'mj' | 'mn', attrs: XmlAttrs, colorXml?: string, latinXml?: string): string {
	return el('a:defRPr', attrs, [
		raw(el('a:solidFill', null, raw(colorXml ?? voidEl('a:schemeClr', { val: 'tx1' })))),
		raw(latinXml ?? voidEl('a:latin', { typeface: `+${font}-lt` })),
		raw(voidEl('a:ea', { typeface: `+${font}-ea` })),
		raw(voidEl('a:cs', { typeface: `+${font}-cs` })),
	])
}

/**
 * One `<a:lvlNpPr>`, with {@link LEVEL_PPR_TAIL} appended to the caller's attributes.
 *
 * The tail is the fixed `defTabSz`/`rtl`/`eaLnBrk`/`latinLnBrk`/`hangingPunct` run every level
 * in the deck carries; it sits last because attribute order is byte-significant.
 * @param level - the 1-based list level (`1` emits `a:lvl1pPr`)
 * @param attrs - the level's own attributes, ahead of the tail
 * @param children - the level's children, in schema order
 */
export function lvlPPr(level: number, attrs: XmlAttrs, children: XmlChild[]): string {
	return el(`a:lvl${level}pPr`, { ...attrs, ...LEVEL_PPR_TAIL }, children)
}
