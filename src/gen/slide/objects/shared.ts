/**
 * ts-pptx: helpers shared by more than one slide-object renderer
 *
 * The `<p:cNvPr>` open tag, its `<a:hlinkClick>` children and the `<a:ln>` outline are each
 * emitted by several shape kinds. They live here — rather than in the dispatch module — so a
 * renderer never has to import from its own caller.
 */

import type { HyperlinkProps, ShapeLineProps } from '../../../types/index.js'
import { encodeXmlAttrValue } from '../../utils.js'
import { createLineCap, genXmlLineFill } from '../../drawingml/line.js'
import { lineWidthToEmu } from '../../../units-internal.js'
import { el, raw, voidEl } from '../../oxml/el.js'

/** PowerPoint 2010 (`p14`) namespace — carries `<p14:modId>` and `<p14:media>`. */
export const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
/** Markup-compatibility namespace, declared by every `<mc:AlternateContent>` this module's callers emit. */
export const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

/**
 * The `<p:cNvPr>` OPEN tag shared by every shape renderer. Callers append `/>` or
 * `>`+children+`</p:cNvPr>` — the element is self-closing for some shape kinds and paired
 * (hyperlink / media-action children) for others.
 *
 * NOT built with the element builder, deliberately. `descr` is escaped here but `name` is **not**,
 * and that asymmetry is intentional and load-bearing: `objectName` is caller-supplied free text,
 * but every `add*Definition` (text.ts, shape.ts, image.ts, chart.ts, media.ts, connector.ts,
 * group.ts, table.ts) already runs it through `encodeXmlAttrValue(validateObjectName(...))` once
 * before it reaches a slide object's `options`. Escaping it again here would double-encode it
 * (`'Q&A'` -> `Q&amp;A` upstream -> `Q&amp;amp;A` if escaped here too). This helper exists so that
 * single escape stays one line in one place instead of eight call sites re-deriving it.
 * @param id - the shape's `<p:cNvPr>` id, unique slide-wide
 * @param name - caller-supplied `objectName`, already escaped once upstream (emitted as-is)
 * @param descr - alt text (escaped)
 * @param openPrefix - byte-significant indentation before `<p:cNvPr`
 * @returns the open tag, without its closing delimiter
 */
export function cNvPrOpen(id: number, name: string | undefined, descr: string, openPrefix = ''): string {
	return `${openPrefix}<p:cNvPr id="${id}" name="${name}" descr="${encodeXmlAttrValue(descr)}"`
}

/**
 * The `<a:hlinkClick>` children of a shape's `<p:cNvPr>` — a URL link and/or a jump to another
 * slide. Shared by the text and image renderers, which emitted identical copies.
 *
 * NOTE: the tooltip is passed through UNESCAPED and escaped once by the element builder. Escaping
 * it here as well would emit `&amp;amp;` for a tooltip containing `&`.
 * @param link - the shape's hyperlink, if any
 * @returns zero, one or two `<a:hlinkClick>` elements
 */
export function cNvPrHyperlink(link: HyperlinkProps | undefined): string {
	if (!link) return ''
	const tooltip = link.tooltip ?? ''
	return (
		(link.url ? voidEl('a:hlinkClick', { 'r:id': `rId${link._rId}`, tooltip }) : '') +
		(link.slide
			? voidEl('a:hlinkClick', { 'r:id': `rId${link._rId}`, tooltip, action: 'ppaction://hlinksldjump' })
			: '') +
		// Action buttons: a self-contained slide-show navigation action. No relationship, so `r:id`
		// is emitted empty (schema-optional on CT_Hyperlink; matches PowerPoint's own output).
		(link.action
			? voidEl('a:hlinkClick', { 'r:id': '', tooltip, action: `ppaction://hlinkshowjump?jump=${link.action}` })
			: '')
	)
}

/**
 * A shape outline: `<a:ln>` with its fill, dash pattern and arrow ends.
 *
 * The text, connector and image renderers each carried a byte-identical copy of this block; they
 * now share one. Attribute and child order are byte-significant and preserved as written.
 *
 * FUTURE: arrow-size support via the `w`/`len` attrs on headEnd/tailEnd
 * (e.g. `<a:headEnd type="arrow" w="lg" len="lg"/>`; each is 'sm'|'med'|'lg', a 3x3 grid)
 * @param ln - the shape's line properties
 * @returns the `<a:ln>` element
 */
export function genXmlShapeLine(ln: ShapeLineProps): string {
	return el('a:ln', { w: ln.width ? lineWidthToEmu(ln.width) : null, cap: ln.cap ? createLineCap(ln.cap) : null }, [
		raw(genXmlLineFill(ln)),
		ln.dashType ? raw(voidEl('a:prstDash', { val: ln.dashType })) : null,
		ln.beginArrowType ? raw(voidEl('a:headEnd', { type: ln.beginArrowType })) : null,
		ln.endArrowType ? raw(voidEl('a:tailEnd', { type: ln.endArrowType })) : null,
	])
}
