/**
 * ts-pptx: DrawingML native equations (OMML)
 *
 * Wrap raw OMML markup into the `<a14:m>` marker PowerPoint uses for editable
 * equations — as a standalone display paragraph (`genXmlMathParagraph`) or as an
 * inline run that flows between the surrounding `<a:r>` runs (`genXmlInlineMath`).
 */

import { el, raw, voidEl } from '../oxml/el.js'
import { OOXML_NS } from '../../ooxml/namespaces.js'

/** The OMML namespace. One module writes it, so it stays here rather than in the registry. */
const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

/** The `<a14:m>` marker, carrying both namespace declarations so the supplied OMML needs none. */
function a14Math(mathXml: string): string {
	return el('a14:m', { 'xmlns:a14': OOXML_NS.a14, 'xmlns:m': M_NS }, raw(mathXml))
}

/**
 * Build a native-equation paragraph (`<a:p>`) from raw OMML.
 *
 * PowerPoint stores an editable equation inside a text body as an `<a14:m>` marker wrapping
 * `<m:oMathPara><m:oMath>…`. We declare both the `a14` (drawing-2010) and `m` (math) namespaces
 * on the `<a14:m>` element so the supplied OMML needs no namespace declarations of its own, then
 * accept three input shapes: a full `<m:oMathPara>`, a full `<m:oMath>`, or the inner OMML
 * (children of `<m:oMath>`). A trailing `<a:endParaRPr>` matches what PowerPoint authors.
 *
 * @param {string} omml - raw OMML markup for the equation
 * @returns {string} an `<a:p>` math paragraph
 */
export function genXmlMathParagraph(omml: string): string {
	const trimmed = (omml || '').trim()
	const paraPr = el('m:oMathParaPr', null, raw(voidEl('m:jc', { 'm:val': 'centerGroup' })))
	const mathXml = trimmed.startsWith('<m:oMathPara')
		? trimmed
		: trimmed.includes('<m:oMath')
			? el('m:oMathPara', null, [paraPr, trimmed].map(raw))
			: el('m:oMathPara', null, [paraPr, el('m:oMath', null, raw(trimmed))].map(raw))
	return el('a:p', null, [a14Math(mathXml), voidEl('a:endParaRPr', { lang: 'en-US' })].map(raw))
}

/**
 * Build an INLINE native-equation run (dn-inline-math) from raw OMML.
 *
 * Unlike the display form ({@link genXmlMathParagraph}), an inline equation is *not* its own
 * paragraph: PowerPoint authors an `<a14:m>` marker wrapping a bare `<m:oMath>` (no `<m:oMathPara>`,
 * no `<m:oMathParaPr>`/`<m:jc>`) that flows between the surrounding `<a:r>` runs in one `<a:p>` —
 * pinned by the `math-omml-inline.pptx` oracle. The `mc:AlternateContent` envelope stays at the
 * shape level (see `objectHasMath`), so nothing wraps the run itself. We declare the `a14` and `m`
 * namespaces on the `<a14:m>` element so the supplied OMML needs none of its own, and accept the
 * same three input shapes as the display helper: a full `<m:oMathPara>` (its inner `<m:oMath>` is
 * unwrapped, since a paragraph block cannot flow inline), a full `<m:oMath>`, or the inner OMML.
 *
 * @param {string} omml - raw OMML markup for the equation
 * @returns {string} an `<a14:m>` inline equation run
 */
export function genXmlInlineMath(omml: string): string {
	const trimmed = (omml || '').trim()
	// Extract a bare <m:oMath>…</m:oMath> — this both strips a display <m:oMathPara> wrapper and
	// leaves an already-bare <m:oMath> untouched; inner-only OMML is wrapped in <m:oMath>.
	const oMathMatch = trimmed.match(/<m:oMath[\s>][\s\S]*<\/m:oMath>/)
	return a14Math(oMathMatch ? oMathMatch[0] : el('m:oMath', null, raw(trimmed)))
}
