/**
 * ts-pptx: building the DOM for a shape added through the read model.
 *
 * Write-side element construction, living in `read/` because that is where the caller is:
 * `Slide.addTextBox` and `Slide.addPicture` mint a `p:sp`/`p:pic` and splice it into a loaded
 * deck's `p:spTree`. It is not the write path's emitter -- that builds strings for a package
 * assembled from nothing, while these build nodes for a package that already exists -- and it
 * shares no state with `Slide`, which is why it sat as a hundred lines of free functions at the
 * bottom of that module with no reference to the class above them.
 */

import type { Document, Element } from '@xmldom/xmldom'
import { createElement, ownerDocumentOf, setAttr } from '../oxml/dom.js'

/** The box a built shape sits in, in EMU. */
interface BoxSpec {
	left: number
	top: number
	width: number
	height: number
}

/**
 * Create `qname` in the parent's document and append it.
 *
 * The builders below are uniformly parent-then-child, so this is the only shape
 * of element creation they need; `ownerDocumentOf` is what lets it take the
 * parent alone rather than threading a `Document` through every call.
 */
function appendEl(parent: Element, qname: string): Element {
	const child = createElement(ownerDocumentOf(parent), qname)
	parent.appendChild(child)
	return child
}

/**
 * Append the `p:spPr` a built `p:sp` and `p:pic` share: the spec's box as an
 * `a:xfrm`, then a rect `prstGeom` with the empty `a:avLst` the schema requires.
 */
function appendSpPr(parent: Element, spec: BoxSpec): Element {
	const spPr = appendEl(parent, 'p:spPr')
	const xfrm = appendEl(spPr, 'a:xfrm')
	const off = appendEl(xfrm, 'a:off')
	setAttr(off, 'x', String(spec.left))
	setAttr(off, 'y', String(spec.top))
	const ext = appendEl(xfrm, 'a:ext')
	setAttr(ext, 'cx', String(spec.width))
	setAttr(ext, 'cy', String(spec.height))
	const prstGeom = appendEl(spPr, 'a:prstGeom')
	setAttr(prstGeom, 'prst', 'rect')
	appendEl(prstGeom, 'a:avLst')
	return spPr
}

interface TextBoxSpec extends BoxSpec {
	id: number
	name: string
	text: string
}

/** Build a minimal, schema-valid text-box `p:sp` element (not yet attached). */
export function buildTextBox(doc: Document, spec: TextBoxSpec): Element {
	const sp = createElement(doc, 'p:sp')

	const nvSpPr = appendEl(sp, 'p:nvSpPr')
	const cNvPr = appendEl(nvSpPr, 'p:cNvPr')
	setAttr(cNvPr, 'id', String(spec.id))
	setAttr(cNvPr, 'name', spec.name)
	const cNvSpPr = appendEl(nvSpPr, 'p:cNvSpPr')
	setAttr(cNvSpPr, 'txBox', '1')
	appendEl(nvSpPr, 'p:nvPr')

	appendSpPr(sp, spec)

	const txBody = appendEl(sp, 'p:txBody')
	appendEl(txBody, 'a:bodyPr')
	appendEl(txBody, 'a:lstStyle')
	const p = appendEl(txBody, 'a:p')
	if (spec.text !== '') {
		const r = appendEl(p, 'a:r')
		const t = appendEl(r, 'a:t')
		t.textContent = spec.text
		if (spec.text !== spec.text.trim()) setAttr(t, 'xml:space', 'preserve')
	}

	return sp
}

interface PictureSpec extends BoxSpec {
	id: number
	name: string
	relId: string
}

/** Build a minimal, schema-valid `p:pic` element (not yet attached). */
export function buildPicture(doc: Document, spec: PictureSpec): Element {
	const pic = createElement(doc, 'p:pic')

	const nvPicPr = appendEl(pic, 'p:nvPicPr')
	const cNvPr = appendEl(nvPicPr, 'p:cNvPr')
	setAttr(cNvPr, 'id', String(spec.id))
	setAttr(cNvPr, 'name', spec.name)
	const cNvPicPr = appendEl(nvPicPr, 'p:cNvPicPr')
	const picLocks = appendEl(cNvPicPr, 'a:picLocks')
	setAttr(picLocks, 'noChangeAspect', '1')
	appendEl(nvPicPr, 'p:nvPr')

	const blipFill = appendEl(pic, 'p:blipFill')
	const blip = appendEl(blipFill, 'a:blip')
	setAttr(blip, 'r:embed', spec.relId)
	const stretch = appendEl(blipFill, 'a:stretch')
	appendEl(stretch, 'a:fillRect')

	appendSpPr(pic, spec)

	return pic
}
