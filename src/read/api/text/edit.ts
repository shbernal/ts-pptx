/**
 * The two free text-body mutators, shared by the three proxies that own an `a:txBody`.
 *
 * `TextFrame.text`, `TableCell.text` and `DiagramPoint.text` all replace a body's content with
 * one paragraph and one run, preserving the first existing run's `a:rPr`. Neither function
 * marks a part dirty: the caller owns the `Part` and calls `markDirty()` after.
 */
import {
	childElements,
	createElement,
	type Element,
	firstChild,
	getElements,
	removeAttr,
	setAttr,
} from '../../oxml/dom.js'
import { InternalError } from '../../../errors.js'
/**
 * Replace a text body's content (`a:txBody` or `a:txBody`-shaped element) with a
 * single paragraph and run, preserving the `a:rPr` of the body's first existing
 * run when there is one. Shared by {@link TextFrame.text} and `TableCell.text`.
 * Does **not** mark any part dirty — the caller owns the `Part` and must call
 * `markDirty()` after.
 */
export function setTextBodyText(txBody: Element, value: string): void {
	const doc = txBody.ownerDocument
	if (!doc) throw new InternalError('oxml/node-has-no-document', 'Cannot set text: text body has no owner document')

	// Collapse to a single paragraph, dropping any extras, then let the paragraph-level
	// setter do the rest — the run rule is the same one, stated once.
	const paragraphs = getElements(txBody, 'a:p')
	for (let i = paragraphs.length - 1; i >= 1; i--) {
		const extra = paragraphs[i]
		if (extra) txBody.removeChild(extra)
	}
	let p = paragraphs[0]
	if (!p) {
		p = createElement(doc, 'a:p')
		txBody.appendChild(p)
	}
	setParagraphText(p, value)
}

/**
 * Replace **one paragraph's** content with a single run, preserving the `a:rPr` of that
 * paragraph's first existing run and leaving its `a:pPr` (level, alignment, bullet) alone.
 * Sibling paragraphs are untouched, which is the whole difference from
 * {@link setTextBodyText} and the reason this exists: a SmartArt drawing cache packs several
 * nodes' text into one `dsp:txBody`, so collapsing the body there would delete the other
 * nodes' strings. Does **not** mark any part dirty — the caller owns the `Part`.
 */
export function setParagraphText(p: Element, value: string): void {
	const doc = p.ownerDocument
	if (!doc) throw new InternalError('oxml/node-has-no-document', 'Cannot set text: paragraph has no owner document')

	// Capture the first run's character formatting before we discard runs.
	const firstRun = firstChild(p, 'a:r')
	const rPrTemplate = firstRun && firstChild(firstRun, 'a:rPr')

	// Remove every run-level child (runs, breaks, fields); keep a:pPr / a:endParaRPr.
	for (const child of childElements(p)) {
		if (child.localName === 'r' || child.localName === 'br' || child.localName === 'fld') p.removeChild(child)
	}

	// Build a single run, carrying over the captured formatting if present.
	const run = createElement(doc, 'a:r')
	if (rPrTemplate) run.appendChild(rPrTemplate.cloneNode(true))
	const t = createElement(doc, 'a:t')
	t.textContent = value
	if (value !== value.trim()) setAttr(t, 'xml:space', 'preserve')
	else removeAttr(t, 'xml:space')
	run.appendChild(t)

	// Insert before a:endParaRPr if present (it must stay last), else append.
	const endParaRPr = firstChild(p, 'a:endParaRPr')
	p.insertBefore(run, endParaRPr)
}
