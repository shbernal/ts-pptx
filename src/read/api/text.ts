/**
 * Read/write proxies for a shape's text: `TextFrame → Paragraph[] → Run[]`.
 *
 * Each proxy wraps a live DOM element (`a:txBody`, `a:p`, `a:r`) and holds the
 * owning `Part`, so a setter can mutate the node in place and call
 * `part.markDirty()` — that single flag is what makes `save()` reserialize the
 * part. Getters compute from the DOM on each access rather than caching.
 */
import type { Part } from '../opc/part.js'
import {
	ELEMENT_NODE,
	attr,
	boolValue,
	firstChild,
	getElements,
	getOrAddChild,
	intValue,
	removeAttr,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import { normalizeHex, setSolidFill, solidFillColor } from '../oxml/fill.js'

// Schema successors used to keep `a:rPr` children in document order when a
// setter has to create one (CT_TextCharacterProperties sequence).
const RPR_AFTER_FILL = [
	'a:effectLst',
	'a:effectDag',
	'a:highlight',
	'a:uLnTx',
	'a:uLn',
	'a:uFillTx',
	'a:uFill',
	'a:latin',
	'a:ea',
	'a:cs',
	'a:sym',
	'a:hlinkClick',
	'a:hlinkMouseOver',
	'a:rtl',
	'a:extLst',
]
const RPR_AFTER_LATIN = ['a:ea', 'a:cs', 'a:sym', 'a:hlinkClick', 'a:hlinkMouseOver', 'a:rtl', 'a:extLst']

/** One text run (`a:r`): a span of text with uniform character formatting. */
export class Run {
	constructor(
		private readonly element: Element,
		private readonly part: Part
	) {}

	/** The run's text (`a:t`), verbatim — whitespace is not normalized. */
	get text(): string {
		return firstChild(this.element, 'a:t')?.textContent ?? ''
	}

	set text(value: string) {
		const t = getOrAddChild(this.element, 'a:t')
		t.textContent = value
		// Preserve significant leading/trailing whitespace per the XML spec.
		if (value !== value.trim()) setAttr(t, 'xml:space', 'preserve')
		else removeAttr(t, 'xml:space')
		this.part.markDirty()
	}

	/** Font size in points (`a:rPr/@sz` is hundredths of a point), or `null` if unset. */
	get fontSizePt(): number | null {
		const size = this.#rPrAttr('sz')
		return size === null ? null : size / 100
	}

	set fontSizePt(value: number | null) {
		if (value === null) {
			this.#removeRPrAttr('sz')
			return
		}
		if (!Number.isFinite(value) || value <= 0) throw new Error(`fontSizePt must be a positive number, got ${value}`)
		setAttr(this.#getOrAddRPr(), 'sz', String(Math.round(value * 100)))
		this.part.markDirty()
	}

	/** Bold (`a:rPr/@b`), or `null` when unset (inherited from style). */
	get bold(): boolean | null {
		return boolValue(this.#rPrAttrRaw('b'))
	}

	set bold(value: boolean | null) {
		this.#setBoolRPrAttr('b', value)
	}

	/** Italic (`a:rPr/@i`), or `null` when unset (inherited from style). */
	get italic(): boolean | null {
		return boolValue(this.#rPrAttrRaw('i'))
	}

	set italic(value: boolean | null) {
		this.#setBoolRPrAttr('i', value)
	}

	/** Underline style token (`a:rPr/@u`, e.g. `sng`), or `null` when unset. */
	get underline(): string | null {
		return this.#rPrAttrRaw('u')
	}

	set underline(value: string | null) {
		if (value === null) {
			this.#removeRPrAttr('u')
			return
		}
		setAttr(this.#getOrAddRPr(), 'u', value)
		this.part.markDirty()
	}

	/** Latin typeface name (`a:rPr/a:latin/@typeface`), or `null` when unset. */
	get fontName(): string | null {
		const rPr = this.#rPr()
		const latin = rPr && firstChild(rPr, 'a:latin')
		return latin ? attr(latin, 'typeface') : null
	}

	set fontName(value: string | null) {
		if (value === null) {
			const rPr = this.#rPr()
			if (rPr) removeChildrenByQName(rPr, ['a:latin'])
			if (rPr) this.part.markDirty()
			return
		}
		const latin = getOrAddChild(this.#getOrAddRPr(), 'a:latin', RPR_AFTER_LATIN)
		setAttr(latin, 'typeface', value)
		this.part.markDirty()
	}

	/** Explicit RGB fill colour as a 6-hex string (`a:solidFill/a:srgbClr/@val`), or `null`. */
	get color(): string | null {
		return solidFillColor(this.#rPr(), 'a:srgbClr')
	}

	set color(value: string | null) {
		this.#setSolidFill(value === null ? null : { qname: 'a:srgbClr', val: normalizeHex(value) })
	}

	/** Theme colour token when the fill is a scheme colour (`a:schemeClr/@val`, e.g. `accent2`), or `null`. */
	get schemeColor(): string | null {
		return solidFillColor(this.#rPr(), 'a:schemeClr')
	}

	set schemeColor(value: string | null) {
		this.#setSolidFill(value === null ? null : { qname: 'a:schemeClr', val: value })
	}

	/** The underlying `a:r` element, for advanced reads and future mutation. */
	get element_(): Element {
		return this.element
	}

	#rPr(): Element | null {
		return firstChild(this.element, 'a:rPr')
	}

	#getOrAddRPr(): Element {
		return getOrAddChild(this.element, 'a:rPr', ['a:t'])
	}

	#rPrAttrRaw(name: string): string | null {
		const rPr = this.#rPr()
		return rPr ? attr(rPr, name) : null
	}

	#rPrAttr(name: string): number | null {
		return intValue(this.#rPrAttrRaw(name))
	}

	#removeRPrAttr(name: string): void {
		const rPr = this.#rPr()
		if (!rPr) return
		removeAttr(rPr, name)
		this.part.markDirty()
	}

	#setBoolRPrAttr(name: string, value: boolean | null): void {
		if (value === null) {
			this.#removeRPrAttr(name)
			return
		}
		setAttr(this.#getOrAddRPr(), name, value ? '1' : '0')
		this.part.markDirty()
	}

	/** Replace the run's solid fill with a single colour element, or clear it when `null`. */
	#setSolidFill(color: { qname: string; val: string } | null): void {
		if (color === null) {
			const rPr = this.#rPr()
			if (!rPr) return
			removeChildrenByQName(rPr, ['a:solidFill'])
			this.part.markDirty()
			return
		}
		setSolidFill(this.#getOrAddRPr(), RPR_AFTER_FILL, color)
		this.part.markDirty()
	}
}

/** One paragraph (`a:p`) of a text frame. */
export class Paragraph {
	constructor(
		private readonly element: Element,
		private readonly part: Part
	) {}

	/** The runs (`a:r`) in document order. Fields (`a:fld`) and breaks are not runs; see `text`. */
	get runs(): Run[] {
		return getElements(this.element, 'a:r').map((element) => new Run(element, this.part))
	}

	/** Indent level (`a:pPr/@lvl`), 0 when unset. */
	get level(): number {
		const pPr = firstChild(this.element, 'a:pPr')
		return (pPr && intValue(attr(pPr, 'lvl'))) ?? 0
	}

	/**
	 * The paragraph's text: run (`a:r`) and field (`a:fld`) text concatenated in
	 * document order, with each line break (`a:br`) rendered as `\n`.
	 */
	get text(): string {
		let out = ''
		for (let node = this.element.firstChild; node; node = node.nextSibling) {
			if (node.nodeType !== ELEMENT_NODE) continue
			const element = node as Element
			if (element.localName === 'r' || element.localName === 'fld') {
				out += firstChild(element, 'a:t')?.textContent ?? ''
			} else if (element.localName === 'br') {
				out += '\n'
			}
		}
		return out
	}

	/** The underlying `a:p` element, for advanced reads and future mutation. */
	get element_(): Element {
		return this.element
	}
}

/** A shape's text frame (`p:txBody`): an ordered list of paragraphs. */
export class TextFrame {
	constructor(
		private readonly txBody: Element,
		private readonly part: Part
	) {}

	/** Paragraphs (`a:p`) in document order. */
	get paragraphs(): Paragraph[] {
		return getElements(this.txBody, 'a:p').map((element) => new Paragraph(element, this.part))
	}

	/** All paragraph text joined by `\n` (mirrors python-pptx `TextFrame.text`). */
	get text(): string {
		return this.paragraphs.map((paragraph) => paragraph.text).join('\n')
	}

	/** The underlying `p:txBody` element, for advanced reads and future mutation. */
	get element_(): Element {
		return this.txBody
	}
}
