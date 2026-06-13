/**
 * Read-model proxies for a shape's text: `TextFrame → Paragraph[] → Run[]`.
 *
 * Each proxy wraps a live DOM element (`a:txBody`, `a:p`, `a:r`) so the same
 * nodes can later be mutated in place. Phase 2 is read-only; getters compute
 * from the DOM on each access rather than caching.
 */
import { ELEMENT_NODE, attr, boolValue, firstChild, getElements, intValue, type Element } from '../oxml/dom.js'

/** One text run (`a:r`): a span of text with uniform character formatting. */
export class Run {
	constructor(private readonly element: Element) {}

	/** The run's text (`a:t`), verbatim — whitespace is not normalized. */
	get text(): string {
		return firstChild(this.element, 'a:t')?.textContent ?? ''
	}

	/** Font size in points (`a:rPr/@sz` is hundredths of a point), or `null` if unset. */
	get fontSizePt(): number | null {
		const size = this.#rPrAttr('sz')
		return size === null ? null : size / 100
	}

	/** Bold (`a:rPr/@b`), or `null` when unset (inherited from style). */
	get bold(): boolean | null {
		return boolValue(this.#rPrAttrRaw('b'))
	}

	/** Italic (`a:rPr/@i`), or `null` when unset (inherited from style). */
	get italic(): boolean | null {
		return boolValue(this.#rPrAttrRaw('i'))
	}

	/** Underline style token (`a:rPr/@u`, e.g. `sng`), or `null` when unset. */
	get underline(): string | null {
		return this.#rPrAttrRaw('u')
	}

	/** Latin typeface name (`a:rPr/a:latin/@typeface`), or `null` when unset. */
	get fontName(): string | null {
		const rPr = this.#rPr()
		const latin = rPr && firstChild(rPr, 'a:latin')
		return latin ? attr(latin, 'typeface') : null
	}

	/** Explicit RGB fill colour as a 6-hex string (`a:solidFill/a:srgbClr/@val`), or `null`. */
	get color(): string | null {
		const srgb = this.#solidFillChild('a:srgbClr')
		return srgb ? attr(srgb, 'val') : null
	}

	/** Theme colour token when the fill is a scheme colour (`a:schemeClr/@val`, e.g. `accent2`), or `null`. */
	get schemeColor(): string | null {
		const scheme = this.#solidFillChild('a:schemeClr')
		return scheme ? attr(scheme, 'val') : null
	}

	/** The underlying `a:r` element, for advanced reads and future mutation. */
	get element_(): Element {
		return this.element
	}

	#rPr(): Element | null {
		return firstChild(this.element, 'a:rPr')
	}

	#rPrAttrRaw(name: string): string | null {
		const rPr = this.#rPr()
		return rPr ? attr(rPr, name) : null
	}

	#rPrAttr(name: string): number | null {
		return intValue(this.#rPrAttrRaw(name))
	}

	#solidFillChild(qname: string): Element | null {
		const rPr = this.#rPr()
		const fill = rPr && firstChild(rPr, 'a:solidFill')
		return fill ? firstChild(fill, qname) : null
	}
}

/** One paragraph (`a:p`) of a text frame. */
export class Paragraph {
	constructor(private readonly element: Element) {}

	/** The runs (`a:r`) in document order. Fields (`a:fld`) and breaks are not runs; see `text`. */
	get runs(): Run[] {
		return getElements(this.element, 'a:r').map((element) => new Run(element))
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
	constructor(private readonly txBody: Element) {}

	/** Paragraphs (`a:p`) in document order. */
	get paragraphs(): Paragraph[] {
		return getElements(this.txBody, 'a:p').map((element) => new Paragraph(element))
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
