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
import type { ColorContext } from '../oxml/theme.js'
import { resolveSolidFillColor, type ResolvedColor } from './theme-context.js'

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
		private readonly part: Part,
		/** The owning slide's theme colour context, for {@link resolvedColor}; absent when the run was reached without one. */
		private readonly themeColors?: ColorContext
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

	/**
	 * The run's own solid fill colour resolved against the owning slide's theme to
	 * a literal hex — the resolved counterpart of {@link color}/{@link schemeColor}.
	 * `null` when the run sets no colour of its own (it then inherits from the
	 * placeholder / list-style chain, which this getter does not walk), the colour
	 * cannot be made literal, or the run was reached without a theme context. The
	 * returned {@link ResolvedColor} carries `effectiveHex` — the base colour with
	 * its child transforms (`lumMod`/`shade`/…) applied — for the final rendered
	 * colour.
	 */
	get resolvedColor(): ResolvedColor | null {
		return this.themeColors ? resolveSolidFillColor(this.#rPr(), this.themeColors) : null
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
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to each {@link Run} for `resolvedColor`. */
		private readonly themeColors?: ColorContext
	) {}

	/** The runs (`a:r`) in document order. Fields (`a:fld`) and breaks are not runs; see `text`. */
	get runs(): Run[] {
		return getElements(this.element, 'a:r').map((element) => new Run(element, this.part, this.themeColors))
	}

	/** Indent level (`a:pPr/@lvl`), 0 when unset. */
	get level(): number {
		const pPr = firstChild(this.element, 'a:pPr')
		return (pPr && intValue(attr(pPr, 'lvl'))) ?? 0
	}

	/**
	 * Horizontal alignment token (`a:pPr/@algn`: `l` | `ctr` | `r` | `just` |
	 * `dist` | `thaiDist`), or `null` when unset (inherited from the list style).
	 */
	get align(): string | null {
		const pPr = firstChild(this.element, 'a:pPr')
		return pPr ? attr(pPr, 'algn') : null
	}

	/**
	 * Space before the paragraph in points (`a:pPr/a:spcBef/a:spcPts/@val` is
	 * hundredths of a point), or `null` when unset or expressed as a percentage
	 * (`a:spcPct`, which has no fixed point value).
	 */
	get spaceBeforePt(): number | null {
		return this.#spacingPt('a:spcBef')
	}

	/** Space after the paragraph in points; see {@link spaceBeforePt} for the percentage caveat. */
	get spaceAfterPt(): number | null {
		return this.#spacingPt('a:spcAft')
	}

	/** Left margin of the paragraph in points (`a:pPr/@marL` is EMU; 12700 EMU = 1pt), or `null` when unset. */
	get marginLeftPt(): number | null {
		return this.#emuAttrPt('marL')
	}

	/** First-line indent in points (`a:pPr/@indent` is EMU; negative for a hanging indent), or `null` when unset. */
	get indentPt(): number | null {
		return this.#emuAttrPt('indent')
	}

	/**
	 * Bullet description for this paragraph, derived from the `a:pPr` bullet
	 * children, or `null` when unset (inherited from the list style):
	 * - `'none'`          — explicit `a:buNone` (bullet suppressed)
	 * - `'char:•'`        — `a:buChar/@char` (the literal glyph follows the colon)
	 * - `'autoNum:arabicPeriod'` — `a:buAutoNum/@type` (auto-numbered)
	 */
	get bullet(): string | null {
		const pPr = firstChild(this.element, 'a:pPr')
		if (!pPr) return null
		if (firstChild(pPr, 'a:buNone')) return 'none'
		const buChar = firstChild(pPr, 'a:buChar')
		if (buChar) return `char:${attr(buChar, 'char') ?? ''}`
		const buAutoNum = firstChild(pPr, 'a:buAutoNum')
		if (buAutoNum) return `autoNum:${attr(buAutoNum, 'type') ?? ''}`
		return null
	}

	/** Points from a spacing child's `a:spcPts/@val` (hundredths of a point), or `null`. */
	#spacingPt(qname: string): number | null {
		const pPr = firstChild(this.element, 'a:pPr')
		const spc = pPr && firstChild(pPr, qname)
		const pts = spc && firstChild(spc, 'a:spcPts')
		const val = pts ? intValue(attr(pts, 'val')) : null
		return val === null ? null : val / 100
	}

	/** Points from an EMU-valued `a:pPr` attribute (`marL` / `indent`), or `null`. */
	#emuAttrPt(name: string): number | null {
		const pPr = firstChild(this.element, 'a:pPr')
		const emu = pPr ? intValue(attr(pPr, name)) : null
		return emu === null ? null : emu / 12700
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
		private readonly part: Part,
		/** The owning slide's theme colour context, threaded to each {@link Paragraph}/{@link Run} for `resolvedColor`. */
		private readonly themeColors?: ColorContext
	) {}

	/** Paragraphs (`a:p`) in document order. */
	get paragraphs(): Paragraph[] {
		return getElements(this.txBody, 'a:p').map((element) => new Paragraph(element, this.part, this.themeColors))
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
