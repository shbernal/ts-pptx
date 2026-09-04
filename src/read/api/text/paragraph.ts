/**
 * The `Paragraph` read/write proxy (`a:p`).
 *
 * A paragraph owns its `a:pPr` — alignment, indents, spacing, the bullet — and hands each
 * `a:r` to a {@link Run}. It reads the same {@link PlaceholderTextContext} rather than
 * resolving inheritance itself.
 */
import type { Part } from '../../opc/part.js'
import type { Relationships } from '../../opc/relationships.js'
import {
	attr,
	type Element,
	ELEMENT_NODE,
	firstChild,
	firstChildElement,
	getElements,
	numberValue,
	pctPointsAttr,
} from '../../oxml/dom.js'
import { colorValueIf } from '../../oxml/fill.js'
import { type ThemeContext } from '../../oxml/theme.js'
import {
	resolveColorElement,
	resolveInheritedRunBold,
	resolveInheritedRunItalic,
	resolveInheritedRunColor,
	resolveInheritedRunFontFace,
	resolveInheritedRunSize,
	type PlaceholderRef,
	type ResolvedColor,
} from '../theme-context.js'
import { HUNDREDTHS_PER_POINT } from '../../../units.js'
import { pctFromThousandths, ptFromEmu, ptFromHundredths } from '../coords.js'
import { Run, type BulletDetail, type BulletStyle, type LineSpacing, type PlaceholderTextContext } from './run.js'
import { setParagraphText } from './edit.js'
/** One paragraph (`a:p`) of a text frame. */
export class Paragraph {
	constructor(
		private readonly element: Element,
		private readonly part: Part,
		/** The owning slide's theme context (colour maps + `fontScheme`), threaded to each {@link Run} for the `resolved*` getters. */
		private readonly themeContext?: ThemeContext,
		/**
		 * Placeholder + slide-list-style context for resolving a placeholder-inherited
		 * run colour/size/face; absent for non-placeholder text. The owning
		 * {@link TextFrame} supplies the placeholder identity and the text body's
		 * `a:lstStyle`.
		 */
		private readonly inherit?: { placeholder: PlaceholderTextContext; slideLstStyle: Element | null },
		/** The owning part's relationships, threaded to each {@link Run} for hyperlink `@r:id` resolution; absent when reached without them. */
		private readonly relationships?: Relationships
	) {}

	/** The runs (`a:r`) in document order. Fields (`a:fld`) and breaks are not runs; see `text`. */
	get runs(): Run[] {
		const inheritedColor = this.#inheritedColorResolver()
		const inheritedSize = this.#inheritedResolver((ph, level, pPr, slideLst, ctx) =>
			resolveInheritedRunSize(ph, level, pPr, slideLst, ctx)
		)
		const inheritedFace = this.#inheritedResolver((ph, level, pPr, slideLst, ctx) =>
			resolveInheritedRunFontFace(ph, level, pPr, slideLst, ctx)
		)
		const inheritedBold = this.#inheritedResolver((ph, level, pPr, slideLst, ctx) =>
			resolveInheritedRunBold(ph, level, pPr, slideLst, ctx)
		)
		const inheritedItalic = this.#inheritedResolver((ph, level, pPr, slideLst, ctx) =>
			resolveInheritedRunItalic(ph, level, pPr, slideLst, ctx)
		)
		const fontRef = this.inherit?.placeholder.fontRef ?? null
		return getElements(this.element, 'a:r').map(
			(element) =>
				new Run(
					element,
					this.part,
					this.themeContext,
					inheritedColor,
					inheritedSize,
					inheritedFace,
					inheritedBold,
					inheritedItalic,
					this.relationships,
					fontRef
				)
		)
	}

	/**
	 * A memoized thunk resolving the colour every run in this paragraph inherits
	 * when it sets none of its own, or `undefined` for non-placeholder paragraphs.
	 * Runs in one paragraph share a level and `a:pPr`, so the lookup runs at most
	 * once per paragraph and only when a colourless run actually asks for it.
	 */
	#inheritedColorResolver(): (() => ResolvedColor | null) | undefined {
		return this.#inheritedResolver((ph, level, pPr, slideLst, ctx) =>
			resolveInheritedRunColor(ph, level, pPr, slideLst, ctx)
		)
	}

	/**
	 * Build a memoized per-paragraph thunk for one inherited run property
	 * (colour/size/face/bold/italic), or `undefined` for non-placeholder paragraphs. All runs in
	 * a paragraph share its level and `a:pPr`, so each `resolve` runs at most once
	 * and only when a run actually lacks its own value and asks.
	 */
	#inheritedResolver<T>(
		resolve: (
			ph: PlaceholderRef | null,
			level: number,
			pPr: Element | null,
			slideLstStyle: Element | null,
			ctx: ThemeContext
		) => T | null
	): (() => T | null) | undefined {
		if (!this.inherit) return undefined
		const { placeholder, slideLstStyle } = this.inherit
		const pPr = firstChild(this.element, 'a:pPr')
		const level = this.level
		let cached: T | null | undefined
		return () =>
			cached === undefined ? (cached = resolve(placeholder.ph, level, pPr, slideLstStyle, placeholder.flatten)) : cached
	}

	/** Indent level (`a:pPr/@lvl`), 0 when unset. */
	get level(): number {
		const pPr = firstChild(this.element, 'a:pPr')
		return (pPr && numberValue(attr(pPr, 'lvl'))) ?? 0
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

	/**
	 * The paragraph's line spacing (`a:pPr/a:lnSpc`), as an exact point height
	 * (`a:spcPts`) or a percentage of the single-line height (`a:spcPct`), or
	 * `null` when unset (inherited from the list style). See {@link LineSpacing}.
	 */
	get lineSpacing(): LineSpacing | null {
		const pPr = firstChild(this.element, 'a:pPr')
		const lnSpc = pPr && firstChild(pPr, 'a:lnSpc')
		if (!lnSpc) return null
		const pts = firstChild(lnSpc, 'a:spcPts')
		if (pts) {
			const val = numberValue(attr(pts, 'val'))
			return val === null ? null : { type: 'points', valuePt: val / HUNDREDTHS_PER_POINT }
		}
		const pct = firstChild(lnSpc, 'a:spcPct')
		if (pct) {
			// `ST_TextSpacingPercentOrPercentString`: the name says it, and the string half is the
			// only spelling Strict has.
			const val = pctPointsAttr(pct, 'val')
			return val === null ? null : { type: 'percent', percent: val }
		}
		return null
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
	 * This paragraph's bullet — its kind, and for a drawn bullet the glyph, the
	 * numbering start, and the bullet's own font/size/colour. `null` when the
	 * paragraph names no bullet at all and inherits one from the list style.
	 *
	 * Structured rather than a tagged string, which is not merely tidier. The
	 * previous `bullet` accessor reported `'none'` / `'char:•'` /
	 * `'autoNum:arabicPeriod'`, and that shape is ambiguous when the glyph is
	 * itself a colon; worse, it reads as a bare glyph, which is exactly how the
	 * script converter first consumed it — `'none'.codePointAt(0)` put a literal
	 * `n` bullet on every converted deck, silently. See {@link BulletDetail}.
	 */
	get bulletDetail(): BulletDetail | null {
		const pPr = firstChild(this.element, 'a:pPr')
		if (!pPr) return null
		if (firstChild(pPr, 'a:buNone')) return { kind: 'none' }

		const style = this.#bulletStyle(pPr)
		const buChar = firstChild(pPr, 'a:buChar')
		if (buChar) return { kind: 'char', char: attr(buChar, 'char') ?? '', ...style }

		const buAutoNum = firstChild(pPr, 'a:buAutoNum')
		if (buAutoNum) {
			return {
				kind: 'autoNum',
				scheme: attr(buAutoNum, 'type') ?? '',
				startAt: numberValue(attr(buAutoNum, 'startAt')),
				...style,
			}
		}

		const buBlip = firstChild(pPr, 'a:buBlip')
		if (buBlip) return { kind: 'picture', imagePartName: this.#bulletImagePartName(buBlip), ...style }

		return null
	}

	/** The `a:buFont` / `a:buSz*` / `a:buClr` siblings of a `a:pPr` bullet choice. */
	#bulletStyle(pPr: Element): BulletStyle {
		const buFont = firstChild(pPr, 'a:buFont')
		const buSzPct = firstChild(pPr, 'a:buSzPct')
		const buSzPts = firstChild(pPr, 'a:buSzPts')
		const buClr = firstChild(pPr, 'a:buClr')
		// `a:buClr` holds the colour element directly (CT_Color), not wrapped in an
		// `a:solidFill` the way a run or a shape fill does.
		const colorEl = buClr ? firstChildElement(buClr) : null
		const pctVal = buSzPct ? numberValue(attr(buSzPct, 'val')) : null
		const ptVal = buSzPts ? numberValue(attr(buSzPts, 'val')) : null
		return {
			font: buFont ? (attr(buFont, 'typeface') ?? null) : null,
			sizePct: pctFromThousandths(pctVal),
			sizePt: ptFromHundredths(ptVal),
			color: colorValueIf(colorEl, 'srgbClr'),
			schemeColor: colorValueIf(colorEl, 'schemeClr'),
			resolvedColor: this.themeContext ? resolveColorElement(colorEl, this.themeContext) : null,
		}
	}

	/** Resolve a picture bullet's `a:buBlip/a:blip/@r:embed` to an absolute partname. */
	#bulletImagePartName(buBlip: Element): string | null {
		const blip = firstChild(buBlip, 'a:blip')
		const relId = blip && attr(blip, 'r:embed')
		if (!relId || !this.relationships) return null
		return this.relationships.resolveTarget(relId)
	}

	/** Points from a spacing child's `a:spcPts/@val` (hundredths of a point), or `null`. */
	#spacingPt(qname: string): number | null {
		const pPr = firstChild(this.element, 'a:pPr')
		const spc = pPr && firstChild(pPr, qname)
		const pts = spc && firstChild(spc, 'a:spcPts')
		const val = pts ? numberValue(attr(pts, 'val')) : null
		return ptFromHundredths(val)
	}

	/** Points from an EMU-valued `a:pPr` attribute (`marL` / `indent`), or `null`. */
	#emuAttrPt(name: string): number | null {
		const pPr = firstChild(this.element, 'a:pPr')
		const emu = pPr ? numberValue(attr(pPr, name)) : null
		return ptFromEmu(emu)
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

	/**
	 * Replace **this paragraph's** content with a single run, preserving the character
	 * formatting (`a:rPr`) of its first existing run and leaving its `a:pPr` (level,
	 * alignment, bullet) alone. Sibling paragraphs are untouched — the difference from
	 * {@link TextFrame.text}, which collapses the whole body to one paragraph.
	 */
	set text(value: string) {
		setParagraphText(this.element, value)
		this.part.markDirty()
	}

	/** Escape hatch: the underlying `a:p` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.element
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}
}

/** A shape's text frame (`p:txBody`): an ordered list of paragraphs. */
