/**
 * The `TextFrame` read/write proxy (`p:txBody` / `a:txBody`), and its `a:bodyPr` value types.
 *
 * The frame is where a shape's text begins: it owns the body properties (insets, anchor,
 * wrap, autofit) and hands each `a:p` to a {@link Paragraph}.
 */
import type { Part } from '../../opc/part.js'
import type { Relationships } from '../../opc/relationships.js'
import { attr, type Element, firstChild, getElements, numberValue, pctPointsAttr } from '../../oxml/dom.js'
import { type ThemeContext } from '../../oxml/theme.js'
import { resolveInheritedAnchor } from '../theme-context.js'
import { EMU_PER_POINT } from '../../../units.js'
import type { PlaceholderTextContext } from './run.js'
import { Paragraph } from './paragraph.js'
import { setTextBodyText } from './edit.js'
/**
 * A text frame's body properties (`a:bodyPr`), as read from a shape. Carries the
 * inset/anchor/vertical-text settings that govern where text sits inside its box
 * — the difference between a label that clears a left rail and one that overlaps
 * it. Only **explicitly set** attributes are reported; an absent inset means the
 * PowerPoint default (`lIns`/`rIns` = 0.1", `tIns`/`bIns` = 0.05").
 */
export interface BodyProperties {
	/** Text direction (`@vert`: `horz`/`vert`/`vert270`/`eaVert`/`wordArtVert`…), or `null` when horizontal/unset. */
	vert: string | null
	/** Vertical anchor (`@anchor`: `t`/`ctr`/`b`), or `null` when unset (defaults to top). */
	anchor: string | null
	/** Wrap mode (`@wrap`: `square`/`none`), or `null` when unset. */
	wrap: string | null
	/** Explicitly-set text insets in points (`@lIns`/`@rIns`/`@tIns`/`@bIns` ÷ 12700); a missing side uses the PowerPoint default. */
	insetsPt: { left?: number; right?: number; top?: number; bottom?: number }
}

/**
 * A text frame's vertical-autofit mode, read from its `a:bodyPr` autofit child:
 * - `'none'`       — `a:noAutofit`, or no autofit child at all (the box has a fixed
 *                    size and text can overflow); write-side `fit: 'none'`.
 * - `'normAutofit'`— `a:normAutofit` (shrink text to fit); write-side `fit: 'shrink'`.
 * - `'spAutoFit'`  — `a:spAutoFit` (resize the shape to fit text); write-side `fit: 'resize'`.
 */
export type AutofitMode = 'none' | 'normAutofit' | 'spAutoFit'

export class TextFrame {
	constructor(
		private readonly txBody: Element,
		private readonly part: Part,
		/** The owning slide's theme context (colour maps + `fontScheme`), threaded to each {@link Paragraph}/{@link Run} for the `resolved*` getters. */
		private readonly themeContext?: ThemeContext,
		/**
		 * The placeholder this text body lives in, when any — enables
		 * placeholder-inherited run colour/size/face resolution. Absent for ordinary
		 * text boxes and table cells.
		 */
		private readonly placeholder?: PlaceholderTextContext,
		/** The owning part's relationships, threaded to each {@link Paragraph}/{@link Run} for hyperlink `@r:id` resolution; absent when reached without them. */
		private readonly relationships?: Relationships
	) {}

	/** Paragraphs (`a:p`) in document order. */
	get paragraphs(): Paragraph[] {
		// The slide text body's own list style is the tier just below the run/paragraph
		// in the placeholder inheritance chain; resolve it once and share it.
		const inherit = this.placeholder
			? { placeholder: this.placeholder, slideLstStyle: firstChild(this.txBody, 'a:lstStyle') }
			: undefined
		return getElements(this.txBody, 'a:p').map(
			(element) => new Paragraph(element, this.part, this.themeContext, inherit, this.relationships)
		)
	}

	/**
	 * The frame's body properties (`a:bodyPr`: insets, anchor, vertical text), or
	 * `null` when there is no `a:bodyPr`. Only explicitly-set insets are reported
	 * (a missing side is the PowerPoint default — see {@link BodyProperties}).
	 */
	get bodyProperties(): BodyProperties | null {
		const bodyPr = firstChild(this.txBody, 'a:bodyPr')
		if (!bodyPr) return null
		const insetsPt: BodyProperties['insetsPt'] = {}
		const inset = (qn: string, key: keyof BodyProperties['insetsPt']): void => {
			const v = numberValue(attr(bodyPr, qn))
			if (v !== null) insetsPt[key] = v / EMU_PER_POINT
		}
		inset('lIns', 'left')
		inset('rIns', 'right')
		inset('tIns', 'top')
		inset('bIns', 'bottom')
		return {
			vert: attr(bodyPr, 'vert') ?? null,
			anchor: attr(bodyPr, 'anchor') ?? null,
			wrap: attr(bodyPr, 'wrap') ?? null,
			insetsPt,
		}
	}

	/**
	 * The effective vertical anchor (`t`/`ctr`/`b`): the frame's own
	 * `a:bodyPr/@anchor` when set, else the anchor it inherits from its
	 * layout → master placeholder `a:bodyPr`. `null` when nothing in the chain sets
	 * one (PowerPoint then defaults to top). Unlike {@link bodyProperties}'s
	 * `anchor` (own attribute only), this surfaces the inherited anchor a
	 * placeholder title relies on — the difference between a top- and
	 * centre-anchored title that own-attribute reads alone cannot see.
	 */
	get resolvedAnchor(): string | null {
		const own = this.bodyProperties?.anchor
		if (own) return own
		// Anchor inheritance is placeholder-only; a non-placeholder frame (ph null) has none.
		const ph = this.placeholder?.ph
		if (!ph) return null
		return resolveInheritedAnchor(ph, this.placeholder.flatten)
	}

	/**
	 * The frame's vertical-autofit mode (`a:bodyPr` autofit child), or `null` when
	 * there is no `a:bodyPr`. A `bodyPr` with no autofit child — or an explicit
	 * `a:noAutofit` — reads `'none'`. See {@link AutofitMode}. This is the deep-model
	 * counterpart of `inspectPptx`'s per-shape `autofit`; unlike the shallow surface
	 * it hangs off the navigable text frame.
	 */
	get autofit(): AutofitMode | null {
		const bodyPr = firstChild(this.txBody, 'a:bodyPr')
		if (!bodyPr) return null
		if (firstChild(bodyPr, 'a:spAutoFit')) return 'spAutoFit'
		if (firstChild(bodyPr, 'a:normAutofit')) return 'normAutofit'
		return 'none'
	}

	/**
	 * The baked shrink scale of a `normAutofit` frame as a percent (62.5 = 62.5%),
	 * or `null` when the frame has no `a:normAutofit` or bakes no scale. A bare
	 * `<a:normAutofit/>` (write-side `fit: 'shrink'`) carries no scale — PowerPoint
	 * computes it on edit and draws at 100% until then — so it too reads `null`; an
	 * explicit scale comes from `fit: { type: 'shrink', fontScale }`.
	 */
	get autofitFontScale(): number | null {
		return this.#normAutofitPct('fontScale')
	}

	/**
	 * The baked line-spacing reduction of a `normAutofit` frame as a percent
	 * (`a:normAutofit/@lnSpcReduction` ÷ 1000), or `null` when unset. The companion
	 * to {@link autofitFontScale} (write-side `fit: { type: 'shrink', lnSpcReduction }`).
	 */
	get autofitLineSpaceReduction(): number | null {
		return this.#normAutofitPct('lnSpcReduction')
	}

	/**
	 * A `a:normAutofit` percentage attribute → percent, or `null`.
	 *
	 * `fontScale` is `ST_TextFontScalePercentOrPercentString` and `lnSpcReduction` is
	 * `ST_TextSpacingPercentOrPercentString`; both admit the `"62.5%"` spelling alongside the
	 * fixed-point one, and it is the only spelling Strict has.
	 */
	#normAutofitPct(name: string): number | null {
		const bodyPr = firstChild(this.txBody, 'a:bodyPr')
		const norm = bodyPr && firstChild(bodyPr, 'a:normAutofit')
		return norm ? pctPointsAttr(norm, name) : null
	}

	/** All paragraph text joined by `\n` (mirrors python-pptx `TextFrame.text`). */
	get text(): string {
		return this.paragraphs.map((paragraph) => paragraph.text).join('\n')
	}

	/**
	 * Replace the frame's text with a single paragraph and run, preserving the
	 * character formatting (`a:rPr`) of the frame's first existing run when there
	 * is one. For multiple runs or per-run formatting, edit
	 * `paragraphs[].runs[]` directly.
	 */
	set text(value: string) {
		setTextBodyText(this.txBody, value)
		this.part.markDirty()
	}

	/** Escape hatch: the underlying `p:txBody` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.txBody
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}
}
