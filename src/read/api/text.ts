/**
 * Read/write proxies for a shape's text: `TextFrame → Paragraph[] → Run[]`.
 *
 * Each proxy wraps a live DOM element (`a:txBody`, `a:p`, `a:r`) and holds the
 * owning `Part`, so a setter can mutate the node in place and call
 * `part.markDirty()` — that single flag is what makes `save()` reserialize the
 * part. Getters compute from the DOM on each access rather than caching.
 */
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import {
	ELEMENT_NODE,
	attr,
	boolValue,
	childElements,
	createElement,
	firstChild,
	firstChildElement,
	getElements,
	getOrAddChild,
	numberValue,
	removeAttr,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../oxml/dom.js'
import { colorValueIf, normalizeHex, setSolidFill, solidFillColor } from '../oxml/fill.js'
import { resolveThemeFont, type ThemeContext } from '../oxml/theme.js'
import {
	resolveColorElement,
	resolveInheritedAnchor,
	resolveInheritedRunBold,
	resolveInheritedRunItalic,
	resolveInheritedRunColor,
	resolveInheritedRunFontFace,
	resolveInheritedRunSize,
	resolveSolidFillColor,
	type PlaceholderRef,
	type ResolvedColor,
	type StyleFontRef,
} from './theme-context.js'
import { InternalError, InvalidOptionError } from '../../errors.js'
import { EMU_PER_POINT, FIXED_PCT_PER_PERCENT, HUNDREDTHS_PER_POINT, ptToHundredths } from '../../units.js'
import { RPR_FILL_AFTER, RPR_LATIN_AFTER } from '../../ooxml/sequence.js'
import { pctFromThousandths, ptFromEmu, ptFromHundredths } from './coords.js'

/**
 * What a {@link Run}'s text body needs to resolve an *inherited* run
 * colour/size/face/bold: which placeholder the text lives in (or `null` for a
 * non-placeholder shape, which still resolves its `p:style/a:fontRef` and the
 * presentation `p:defaultTextStyle`), the slide theme context (with the
 * layout/master roots) to resolve against, and the shape's resolved
 * `p:style/a:fontRef` text tier. The owning slide's text body `a:lstStyle` is added
 * per text frame. Absent only for text reached without a theme context (table cells).
 */
interface PlaceholderTextContext {
	ph: PlaceholderRef | null
	flatten: ThemeContext
	/** The shape's resolved `p:style/a:fontRef` colour + face tier, or `null` when it has none. */
	fontRef?: StyleFontRef | null
}

/**
 * A run's click hyperlink (`a:rPr/a:hlinkClick`): a link on a span of text. A URL
 * link carries an external {@link url}; a slide jump carries the internal
 * {@link targetPartName} (the linked slide's part) alongside its `hlinksldjump`
 * {@link action}. `tooltip` and `relId` are surfaced when present.
 */
export interface RunHyperlink {
	/** External URL target (its `@r:id` resolves to a `TargetMode="External"` rel), or `null` for an internal/action-only link. */
	url: string | null
	/** Absolute partname of an internal target (e.g. the slide a jump points at), or `null`. */
	targetPartName: string | null
	/** Navigation action token (`@action`, e.g. `ppaction://hlinksldjump`), or `null` when absent/empty. */
	action: string | null
	/** Tooltip text (`@tooltip`), or `null` when absent/empty. */
	tooltip: string | null
	/** The relationship id (`@r:id`) backing the link, or `null` when the link is action-only. */
	relId: string | null
}

/**
 * A paragraph's line spacing (`a:pPr/a:lnSpc`), in whichever of the two OOXML
 * forms the file uses: an exact point height (`a:spcPts`) or a percentage of the
 * single-line height (`a:spcPct` — e.g. `percent: 150` for 1.5× spacing).
 */
export type LineSpacing = { type: 'points'; valuePt: number } | { type: 'percent'; percent: number }

/**
 * The bullet's own font, size and colour (`a:buFont` / `a:buSzPct` / `a:buSzPts` /
 * `a:buClr`) — the properties that style the glyph or number itself rather than
 * the text after it. Every field is `null` when the paragraph leaves it to be
 * inherited from the list style.
 *
 * These are siblings of the bullet kind in `a:pPr`, not children of it, so they
 * are carried alongside each kind rather than inside it.
 */
export interface BulletStyle {
	/** `a:buFont/@typeface` — the face the glyph is drawn in (a symbol font such as `Wingdings`), or `null`. */
	font: string | null
	/** `a:buSzPct/@val` as a percentage of the run size (e.g. `80`; the raw attribute is thousandths of a percent), or `null`. */
	sizePct: number | null
	/** `a:buSzPts/@val` as an absolute point size (the raw attribute is hundredths of a point), or `null`. The alternative to {@link sizePct}; at most one is set. */
	sizePt: number | null
	/** Explicit RGB bullet colour as 6-hex (`a:buClr/a:srgbClr/@val`), or `null`. */
	color: string | null
	/** Theme colour token when the bullet colour is a scheme colour (`a:buClr/a:schemeClr/@val`), or `null`. */
	schemeColor: string | null
	/** The bullet colour resolved against the slide theme with its transforms applied, or `null` when unset or not resolvable to a literal. */
	resolvedColor: ResolvedColor | null
}

/**
 * A paragraph's bullet, as the structured counterpart of the `a:pPr` bullet
 * children. Discriminated on {@link kind}, which is one of the four mutually
 * exclusive choices the schema allows:
 *
 * - `'none'`    — `a:buNone`, the bullet explicitly suppressed. Carries no style,
 *                 because there is no glyph to style.
 * - `'char'`    — `a:buChar`, a literal glyph.
 * - `'autoNum'` — `a:buAutoNum`, an auto-numbered list.
 * - `'picture'` — `a:buBlip`, an image used as the glyph.
 *
 * A paragraph that names none of them inherits its bullet from the list style and
 * reports `null` rather than a member of this union.
 */
export type BulletDetail =
	| { kind: 'none' }
	| ({
			kind: 'char'
			/** The glyph itself (`a:buChar/@char`) — a bare character, never a tagged string. */
			char: string
	  } & BulletStyle)
	| ({
			kind: 'autoNum'
			/** The numbering scheme (`a:buAutoNum/@type`, e.g. `arabicPeriod`). */
			scheme: string
			/**
			 * The number this list starts at (`a:buAutoNum/@startAt`), or `null` when
			 * unset (the schema default is 1). Content rather than styling: a list
			 * continuing "5. Deploy" that restarts at 1 is a different slide.
			 */
			startAt: number | null
	  } & BulletStyle)
	| ({
			kind: 'picture'
			/**
			 * Absolute partname of the image used as the glyph (`a:buBlip/a:blip/@r:embed`,
			 * resolved through the owning part's relationships), or `null` when the
			 * paragraph was reached without them or the blip carries no `r:embed`.
			 */
			imagePartName: string | null
	  } & BulletStyle)

/** One text run (`a:r`): a span of text with uniform character formatting. */
export class Run {
	constructor(
		private readonly element: Element,
		private readonly part: Part,
		/** The owning slide's theme context (colour maps + `fontScheme`), for the `resolved*` getters; absent when the run was reached without one. */
		private readonly themeContext?: ThemeContext,
		/**
		 * Resolves the colour this run inherits from its placeholder/list-style chain
		 * when it sets none of its own (item A). Built by the owning {@link Paragraph}
		 * for placeholder text; absent for non-placeholder runs. Called lazily.
		 */
		private readonly inheritedColor?: () => ResolvedColor | null,
		/**
		 * Resolves the point size this run inherits from the same chain when it sets no
		 * own `@sz`. Built by the owning {@link Paragraph} for placeholder text; absent
		 * for non-placeholder runs. Called lazily.
		 */
		private readonly inheritedSize?: () => number | null,
		/**
		 * Resolves the typeface this run inherits from the same chain (a `+mj-*`/`+mn-*`
		 * theme token already resolved to a literal face) when it sets no own `a:latin`.
		 * Built by the owning {@link Paragraph} for placeholder text; absent otherwise.
		 * Called lazily.
		 */
		private readonly inheritedFace?: () => string | null,
		/**
		 * Resolves whether this run inherits bold from the same chain when it sets no
		 * own `@b`. Built by the owning {@link Paragraph} for placeholder text; absent
		 * for non-placeholder runs. Called lazily.
		 */
		private readonly inheritedBold?: () => boolean | null,
		/**
		 * Resolves whether this run inherits italic from the same chain when it sets no
		 * own `@i`. Built by the owning {@link Paragraph} for placeholder text; absent
		 * for non-placeholder runs. Called lazily.
		 */
		private readonly inheritedItalic?: () => boolean | null,
		/**
		 * The owning part's relationships, used to resolve a run hyperlink's `@r:id`
		 * to its external URL or internal target partname. Every frame the read model builds
		 * now carries them — a table cell's and a notes placeholder's included — but the
		 * parameter stays optional for a frame constructed by hand, in which case
		 * {@link hyperlink} still reports the raw `@r:id`/`@action`/`@tooltip`.
		 */
		private readonly relationships?: Relationships,
		/**
		 * The owning shape's resolved `p:style/a:fontRef` colour + face tier — the
		 * fallback consulted for {@link resolvedColor}/{@link resolvedFontFace} just
		 * below the run's own `a:rPr` and above the placeholder/`p:defaultTextStyle`
		 * chain. Absent when the shape has no `p:style/a:fontRef`.
		 */
		private readonly fontRef?: StyleFontRef | null
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
		return ptFromHundredths(size)
	}

	set fontSizePt(value: number | null) {
		if (value === null) {
			this.#removeRPrAttr('sz')
			return
		}
		if (!Number.isFinite(value) || value <= 0)
			throw new InvalidOptionError('font/size-not-positive', `fontSizePt must be a positive number, got ${value}`)
		setAttr(this.#getOrAddRPr(), 'sz', String(ptToHundredths(value)))
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

	/**
	 * Strikethrough token (`a:rPr/@strike`: `noStrike` | `sngStrike` |
	 * `dblStrike`), or `null` when unset (inherited from the style). Surfaced as
	 * the raw token — `sngStrike` is the writer's single-strike value.
	 */
	get strike(): string | null {
		return this.#rPrAttrRaw('strike')
	}

	/**
	 * Capitalization token (`a:rPr/@cap`: `none` | `small` | `all`), or `null`
	 * when unset. `small` renders small-caps, `all` renders all-caps.
	 */
	get caps(): string | null {
		return this.#rPrAttrRaw('cap')
	}

	/**
	 * Baseline shift as a percentage of the font size (`a:rPr/@baseline`, stored
	 * in 1000ths of a percent): positive for superscript (the writer's default is
	 * `30`), negative for subscript (`-40`), or `null` when unset. Reported as the
	 * percentage (`@baseline` ÷ 1000).
	 */
	get baselinePct(): number | null {
		const raw = this.#rPrAttr('baseline')
		return pctFromThousandths(raw)
	}

	/**
	 * Character spacing (tracking) in points (`a:rPr/@spc`, authored in hundredths
	 * of a point), or `null` when unset. Negative tightens. The read counterpart of
	 * the write-side `charSpacing` option.
	 */
	get charSpacingPt(): number | null {
		const raw = this.#rPrAttr('spc')
		return ptFromHundredths(raw)
	}

	/**
	 * The run's highlight colour (`a:rPr/a:highlight`), resolved to a literal hex
	 * through the owning slide's theme, or `null` when the run has no highlight
	 * (or a token colour cannot be made literal without a theme context). The
	 * writer authors highlights from a hex colour, so `effectiveHex` is that
	 * colour; imported decks may carry a theme token, resolved here when possible.
	 */
	get highlight(): ResolvedColor | null {
		const rPr = this.#rPr()
		const hl = rPr && firstChild(rPr, 'a:highlight')
		if (!hl) return null
		const colorEl = firstChildElement(hl)
		if (!colorEl) return null
		if (this.themeContext) return resolveColorElement(colorEl, this.themeContext)
		// Without a theme context only a literal srgbClr can be made concrete.
		const hex = colorValueIf(colorEl, 'srgbClr')
		return hex ? { hex, transforms: [], effectiveHex: hex } : null
	}

	/**
	 * The run's click hyperlink (`a:rPr/a:hlinkClick`), or `null` when the run
	 * carries none. A URL link resolves its `@r:id` to the external target
	 * ({@link RunHyperlink.url}); a slide jump resolves it to the linked slide's
	 * partname ({@link RunHyperlink.targetPartName}). When the run was reached
	 * without the owning part's relationships, only the raw `@r:id`/`@action`/
	 * `@tooltip` are reported (the target stays `null`).
	 */
	get hyperlink(): RunHyperlink | null {
		const rPr = this.#rPr()
		const hlink = rPr && firstChild(rPr, 'a:hlinkClick')
		if (!hlink) return null
		const relId = attr(hlink, 'r:id') || null
		const action = attr(hlink, 'action') || null
		const tooltip = attr(hlink, 'tooltip') || null
		let url: string | null = null
		let targetPartName: string | null = null
		if (relId && this.relationships) {
			const rel = this.relationships.get(relId)
			if (rel?.targetMode === 'External') url = rel.target
			else if (rel) targetPartName = this.relationships.resolveTarget(relId)
		}
		return { url, targetPartName, action, tooltip, relId }
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
		const latin = getOrAddChild(this.#getOrAddRPr(), 'a:latin', RPR_LATIN_AFTER)
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
	 * The colour this run effectively renders, resolved against the owning slide's
	 * theme to a literal hex. It is the run's own solid fill
	 * ({@link color}/{@link schemeColor}) when set; otherwise the shape's
	 * `p:style/a:fontRef` colour, then — for a run inside a placeholder — the colour
	 * it inherits from the placeholder/list-style chain (layout → master placeholder
	 * `a:lstStyle` → master `p:txStyles`), then the presentation's
	 * `p:defaultTextStyle`. `null` when the run sets no colour and inherits none, the
	 * colour cannot be made literal, or the run was reached without a theme context.
	 * The returned {@link ResolvedColor} carries `effectiveHex` — the base colour with
	 * its child transforms (`lumMod`/`shade`/…) applied — for the final rendered colour.
	 */
	get resolvedColor(): ResolvedColor | null {
		if (!this.themeContext) return null
		return (
			resolveSolidFillColor(this.#rPr(), this.themeContext) ?? this.fontRef?.color ?? this.inheritedColor?.() ?? null
		)
	}

	/**
	 * The point size this run effectively renders. It is the run's own `@sz`
	 * ({@link fontSizePt}) when set; otherwise, for a run inside a placeholder, the
	 * size it inherits from the placeholder/list-style chain (paragraph `a:defRPr` →
	 * slide `a:lstStyle` → layout → master placeholder `a:lstStyle` → master
	 * `p:txStyles`). `null` when the run sets no size and inherits none — the
	 * resolved counterpart of {@link fontSizePt}, which reports only the run's own value.
	 */
	get resolvedSizePt(): number | null {
		return this.fontSizePt ?? this.inheritedSize?.() ?? null
	}

	/**
	 * The typeface this run effectively renders, resolved to a literal face name. It
	 * is the run's own `a:latin` ({@link fontName}) when set; otherwise the face named
	 * by the shape's `p:style/a:fontRef` (`idx` → theme major/minor font), then — for a
	 * run inside a placeholder — the face it inherits from the placeholder/list-style
	 * chain, then the presentation's `p:defaultTextStyle`. A `+mj-*`/`+mn-*`
	 * major/minor theme-font token — on the run itself or reached through the chain —
	 * is resolved through the theme `fontScheme` to its concrete face. `null` when the
	 * run names no face and inherits none, or a token cannot be resolved — the resolved
	 * counterpart of {@link fontName}, which reports the raw `@typeface` (possibly a token).
	 */
	get resolvedFontFace(): string | null {
		const own = this.fontName
		if (own !== null) return resolveThemeFont(own, this.themeContext?.fontScheme ?? null)
		return this.fontRef?.face ?? this.inheritedFace?.() ?? null
	}

	/**
	 * Whether this run effectively renders bold. It is the run's own `@b`
	 * ({@link bold}) when set; otherwise, for a run inside a placeholder, the bold
	 * state it inherits from the placeholder/list-style chain (paragraph `a:defRPr` →
	 * slide `a:lstStyle` → layout → master placeholder `a:lstStyle` → master
	 * `p:txStyles`). `null` when the run sets no `@b` and inherits none — the
	 * resolved counterpart of {@link bold}, which reports only the run's own value.
	 */
	get resolvedBold(): boolean | null {
		return this.bold ?? this.inheritedBold?.() ?? null
	}

	/**
	 * Whether this run effectively renders italic. It is the run's own `@i`
	 * ({@link italic}) when set; otherwise, for a run inside a placeholder, the italic
	 * state it inherits from the placeholder/list-style chain (paragraph `a:defRPr` →
	 * slide `a:lstStyle` → layout → master placeholder `a:lstStyle` → master
	 * `p:txStyles`). `null` when the run sets no `@i` and inherits none — the
	 * resolved counterpart of {@link italic}, and the twin of {@link resolvedBold}:
	 * `@b` and `@i` are siblings a master text style states together, so a deck that
	 * can be authored with an inherited italic can be read back with one.
	 */
	get resolvedItalic(): boolean | null {
		return this.italic ?? this.inheritedItalic?.() ?? null
	}

	/** Escape hatch: the underlying `a:r` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.element
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
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
		return numberValue(this.#rPrAttrRaw(name))
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
		setSolidFill(this.#getOrAddRPr(), RPR_FILL_AFTER, color)
		this.part.markDirty()
	}
}

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
			const val = numberValue(attr(pct, 'val'))
			return val === null ? null : { type: 'percent', percent: val / FIXED_PCT_PER_PERCENT }
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

	/** A `a:normAutofit` percentage attribute (stored in 1000ths of a percent → percent), or `null`. */
	#normAutofitPct(name: string): number | null {
		const bodyPr = firstChild(this.txBody, 'a:bodyPr')
		const norm = bodyPr && firstChild(bodyPr, 'a:normAutofit')
		const raw = norm ? numberValue(attr(norm, name)) : null
		return pctFromThousandths(raw)
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
