/**
 * The `Run` read/write proxy (`a:r`), and the value types a run's own reads answer with.
 *
 * A run is where inheritance actually bites: an uncoloured, unsized, unfaced run resolves
 * through the placeholder tier, the shape's `p:style/a:fontRef` and the presentation's
 * `p:defaultTextStyle`, which is why {@link PlaceholderTextContext} threads down to here from
 * the frame.
 */
import type { Part } from '../../opc/part.js'
import type { Relationships } from '../../opc/relationships.js'
import {
	attr,
	boolValue,
	type Element,
	firstChild,
	firstChildElement,
	getOrAddChild,
	numberValue,
	pctPointsAttr,
	removeAttr,
	removeChildrenByQName,
	setAttr,
} from '../../oxml/dom.js'
import { colorValueIf, normalizeHex, setSolidFill, solidFillColor } from '../../oxml/fill.js'
import { resolveThemeFont, type ThemeContext } from '../../oxml/theme.js'
import {
	resolveColorElement,
	resolveSolidFillColor,
	type PlaceholderRef,
	type ResolvedColor,
	type StyleFontRef,
} from '../theme-context.js'
import { InvalidOptionError } from '../../../errors.js'
import { ptToHundredths } from '../../../units.js'
import { RPR_FILL_AFTER, RPR_LATIN_AFTER } from '../../../ooxml/sequence.js'
import { ptFromHundredths } from '../coords.js'

/**
 * What a {@link Run}'s text body needs to resolve an *inherited* run
 * colour/size/face/bold: which placeholder the text lives in (or `null` for a
 * non-placeholder shape, which still resolves its `p:style/a:fontRef` and the
 * presentation `p:defaultTextStyle`), the slide theme context (with the
 * layout/master roots) to resolve against, and the shape's resolved
 * `p:style/a:fontRef` text tier. The owning slide's text body `a:lstStyle` is added
 * per text frame. Absent only for text reached without a theme context (table cells).
 */
export interface PlaceholderTextContext {
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
	 *
	 * `@baseline` is `a:ST_Percentage`, a union that also admits `"62.5%"` — the only form the
	 * Strict profile has — so it is read through `parsePercentPoints` rather than as a bare number.
	 */
	get baselinePct(): number | null {
		const rPr = this.#rPr()
		return rPr ? pctPointsAttr(rPr, 'baseline') : null
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
