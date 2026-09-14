/**
 * The `Run` read/write proxy (`a:r`), and the value types a run's own reads answer with.
 *
 * A run is where inheritance actually bites: an uncoloured, unsized, unfaced run resolves
 * through the placeholder tier, the shape's `p:style/a:fontRef` and the presentation's
 * `p:defaultTextStyle`, which is why {@link TextContext} threads down to here from the frame.
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
import { applySolidFill, hasFillChoice, solidFillColor, type SolidFillEdit } from '../../oxml/fill.js'
import { checkEnumOrThrow } from '../../../ooxml/check-enum.js'
import { TEXT_UNDERLINE_TYPES } from '../../../ooxml/st-enums.js'
import { resolveThemeFont, type ThemeContext } from '../../oxml/theme.js'
import {
	resolveColorElement,
	resolveSolidFillColor,
	type ColorRef,
	type PlaceholderRef,
	type ResolvedColor,
	type StyleFontRef,
} from '../theme-context.js'
import { InvalidOptionError } from '../../../errors.js'
import { ptToHundredths } from '../../../units.js'
import { RPR_FILL_AFTER, RPR_LATIN_AFTER } from '../../../ooxml/sequence.js'
import { ptFromHundredths } from '../coords.js'
import type { TableCellTextStyle } from '../table-style-resolve.js'

/**
 * What a text body is read against, handed from a {@link TextFrame} to each paragraph and run: the
 * part an edit marks dirty, the theme the `resolved*` getters resolve against, the relationships
 * a hyperlink or a picture bullet resolves through, and the inheritance its runs fall back through.
 */
export interface TextContext {
	/** The part the text lives in. */
	part: Part
	/** The theme the text resolves against: the colour maps, `fontScheme`, and the layout and master roots. */
	ctx: ThemeContext
	/**
	 * The owning part's relationships, or `null` for text read without them — a SmartArt drawing's
	 * cached text — whose run hyperlinks then report their raw `@r:id`, `@action` and `@tooltip`.
	 */
	rels: Relationships | null
	/**
	 * What a run that sets no colour, size, face, bold or italic of its own inherits through, or
	 * `null` for text that inherits through nothing, such as a SmartArt point's.
	 */
	inherit: TextInheritance | null
}

/** The inheritance a shape's or a table cell's text resolves through below its own run properties. */
export interface TextInheritance {
	/**
	 * The placeholder the text lives in, or `null` for a non-placeholder shape, which still resolves
	 * its `p:style/a:fontRef` and the presentation's `p:defaultTextStyle`.
	 */
	ph: PlaceholderRef | null
	/** The shape's resolved `p:style/a:fontRef` colour and face, or `null` when it has none. */
	fontRef: StyleFontRef | null
	/**
	 * A table cell's text style: what the table style's `a:tcTxStyle` gives the cell's region, with
	 * the theme's `tx1` and minor font where it names no colour or face. Read after {@link fontRef}
	 * and before the list-style chain. Absent for text outside a table.
	 */
	tableText?: TableCellTextStyle | null
}

/**
 * What the runs of one paragraph inherit, each property resolved at most once and only when a run
 * that sets no value of its own asks. Every run in a paragraph shares its level and `a:pPr`, so
 * they share one of these.
 */
export interface InheritedRunProps {
	color(): ResolvedColor | null
	size(): number | null
	face(): string | null
	bold(): boolean | null
	italic(): boolean | null
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
	/** The bullet's own colour (`a:buClr`); every field is `null` when the paragraph leaves it to be inherited. */
	colorRef: ColorRef
}

/**
 * A paragraph's bullet, as the structured counterpart of the `a:pPr` bullet
 * children. Discriminated on `kind`, which is one of the four mutually
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
		/**
		 * What the run is read against; see {@link TextContext}. Its `inherit.fontRef` is the tier
		 * {@link resolvedColor} and {@link resolvedFontFace} consult just below the run's own `a:rPr`
		 * and above the placeholder and `p:defaultTextStyle` chain.
		 */
		private readonly context: TextContext,
		/**
		 * What the run inherits when it sets no colour, size, face, bold or italic of its own, or `null`
		 * for text that inherits through nothing. Built by the owning {@link Paragraph} and shared by
		 * its runs; nothing in it is resolved until a run asks.
		 */
		private readonly inherited: InheritedRunProps | null
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
		this.context.part.markDirty()
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
		this.context.part.markDirty()
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

	/** @throws {InvalidOptionError} when the token is not an `ST_TextUnderlineType` */
	set underline(value: string | null) {
		if (value === null) {
			this.#removeRPrAttr('u')
			return
		}
		const token = checkEnumOrThrow(value, TEXT_UNDERLINE_TYPES, 'underline', 'text/invalid-underline')
		setAttr(this.#getOrAddRPr(), 'u', token)
		this.context.part.markDirty()
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
	 * through the owning slide's theme, or `null` when the run has no highlight or
	 * its colour cannot be made literal. The writer authors highlights from a hex
	 * colour, so `effectiveHex` is that colour; imported decks may carry a theme token.
	 */
	get highlight(): ResolvedColor | null {
		const rPr = this.#rPr()
		const hl = rPr && firstChild(rPr, 'a:highlight')
		if (!hl) return null
		const colorEl = firstChildElement(hl)
		return colorEl ? resolveColorElement(colorEl, this.context.ctx) : null
	}

	/**
	 * The run's click hyperlink (`a:rPr/a:hlinkClick`), or `null` when the run
	 * carries none. A URL link resolves its `@r:id` to the external target
	 * ({@link RunHyperlink.url}); a slide jump resolves it to the linked slide's
	 * partname ({@link RunHyperlink.targetPartName}). When the run is read without
	 * its part's relationships (see {@link TextContext.rels}), only the raw
	 * `@r:id`/`@action`/`@tooltip` are reported (the target stays `null`).
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
		const rels = this.context.rels
		if (relId && rels) {
			const rel = rels.get(relId)
			if (rel?.targetMode === 'External') url = rel.target
			else if (rel) targetPartName = rels.resolveTarget(relId)
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
			if (!rPr || !firstChild(rPr, 'a:latin')) return
			removeChildrenByQName(rPr, ['a:latin'])
			this.context.part.markDirty()
			return
		}
		const latin = getOrAddChild(this.#getOrAddRPr(), 'a:latin', RPR_LATIN_AFTER)
		setAttr(latin, 'typeface', value)
		this.context.part.markDirty()
	}

	/** Explicit RGB fill colour as a 6-hex string (`a:solidFill/a:srgbClr/@val`), or `null`. */
	get color(): string | null {
		return solidFillColor(this.#rPr(), 'a:srgbClr')
	}

	set color(value: string | null) {
		this.#applyFill(value === null ? null : { hex: value })
	}

	/** Theme colour token when the fill is a scheme colour (`a:schemeClr/@val`, e.g. `accent2`), or `null`. */
	get schemeColor(): string | null {
		return solidFillColor(this.#rPr(), 'a:schemeClr')
	}

	/** @throws {InvalidOptionError} when the token is not an `ST_SchemeColorVal` */
	set schemeColor(value: string | null) {
		this.#applyFill(value === null ? null : { scheme: value })
	}

	/**
	 * The colour this run effectively renders, resolved against the owning slide's
	 * theme to a literal hex. It is the run's own solid fill
	 * ({@link color}/{@link schemeColor}) when set; otherwise the shape's
	 * `p:style/a:fontRef` colour, then — for a run inside a placeholder — the colour
	 * it inherits from the placeholder/list-style chain (layout → master placeholder
	 * `a:lstStyle` → master `p:txStyles`), then the presentation's
	 * `p:defaultTextStyle`. `null` when the run sets no colour and inherits none, or the
	 * colour cannot be made literal.
	 * `null` too when the run's own fill is not a solid colour (`a:noFill`, a gradient):
	 * the run decides its own colour then, and the one it would otherwise inherit is not
	 * what it paints in, the rule `Shape.resolvedFill` and `TableCell.resolvedFill` follow.
	 * The returned {@link ResolvedColor} carries `effectiveHex` — the base colour with
	 * its child transforms (`lumMod`/`shade`/…) applied — for the final rendered colour.
	 */
	get resolvedColor(): ResolvedColor | null {
		const rPr = this.#rPr()
		if (hasFillChoice(rPr)) return resolveSolidFillColor(rPr, this.context.ctx)
		const inherit = this.context.inherit
		return inherit?.fontRef?.color ?? inherit?.tableText?.color ?? this.inherited?.color() ?? null
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
		return this.fontSizePt ?? this.inherited?.size() ?? null
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
		if (own !== null) return resolveThemeFont(own, this.context.ctx.fontScheme ?? null)
		const inherit = this.context.inherit
		return inherit?.fontRef?.face ?? inherit?.tableText?.face ?? this.inherited?.face() ?? null
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
		return this.bold ?? this.context.inherit?.tableText?.bold ?? this.inherited?.bold() ?? null
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
		return this.italic ?? this.context.inherit?.tableText?.italic ?? this.inherited?.italic() ?? null
	}

	/** Escape hatch: the underlying `a:r` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.element
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.context.part.markDirty()
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
		this.context.part.markDirty()
	}

	#setBoolRPrAttr(name: string, value: boolean | null): void {
		if (value === null) {
			this.#removeRPrAttr(name)
			return
		}
		setAttr(this.#getOrAddRPr(), name, value ? '1' : '0')
		this.context.part.markDirty()
	}

	/** Replace the run's solid fill with a single colour, or clear it when `null`. */
	#applyFill(edit: SolidFillEdit | null): void {
		const changed =
			edit === null
				? applySolidFill(this.#rPr(), null)
				: applySolidFill(this.#rPr(), edit, () => this.#getOrAddRPr(), RPR_FILL_AFTER)
		if (changed) this.context.part.markDirty()
	}
}
