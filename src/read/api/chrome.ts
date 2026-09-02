/**
 * Read-model proxies for a deck's shared *chrome* — the slide masters, slide
 * layouts, and themes reachable through the presentation → master → layout → theme
 * graph but owned by no single slide. A slide's concrete content is modeled by
 * {@link import('./slide.js').Slide}; this module models the property tiers a slide
 * *inherits* from: the theme's colour scheme and font scheme, the master's colour
 * map, and each master/layout placeholder's own geometry.
 *
 * The three parts form a chain — a {@link SlideLayout} resolves its {@link SlideMaster}
 * (which resolves its {@link Theme}) — and a {@link Slide} enters it through
 * `slide.layout`. Placeholders (master/layout `p:sp` shapes carrying a `p:ph`) are
 * modeled as {@link Placeholder}s: type/idx/name/id, own-`a:xfrm` geometry (EMU),
 * and a navigable {@link TextFrame}. Unlike a notes placeholder — whose geometry is
 * inherited and so reads `null` on an authored deck — a master/layout placeholder
 * carries its geometry explicitly, so position/size round-trip.
 *
 * A master's and a layout's *non*-placeholder content — the bands, rules, logos and
 * decorative furniture a template is recognized by — is reached through
 * {@link SlideMaster.shapes} / {@link SlideLayout.shapes}, which return the same
 * `AnyShape` union `Slide.shapes` does (both classes are {@link ShapeHost}s). Those
 * shapes carry the full paint surface — `resolvedFill`, `resolvedLine`,
 * `presetGeometry`, `rotation`, `absoluteFrame` — so a template renders, not merely
 * positions. `placeholders` remains the filtered convenience view over the same tree.
 *
 * A *slide* placeholder's effective inherited geometry against this chain resolves
 * via `Shape.resolvedFrame` (`shapes.ts`), the geometry sibling of the run
 * colour/size/face resolution already backed by `Slide.themeContext` → `Run.resolved*`.
 */
import type { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import type { Relationships } from '../opc/relationships.js'
import { attr, boolValue, firstChild, getElements, type Element } from '../oxml/dom.js'
import { parseClrMap, parseClrScheme, type ThemeContext } from '../oxml/theme.js'
import { resolveLayoutColorContext, resolveMasterColorContext } from './theme-context.js'
import { placeholderOf } from '../oxml/placeholder-inherit.js'
import { backgroundElementOf, readSlideBackground, type SlideBackground } from './slide-background.js'
import { AutoShape, buildShapes, findShapeByIdDeep, type AnyShape, type ShapeHost } from './shapes.js'
import { TextFrame } from './text.js'
import { SLIDE_MASTER_REL, THEME_REL } from '../../ooxml/rel-types.js'
import { cSldOf, spTreeOf } from '../oxml/slide-dom.js'
import { COLOR_MAP_TOKENS, THEME_COLOR_SLOTS, type ColorMapToken, type ThemeColorSlot } from '../../ooxml/st-enums.js'

/**
 * The twelve theme colour slots and the twelve colour-map tokens that point at them, plus
 * the tuples this module iterates. Both are schema vocabulary shared with the write half
 * (`gen/slide/master.ts` emits the identity map, `script/from-read` recognises it), so they
 * are declared in `ooxml/st-enums.ts` and re-exported here to keep `ts-pptx/read`'s surface
 * unchanged.
 */
export type { ThemeColorSlot, ColorMapToken }

/**
 * One theme font family (`a:majorFont` or `a:minorFont`): the primary Latin face
 * plus the East-Asian and complex-script faces. Each is `null` when the theme
 * leaves that slot empty (`typeface=""`, PowerPoint's default for `ea`/`cs`). The
 * per-script `<a:font>` fallback list is not decoded here.
 */
export interface ThemeFontFace {
	/** `<a:latin@typeface>` — the primary face (e.g. `Calibri Light`), or `null` when empty. */
	latin: string | null
	/** `<a:ea@typeface>` — the East-Asian face, or `null` when empty. */
	ea: string | null
	/** `<a:cs@typeface>` — the complex-script face, or `null` when empty. */
	cs: string | null
}

/** A theme's `a:fontScheme`: its name plus the major (heading) and minor (body) font families. */
export interface ThemeFontScheme {
	/** `<a:fontScheme@name>` (e.g. `Office`), or `null` when unnamed. */
	name: string | null
	/** The `a:majorFont` family — the heading (`+mj-*`) fonts. */
	major: ThemeFontFace
	/** The `a:minorFont` family — the body (`+mn-*`) fonts. */
	minor: ThemeFontFace
}

/** Read `<a:latin>`/`<a:ea>`/`<a:cs>` typefaces of a theme font family; empty string ⇒ `null`. */
function readThemeFontFace(font: Element | null): ThemeFontFace {
	const face = (child: string): string | null => {
		const el = font && firstChild(font, child)
		return (el && attr(el, 'typeface')) || null
	}
	return { latin: face('a:latin'), ea: face('a:ea'), cs: face('a:cs') }
}

/**
 * A deck's theme part (`theme1.xml`) as a modeled object. Exposes the two property
 * tiers a slide resolves against: the {@link colorScheme} (the 12 `a:clrScheme`
 * slots as literal hex, `a:sysClr` resolved via its `lastClr`) and the
 * {@link fontScheme} (major/minor Latin/EA/CS faces). Reachable from
 * {@link SlideMaster.theme} and `Slide.theme`.
 */
export class Theme {
	constructor(
		/** The theme's OPC part (`/ppt/theme/themeN.xml`). */
		readonly part: Part
	) {}

	/** Partname of the theme part. */
	get partName(): string {
		return this.part.partName
	}

	/** Theme name (`a:theme/@name`, e.g. `Office Theme`), or `null` when unnamed. */
	get name(): string | null {
		const root = this.part.dom.documentElement
		return root ? attr(root, 'name') : null
	}

	/** Colour-scheme name (`a:clrScheme/@name`, e.g. `Office`), or `null`. */
	get colorSchemeName(): string | null {
		const clrScheme = this.#clrScheme()
		return clrScheme ? attr(clrScheme, 'name') : null
	}

	/**
	 * The theme colour scheme: the 12 {@link ThemeColorSlot}s resolved to a 6-hex RGB
	 * (an `a:sysClr` slot resolves through its `lastClr`). A slot the theme does not
	 * define reads `null`. Use {@link color} to look one slot up.
	 */
	get colorScheme(): Record<ThemeColorSlot, string | null> {
		const parsed = parseClrScheme(this.#clrScheme())
		const out = {} as Record<ThemeColorSlot, string | null>
		for (const slot of THEME_COLOR_SLOTS) out[slot] = parsed.get(slot) ?? null
		return out
	}

	/** The 6-hex RGB of one colour-scheme slot, or `null` when the theme does not define it. */
	color(slot: ThemeColorSlot): string | null {
		return parseClrScheme(this.#clrScheme()).get(slot) ?? null
	}

	/** The theme's `a:fontScheme` (major/minor font families), or `null` when the theme declares none. */
	get fontScheme(): ThemeFontScheme | null {
		const themeElements = this.#themeElements()
		const fontScheme = themeElements && firstChild(themeElements, 'a:fontScheme')
		if (!fontScheme) return null
		return {
			name: attr(fontScheme, 'name'),
			major: readThemeFontFace(firstChild(fontScheme, 'a:majorFont')),
			minor: readThemeFontFace(firstChild(fontScheme, 'a:minorFont')),
		}
	}

	/** Escape hatch: the underlying `a:theme` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element | null {
		return this.part.dom.documentElement
	}

	/** Mark the theme part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}

	#themeElements(): Element | null {
		const root = this.part.dom.documentElement
		return root ? firstChild(root, 'a:themeElements') : null
	}

	#clrScheme(): Element | null {
		const themeElements = this.#themeElements()
		return themeElements ? firstChild(themeElements, 'a:clrScheme') : null
	}
}

/** The placeholder shapes (`p:sp` carrying a `p:ph`) of a master/layout `p:spTree`, in document order. */
function placeholderShapes(spTree: Element | null): Element[] {
	if (!spTree) return []
	return getElements(spTree, 'p:sp').filter((sp) => placeholderOf(sp) !== null)
}

/**
 * One placeholder shape (`p:sp` with a `p:ph`) of a slide master or slide layout —
 * the definition a slide's placeholder of the same `type`/`idx` inherits geometry
 * and text style from. Geometry getters read the placeholder's *own* `a:xfrm`,
 * which a master/layout placeholder carries explicitly (so it round-trips, unlike a
 * notes placeholder's inherited geometry).
 *
 * This is the identity-and-geometry view. The same `p:sp` also appears in
 * {@link SlideMaster.shapes} / {@link SlideLayout.shapes} as an `AutoShape`, which
 * additionally carries the paint surface (`resolvedFill`, `resolvedLine`,
 * `presetGeometry`, `rotation`, `absoluteFrame`) and reports its `p:ph` through
 * `AutoShape.placeholder`. Read a placeholder here to *place* it; read it there to
 * *draw* it.
 */
export class Placeholder {
	/**
	 * The same `p:sp` as an {@link AutoShape}. Identity, geometry and the text frame are
	 * forwarded to it rather than re-derived here.
	 *
	 * They used to be re-derived, and the two answers had drifted apart. `#cNvPr()` hard-coded
	 * `p:nvSpPr` where `nonVisualCNvPr` finds any `p:nv*Pr`; the four geometry getters read the
	 * `p:spPr/a:xfrm` chain through a helper no other caller used, against `Shape`'s own
	 * `xfrm()` + `emuFrom`; and,
	 * the one that showed, the text frame was built with **no** inheritance context, so
	 * `Run.resolvedSizePt`, `resolvedFontFace`, `resolvedColor` and `resolvedBold` resolved
	 * through `SlideMaster.shapes` and came back `null` through `SlideMaster.placeholders` —
	 * two views of one element disagreeing about the same run.
	 */
	protected readonly shape: AutoShape

	constructor(
		/** The placeholder's `p:sp`. `protected` so {@link NotesPlaceholder} can read it. */
		protected readonly sp: Element,
		/** The part the `p:sp` lives in — a master, a layout, or a notes slide. */
		protected readonly host: ShapeHost
	) {
		this.shape = new AutoShape(sp, host)
	}

	/** Placeholder type (`p:ph/@type`: `title` | `body` | `sldNum` | …), or `null` when absent (a body placeholder). */
	get type(): string | null {
		const ph = placeholderOf(this.sp)
		return ph ? attr(ph, 'type') : null
	}

	/**
	 * Placeholder index (`p:ph/@idx`), or `null` when unset.
	 *
	 * Not forwarded to {@link AutoShape.placeholder}, whose `idx` defaults to `'0'` the way
	 * PowerPoint resolves it. Here the absence is the contract: this view reports what the
	 * part says.
	 */
	get idx(): string | null {
		const ph = placeholderOf(this.sp)
		return ph ? attr(ph, 'idx') : null
	}

	/** Shape name (`p:cNvPr/@name`), or `''` when unnamed. */
	get name(): string {
		return this.shape.name
	}

	/** Drawing id (`p:cNvPr/@id`), or `null` when absent. */
	get id(): number | null {
		return this.shape.id
	}

	/** Left edge in EMU (`a:off/@x`), or `null` when the placeholder carries no own `a:xfrm`. */
	get left(): number | null {
		return this.shape.left
	}

	/** Top edge in EMU (`a:off/@y`), or `null` when the placeholder carries no own `a:xfrm`. */
	get top(): number | null {
		return this.shape.top
	}

	/** Width in EMU (`a:ext/@cx`), or `null` when the placeholder carries no own `a:xfrm`. */
	get width(): number | null {
		return this.shape.width
	}

	/** Height in EMU (`a:ext/@cy`), or `null` when the placeholder carries no own `a:xfrm`. */
	get height(): number | null {
		return this.shape.height
	}

	/**
	 * The placeholder's text as a navigable {@link TextFrame} (`p:txBody`), or `null`
	 * when it carries no text body. Threaded with the owning tier's theme context and
	 * placeholder identity, so a run's `resolved*` getters walk the same inheritance chain
	 * they walk through {@link SlideMaster.shapes}.
	 */
	get textFrame(): TextFrame | null {
		return this.shape.textFrame
	}

	/** Escape hatch: the underlying `p:sp` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.sp
	}

	/** Mark the owning part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.host.part.markDirty()
	}
}

/**
 * The shared body of a slide master and a slide layout — the two template tiers a slide
 * inherits from.
 *
 * Both are an OPC part holding a `p:cSld` with a shape tree and a resolvable theme context, and
 * both were written out in full, member for member: `partName`, `relationships`, `name`,
 * `shapes`, `shapeByIdDeep`, `placeholders`, `themeContext` and the two DOM lookups behind them
 * were character-for-character identical in the two classes. Nothing about that was broken, but
 * it meant every accessor added to a template tier had to be added twice, and the ones meant to
 * agree could quietly stop agreeing.
 *
 * What genuinely differs stays on the subclasses. `theme` differs because a layout has to walk
 * to its master first. `background` differs in the tier label it decodes with and in where it
 * finds the theme part. And the context itself differs, which is why {@link resolveThemeContext}
 * is abstract: a master resolves its own `p:clrMap` + theme, a layout walks layout → master →
 * theme. The caching around it is shared, so a subclass supplies the walk and nothing else.
 *
 * `Slide` deliberately does not extend this. It is a shape host too, but it is constructed from
 * a `Presentation` rather than an `OpcPackage`, reaches its relationships through that, and
 * reports an unnamed slide as `null` where a template tier reports `''` — three differences in
 * the same handful of members, which is more contortion than the sharing is worth.
 */
abstract class TemplatePart implements ShapeHost {
	#themeContext?: ThemeContext

	constructor(
		/** The package this part belongs to, for reaching related parts (theme, master, media). */
		readonly opc: OpcPackage,
		/** This tier's OPC part. */
		readonly part: Part
	) {}

	/** Partname of this part. */
	get partName(): string {
		return this.part.partName
	}

	/** This part's relationships (its theme or master, image embeds, hyperlinks, …). */
	get relationships(): Relationships {
		return this.opc.relationshipsFor(this.partName)
	}

	/** The authoring name (`p:cSld/@name`), or `''` when unnamed. */
	get name(): string {
		const cSld = this.cSld()
		return (cSld && attr(cSld, 'name')) ?? ''
	}

	/**
	 * Every shape in this tier's `p:cSld/p:spTree`, in document order — the same `AnyShape`
	 * union `Slide.shapes` returns, so shape-walking code applies unchanged.
	 *
	 * {@link placeholders} is the filtered view of the same tree.
	 */
	get shapes(): AnyShape[] {
		const spTree = this.spTree()
		return spTree ? buildShapes(spTree, this) : []
	}

	/** The shape anywhere in this tier's tree with the given drawing id, or `undefined`. */
	shapeByIdDeep(id: number): AnyShape | undefined {
		return findShapeByIdDeep(this.shapes, id)
	}

	/** This tier's placeholder shapes (`p:sp` carrying a `p:ph`), in document order. */
	get placeholders(): Placeholder[] {
		return placeholderShapes(this.spTree()).map((sp) => new Placeholder(sp, this))
	}

	/** This tier's colour/font context, resolved once and cached. Backs each {@link Placeholder}'s text frame. */
	themeContext(): ThemeContext {
		return (this.#themeContext ??= this.resolveThemeContext())
	}

	/** The theme this tier resolves against, or `null` when the chain is incomplete. */
	abstract get theme(): Theme | null

	/** This tier's *own* background (`p:cSld/p:bg`), decoded, or `null` when it defines none. */
	abstract get background(): SlideBackground | null

	/** Walk this tier's colour/font chain. Called once; {@link themeContext} caches the result. */
	protected abstract resolveThemeContext(): ThemeContext

	/** The `p:cSld` of this part, or `null`. */
	protected cSld(): Element | null {
		return cSldOf(this.part.dom.documentElement)
	}

	/** The `p:cSld/p:spTree` of this part, or `null`. */
	protected spTree(): Element | null {
		return spTreeOf(this.part.dom.documentElement)
	}

	/** The part this one's first relationship of `type` points at, or `null`. */
	protected relTarget(type: string): Part | null {
		const rels = this.relationships
		const rel = rels.byType(type)[0]
		return rel ? (this.opc.part(rels.resolveTarget(rel.id)) ?? null) : null
	}
}

/**
 * A deck's slide master (`slideMasterN.xml`) as a modeled object. Exposes the
 * property tiers a bound slide inherits: the {@link colorMap} (token → theme slot),
 * its {@link theme}, its `shapes` and the `placeholders` subset of them, its own
 * {@link background}, and the {@link layouts} that build on it. Reachable from
 * {@link SlideLayout.master} and `Slide.master`.
 *
 * `shapes` is where a template's non-placeholder furniture lives: the header band, the
 * rule under the title, the logo. Whether a given slide draws them is
 * `Slide.showMasterSp` (and, for a layout, {@link SlideLayout.showMasterSp}).
 */
export class SlideMaster extends TemplatePart {
	/** The master's theme (via its `theme` relationship), or `null` when absent. */
	override get theme(): Theme | null {
		const themePart = this.relTarget(THEME_REL)
		return themePart ? new Theme(themePart) : null
	}

	/**
	 * The master's colour map (`p:clrMap`): each of the 12 {@link ColorMapToken}s
	 * mapped to the {@link ThemeColorSlot} it resolves through. This is the indirection
	 * that lets a slide's `schemeClr val="tx1"` reach a theme slot (e.g. `tx1` → `dk1`).
	 * A token the map omits reads `null`.
	 */
	get colorMap(): Record<ColorMapToken, string | null> {
		const root = this.part.dom.documentElement
		const parsed = parseClrMap(root ? firstChild(root, 'p:clrMap') : null)
		const out = {} as Record<ColorMapToken, string | null>
		for (const token of COLOR_MAP_TOKENS) out[token] = parsed.get(token) ?? null
		return out
	}

	/** The layouts built on this master (via `p:sldLayoutIdLst` → its relationships), in list order. */
	get layouts(): SlideLayout[] {
		const root = this.part.dom.documentElement
		const lst = root && firstChild(root, 'p:sldLayoutIdLst')
		if (!lst) return []
		const rels = this.relationships
		const out: SlideLayout[] = []
		for (const entry of getElements(lst, 'p:sldLayoutId')) {
			const relId = attr(entry, 'r:id')
			const part = relId ? this.opc.part(rels.resolveTarget(relId)) : null
			if (part) out.push(new SlideLayout(this.opc, part))
		}
		return out
	}

	/**
	 * The master's *own* background (`p:cSld/p:bg`), decoded, or `null` when it defines
	 * none. Scoped to the master's own element — a slide's *effective* background,
	 * resolved through the slide → layout → master chain, is `Slide.background`.
	 */
	override get background(): SlideBackground | null {
		const bg = backgroundElementOf(this.part.dom.documentElement)
		if (!bg) return null
		const themePart = this.relTarget(THEME_REL)
		return readSlideBackground(
			bg,
			'master',
			this.themeContext(),
			this.relationships,
			themePart ? this.opc.relationshipsFor(themePart.partName) : null
		)
	}

	/** A master's context is its own `p:clrMap` plus its theme — no walk beyond this part. */
	protected override resolveThemeContext(): ThemeContext {
		return resolveMasterColorContext(this.opc, this.partName)
	}
}

/**
 * A deck's slide layout (`slideLayoutN.xml`) as a modeled object — the gallery entry
 * a slide binds to. Exposes its {@link name}, its {@link type} (import-only; the
 * writer authors none), its {@link master} (and, through it, the {@link theme}), its
 * `shapes` and the `placeholders` subset of them, and its own {@link background}.
 * Reachable from `Slide.layout` and {@link SlideMaster.layouts}.
 *
 * The write API authors non-placeholder shapes into a layout: every
 * `defineSlideMaster({ objects })` member that is not a `placeholder` — a `rect`, a
 * `line`, an `image`, a `chart`, a `text` box — lands in this tier's tree.
 */
export class SlideLayout extends TemplatePart {
	/**
	 * The layout type (`p:sldLayout/@type`, e.g. `title` | `obj` | `blank`), or `null`
	 * when the layout declares none. Import-only: the writer authors no `@type`, so an
	 * authored deck reads `null`; an imported deck carries PowerPoint's value.
	 */
	get type(): string | null {
		const root = this.part.dom.documentElement
		return root ? attr(root, 'type') : null
	}

	/**
	 * Whether this layout draws the master's non-placeholder shapes
	 * (`p:sldLayout/@showMasterSp`). `xsd:boolean` defaulting to `true`, so an absent
	 * attribute means shown — the layout-tier counterpart of `Slide.showMasterSp`,
	 * and the one PowerPoint actually writes on a section-divider or full-bleed
	 * layout. A renderer resolving what to paint under a slide's own shapes has to
	 * consult both: the slide's flag, then its layout's.
	 */
	get showMasterSp(): boolean {
		const root = this.part.dom.documentElement
		return boolValue(root && attr(root, 'showMasterSp')) !== false
	}

	/** The master this layout is built on (via its `slideMaster` relationship), or `null` when absent. */
	get master(): SlideMaster | null {
		const rels = this.relationships
		const rel = rels.byType(SLIDE_MASTER_REL)[0]
		const part = rel ? this.opc.part(rels.resolveTarget(rel.id)) : null
		return part ? new SlideMaster(this.opc, part) : null
	}

	/** The theme this layout resolves against, via its {@link master}, or `null`. */
	override get theme(): Theme | null {
		return this.master?.theme ?? null
	}

	/**
	 * The layout's *own* background (`p:cSld/p:bg`), decoded, or `null` when it defines
	 * none. A layout usually authors one — an explicit fill (`p:bgPr`) or a
	 * theme-indexed reference (`p:bgRef`). The slide's *effective* background is
	 * `Slide.background`.
	 */
	override get background(): SlideBackground | null {
		const bg = backgroundElementOf(this.part.dom.documentElement)
		if (!bg) return null
		const themePart = this.master?.theme?.part ?? null
		return readSlideBackground(
			bg,
			'layout',
			this.themeContext(),
			this.relationships,
			themePart ? this.opc.relationshipsFor(themePart.partName) : null
		)
	}

	/** A layout's context is walked layout → master → theme, which is why it is not the master's. */
	protected override resolveThemeContext(): ThemeContext {
		return resolveLayoutColorContext(this.opc, this.partName)
	}
}
