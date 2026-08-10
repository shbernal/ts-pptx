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
import { attr, boolValue, firstChild, getElements, intValue, type Element } from '../oxml/dom.js'
import { parseClrMap, parseClrScheme, type ThemeContext } from '../oxml/theme.js'
import { resolveLayoutColorContext, resolveMasterColorContext } from './theme-context.js'
import { placeholderOf } from '../oxml/placeholder-inherit.js'
import { spPrXfrmEmu } from './shapes/oxml.js'
import { backgroundElementOf, readSlideBackground, type SlideBackground } from './slide-background.js'
import { buildShapes, findShapeByIdDeep, type AnyShape, type ShapeHost } from './shapes.js'
import { TextFrame } from './text.js'
import { SLIDE_MASTER_REL, THEME_REL } from '../../ooxml/rel-types.js'

/** The 12 theme colour-scheme slots (`a:clrScheme` children), in schema order. */
export type ThemeColorSlot =
	| 'dk1'
	| 'lt1'
	| 'dk2'
	| 'lt2'
	| 'accent1'
	| 'accent2'
	| 'accent3'
	| 'accent4'
	| 'accent5'
	| 'accent6'
	| 'hlink'
	| 'folHlink'

const THEME_COLOR_SLOTS: readonly ThemeColorSlot[] = [
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
]

/** The 12 colour-map tokens (`p:clrMap` attributes), each pointing at a {@link ThemeColorSlot}. */
export type ColorMapToken =
	| 'bg1'
	| 'tx1'
	| 'bg2'
	| 'tx2'
	| 'accent1'
	| 'accent2'
	| 'accent3'
	| 'accent4'
	| 'accent5'
	| 'accent6'
	| 'hlink'
	| 'folHlink'

const COLOR_MAP_TOKENS: readonly ColorMapToken[] = [
	'bg1',
	'tx1',
	'bg2',
	'tx2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
]

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
	constructor(
		private readonly sp: Element,
		private readonly part: Part,
		/** The owning master/layout theme context, threaded to {@link textFrame}. */
		private readonly themeContext: ThemeContext,
		/** The owning part's relationships, threaded to {@link textFrame} for hyperlink resolution. */
		private readonly relationships: Relationships
	) {}

	#cNvPr(): Element | null {
		const nvSpPr = firstChild(this.sp, 'p:nvSpPr')
		return nvSpPr ? firstChild(nvSpPr, 'p:cNvPr') : null
	}

	/** Placeholder type (`p:ph/@type`: `title` | `body` | `sldNum` | …), or `null` when absent (a body placeholder). */
	get type(): string | null {
		const ph = placeholderOf(this.sp)
		return ph ? attr(ph, 'type') : null
	}

	/** Placeholder index (`p:ph/@idx`), or `null` when unset. */
	get idx(): string | null {
		const ph = placeholderOf(this.sp)
		return ph ? attr(ph, 'idx') : null
	}

	/** Shape name (`p:cNvPr/@name`), or `''` when unnamed. */
	get name(): string {
		const cNvPr = this.#cNvPr()
		return (cNvPr && attr(cNvPr, 'name')) ?? ''
	}

	/** Drawing id (`p:cNvPr/@id`), or `null` when absent. */
	get id(): number | null {
		const cNvPr = this.#cNvPr()
		return cNvPr ? intValue(attr(cNvPr, 'id')) : null
	}

	/** Left edge in EMU (`a:off/@x`), or `null` when the placeholder carries no own `a:xfrm`. */
	get left(): number | null {
		return spPrXfrmEmu(this.sp, 'a:off', 'x')
	}

	/** Top edge in EMU (`a:off/@y`), or `null` when the placeholder carries no own `a:xfrm`. */
	get top(): number | null {
		return spPrXfrmEmu(this.sp, 'a:off', 'y')
	}

	/** Width in EMU (`a:ext/@cx`), or `null` when the placeholder carries no own `a:xfrm`. */
	get width(): number | null {
		return spPrXfrmEmu(this.sp, 'a:ext', 'cx')
	}

	/** Height in EMU (`a:ext/@cy`), or `null` when the placeholder carries no own `a:xfrm`. */
	get height(): number | null {
		return spPrXfrmEmu(this.sp, 'a:ext', 'cy')
	}

	/**
	 * The placeholder's text as a navigable {@link TextFrame} (`p:txBody`), or `null`
	 * when it carries no text body. Threaded with the owning master/layout theme
	 * context so a run's own `schemeClr` resolves through `Run.resolvedColor`.
	 */
	get textFrame(): TextFrame | null {
		const txBody = firstChild(this.sp, 'p:txBody')
		return txBody ? new TextFrame(txBody, this.part, this.themeContext, undefined, this.relationships) : null
	}

	/** Escape hatch: the underlying `p:sp` element. After mutating it call {@link markDirty}, or `save()` writes the original bytes. */
	get element_(): Element {
		return this.sp
	}

	/** Mark the owning master/layout part dirty so `save()` reserializes it. Call after mutating {@link element_}. */
	markDirty(): void {
		this.part.markDirty()
	}
}

/**
 * A deck's slide master (`slideMasterN.xml`) as a modeled object. Exposes the
 * property tiers a bound slide inherits: the {@link colorMap} (token → theme slot),
 * its {@link theme}, its {@link shapes} and the {@link placeholders} subset of them,
 * its own {@link background}, and the {@link layouts} that build on it. Reachable
 * from {@link SlideLayout.master} and `Slide.master`.
 */
export class SlideMaster implements ShapeHost {
	#themeContext?: ThemeContext

	constructor(
		/** The package the master belongs to, for reaching its layouts, theme, and media. */
		readonly opc: OpcPackage,
		/** The master's OPC part (`/ppt/slideMasters/slideMasterN.xml`). */
		readonly part: Part
	) {}

	/** Partname of the master part. */
	get partName(): string {
		return this.part.partName
	}

	/** This master part's relationships (its layouts, its theme, image embeds, …). */
	get relationships(): Relationships {
		return this.opc.relationshipsFor(this.partName)
	}

	/** The master's authoring name (`p:cSld/@name`), or `''` when unnamed (the writer's default). */
	get name(): string {
		const cSld = this.#cSld()
		return (cSld && attr(cSld, 'name')) ?? ''
	}

	/** The master's theme (via its `theme` relationship), or `null` when absent. */
	get theme(): Theme | null {
		const themePart = this.#relTarget(THEME_REL)
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

	/**
	 * Every shape in the master's `p:cSld/p:spTree`, in document order — the same
	 * `AnyShape` union `Slide.shapes` returns, so shape-walking code applies
	 * unchanged. This is where a template's non-placeholder furniture lives: the
	 * header band, the rule under the title, the logo. Whether a given slide draws
	 * them is `Slide.showMasterSp` (and, for a layout, {@link SlideLayout.showMasterSp}).
	 *
	 * {@link placeholders} is the filtered view of the same tree.
	 */
	get shapes(): AnyShape[] {
		const spTree = this.#spTree()
		return spTree ? buildShapes(spTree, this) : []
	}

	/** The shape anywhere in the master's tree with the given drawing id, or `undefined`. */
	shapeByIdDeep(id: number): AnyShape | undefined {
		return findShapeByIdDeep(this.shapes, id)
	}

	/** The master's placeholder shapes (`p:sp` carrying a `p:ph`), in document order. */
	get placeholders(): Placeholder[] {
		const ctx = this.themeContext()
		const rels = this.relationships
		return placeholderShapes(this.#spTree()).map((sp) => new Placeholder(sp, this.part, ctx, rels))
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
	get background(): SlideBackground | null {
		const bg = backgroundElementOf(this.part.dom.documentElement)
		if (!bg) return null
		const themePart = this.#relTarget(THEME_REL)
		return readSlideBackground(
			bg,
			'master',
			this.themeContext(),
			this.relationships,
			themePart ? this.opc.relationshipsFor(themePart.partName) : null
		)
	}

	/**
	 * The master's colour/font context (its `p:clrMap` + theme), resolved once and
	 * cached. Backs each {@link Placeholder}'s text frame.
	 */
	themeContext(): ThemeContext {
		return (this.#themeContext ??= resolveMasterColorContext(this.opc, this.partName))
	}

	#relTarget(type: string): Part | null {
		const rels = this.relationships
		const rel = rels.byType(type)[0]
		return rel ? (this.opc.part(rels.resolveTarget(rel.id)) ?? null) : null
	}

	#cSld(): Element | null {
		const root = this.part.dom.documentElement
		return root ? firstChild(root, 'p:cSld') : null
	}

	#spTree(): Element | null {
		const cSld = this.#cSld()
		return cSld ? firstChild(cSld, 'p:spTree') : null
	}
}

/**
 * A deck's slide layout (`slideLayoutN.xml`) as a modeled object — the gallery entry
 * a slide binds to. Exposes its {@link name}, its {@link type} (import-only; the
 * writer authors none), its {@link master} (and, through it, the {@link theme}), its
 * {@link shapes} and the {@link placeholders} subset of them, and its own
 * {@link background}. Reachable from `Slide.layout` and {@link SlideMaster.layouts}.
 */
export class SlideLayout implements ShapeHost {
	#themeContext?: ThemeContext

	constructor(
		/** The package the layout belongs to, for reaching its master, theme, and media. */
		readonly opc: OpcPackage,
		/** The layout's OPC part (`/ppt/slideLayouts/slideLayoutN.xml`). */
		readonly part: Part
	) {}

	/** Partname of the layout part. */
	get partName(): string {
		return this.part.partName
	}

	/** This layout part's relationships (its master, image embeds, hyperlinks, …). */
	get relationships(): Relationships {
		return this.opc.relationshipsFor(this.partName)
	}

	/** The layout's authoring name (`p:cSld/@name`, e.g. `Title and Content`), or `''` when unnamed. */
	get name(): string {
		const cSld = this.#cSld()
		return (cSld && attr(cSld, 'name')) ?? ''
	}

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
	get theme(): Theme | null {
		return this.master?.theme ?? null
	}

	/**
	 * Every shape in the layout's `p:cSld/p:spTree`, in document order — the same
	 * `AnyShape` union `Slide.shapes` returns, so shape-walking code applies
	 * unchanged. The write API authors non-placeholder shapes here: every
	 * `defineSlideMaster({ objects })` member that is not a `placeholder` (a `rect`,
	 * a `line`, an `image`, a `chart`, a `text` box) lands in this tree.
	 *
	 * {@link placeholders} is the filtered view of the same tree.
	 */
	get shapes(): AnyShape[] {
		const spTree = this.#spTree()
		return spTree ? buildShapes(spTree, this) : []
	}

	/** The shape anywhere in the layout's tree with the given drawing id, or `undefined`. */
	shapeByIdDeep(id: number): AnyShape | undefined {
		return findShapeByIdDeep(this.shapes, id)
	}

	/** The layout's placeholder shapes (`p:sp` carrying a `p:ph`), in document order. */
	get placeholders(): Placeholder[] {
		const ctx = this.themeContext()
		const rels = this.relationships
		return placeholderShapes(this.#spTree()).map((sp) => new Placeholder(sp, this.part, ctx, rels))
	}

	/**
	 * The layout's *own* background (`p:cSld/p:bg`), decoded, or `null` when it defines
	 * none. A layout usually authors one — an explicit fill (`p:bgPr`) or a
	 * theme-indexed reference (`p:bgRef`). The slide's *effective* background is
	 * `Slide.background`.
	 */
	get background(): SlideBackground | null {
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

	/**
	 * The layout's colour/font context (walked layout → master → theme), resolved once
	 * and cached. Backs each {@link Placeholder}'s text frame.
	 */
	themeContext(): ThemeContext {
		return (this.#themeContext ??= resolveLayoutColorContext(this.opc, this.partName))
	}

	#cSld(): Element | null {
		const root = this.part.dom.documentElement
		return root ? firstChild(root, 'p:cSld') : null
	}

	#spTree(): Element | null {
		const cSld = this.#cSld()
		return cSld ? firstChild(cSld, 'p:spTree') : null
	}
}
