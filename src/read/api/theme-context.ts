/**
 * Resolve a slide's theme colour context: the slide → slideLayout → slideMaster
 * → theme walk that turns a `schemeClr` token into a literal hex, plus the
 * `a:solidFill` helper the read-model colour getters share.
 *
 * One implementation backs two callers — the read-model getters
 * (`Slide.themeContext` → `Shape.resolvedFill` / `Run.resolvedColor`) and the
 * `importSlide` `theme: 'preserve'` flatten path, which layers its `fmtScheme` /
 * background needs on top of {@link resolveSlideThemeParts} — so a token resolves
 * identically whether it is read or baked.
 */
import { applyColorTransforms, type ColorTransform } from '../oxml/color-transform.js'
import { attr, boolValue, firstChild, firstChildElement, numberValue, type Element } from '../oxml/dom.js'
import {
	parseClrMap,
	parseClrScheme,
	resolveColor,
	resolveThemeFont,
	styleRefFill,
	styleRefLine,
	themeElementsOfRoot,
	type ColorContext,
	type ThemeContext,
} from '../oxml/theme.js'
import {
	lstStyleLevelDefRPr,
	lstStyleLevelFill,
	placeholderInheritedAnchor,
	placeholderInheritedDefRPrs,
	placeholderInheritedFill,
	placeholderInheritedXfrm,
} from '../oxml/placeholder-inherit.js'
import type { OpcPackage } from '../opc/package.js'
import { ptFromHundredths } from './coords.js'
import { readBox } from './shapes/geometry.js'
import { resolveSingleRel } from '../opc/partnames.js'
import {
	NOTES_MASTER_REL,
	OFFICE_DOCUMENT_REL,
	SLIDE_LAYOUT_REL,
	SLIDE_MASTER_REL,
	THEME_REL,
} from '../../ooxml/rel-types.js'

/** The resolved theme subgraph a slide depends on, plus its parsed colour maps. */
export interface SlideThemeParts extends ColorContext {
	slideRoot: Element | null
	layoutPartName: string | null
	masterPartName: string | null
	layoutRoot: Element | null
	masterRoot: Element | null
	/**
	 * Partname of the theme part the chain bottoms out in, or `null`. Callers that
	 * materialize a fill out of the theme (a `p:bgRef`'s `fmtScheme` entry) need it to
	 * resolve that fill's relationship ids against the *theme's* rels.
	 */
	themePartName: string | null
	/** The theme's `a:themeElements`, for callers that also need `a:fmtScheme`. */
	themeElements: Element | null
}

/** The document element of a part, or `null` when the partname/part is absent. */
function documentElement(opc: OpcPackage, partName: string | null): Element | null {
	return partName ? (opc.part(partName)?.dom.documentElement ?? null) : null
}

/**
 * The presentation's `p:defaultTextStyle` (`presentation.xml`, reached via the
 * package `officeDocument` relationship), or `null` when absent. This is
 * PowerPoint's lowest-priority text fallback — the bottom tier of a slide run's
 * size/colour/face/bold resolution (see `ThemeContext.defaultTextStyle`).
 */
function presentationDefaultTextStyle(opc: OpcPackage): Element | null {
	const root = documentElement(opc, resolveSingleRel(opc, '/', OFFICE_DOCUMENT_REL))
	return root ? firstChild(root, 'p:defaultTextStyle') : null
}

/**
 * Walk slide → slideLayout → slideMaster → theme, returning the resolved part
 * roots plus the parsed colour map (`clrMap`, honouring the slide's `clrMapOvr`)
 * and colour scheme (`clrScheme`). A missing link degrades to `null` roots and
 * empty maps rather than throwing.
 */
export function resolveSlideThemeParts(opc: OpcPackage, slidePartName: string): SlideThemeParts {
	const layoutPartName = resolveSingleRel(opc, slidePartName, SLIDE_LAYOUT_REL)
	const masterPartName = layoutPartName ? resolveSingleRel(opc, layoutPartName, SLIDE_MASTER_REL) : null
	const themePartName = masterPartName ? resolveSingleRel(opc, masterPartName, THEME_REL) : null

	const slideRoot = documentElement(opc, slidePartName)
	const layoutRoot = documentElement(opc, layoutPartName)
	const masterRoot = documentElement(opc, masterPartName)
	const themeRoot = documentElement(opc, themePartName)

	const themeElements = themeElementsOfRoot(themeRoot)

	return {
		slideRoot,
		layoutPartName,
		masterPartName,
		layoutRoot,
		masterRoot,
		themePartName,
		themeElements,
		clrMap: parseClrMap(effectiveClrMap(slideRoot, masterRoot)),
		clrScheme: themeTier(themeElements).clrScheme,
	}
}

/**
 * The colour context a slide's read-model getters resolve against: the
 * {@link ColorContext} maps plus the theme's `a:fmtScheme` (so a colour delivered
 * through a shape's `p:style` `fillRef`/`lnRef` can be resolved like the
 * `theme: 'preserve'` flatten path does). The `fmtScheme` is `null` when the
 * slide's theme is missing.
 */
export function resolveSlideColorContext(opc: OpcPackage, slidePartName: string): ThemeContext {
	const { clrMap, themeElements, layoutRoot, masterRoot } = resolveSlideThemeParts(opc, slidePartName)
	return {
		clrMap,
		// `fontScheme` is what lets the run font getters resolve a +mj-*/+mn-* token (the
		// placeholder-inherited typeface chain bottoms out in one) to a literal face.
		...themeTier(themeElements),
		// layout/master roots let the read-model run-colour/size/face getters resolve
		// a placeholder-inherited value the same way the flatten path does.
		layoutRoot,
		masterRoot,
		// The presentation's default text style — the lowest-priority tier a slide run
		// (placeholder or not) falls back to when it resolves nothing above it.
		defaultTextStyle: presentationDefaultTextStyle(opc),
	}
}

/**
 * The colour context a slide's *speaker-notes* getters resolve against. A notes
 * slide inherits its theme through its notesMaster (`notesSlideN.xml.rels`
 * carries the `notesMaster` relationship), which in turn points at the notes
 * theme (`theme2.xml`). Walk notesSlide → notesMaster → theme, taking `clrMap`
 * from the notesMaster's own `p:clrMap` and `clrScheme`/`fontScheme` from that
 * theme, so a notes run's own `schemeClr` fill resolves to a literal hex the same
 * way a slide run's does (backing `Run.resolvedColor` on {@link Slide.notesTextFrame}).
 *
 * Notes runs inherit character properties from the notesMaster's `p:notesStyle`,
 * not from a slide layout/master placeholder chain, so `layoutRoot`/`masterRoot`
 * are deliberately left absent. Instead the notesMaster's `p:notesStyle` is carried
 * as `notesStyle`, and the notes body `TextFrame` is built with a placeholder
 * context (see {@link import('./notes.js').NotesPlaceholder}) so a body run's
 * effective size/face/bold (and inherited colour) resolve against it — the notes
 * analogue of the slide placeholder chain. The maps are empty when the
 * notesMaster/theme chain is incomplete, in which case tokens stay unresolved.
 */
export function resolveNotesColorContext(opc: OpcPackage, notesPartName: string): ThemeContext {
	const masterPartName = resolveSingleRel(opc, notesPartName, NOTES_MASTER_REL)
	const themePartName = masterPartName ? resolveSingleRel(opc, masterPartName, THEME_REL) : null
	const masterRoot = documentElement(opc, masterPartName)
	const themeRoot = documentElement(opc, themePartName)
	const themeElements = themeElementsOfRoot(themeRoot)
	return {
		clrMap: parseClrMap(effectiveClrMap(masterRoot, masterRoot)),
		// `fontScheme` is what lets a notes run resolve a +mj-*/+mn-* theme-font token to a
		// literal face; the writer's notesStyle uses +mn-lt, so an authored notes run that
		// omits its own face resolves through here.
		...themeTier(themeElements),
		// The notesMaster's text style — the tier a notes-body run's inherited
		// size/face/bold/colour resolves against (see `ThemeContext.notesStyle`).
		notesStyle: masterRoot ? firstChild(masterRoot, 'p:notesStyle') : null,
	}
}

/**
 * The colour context a *slide master*'s own getters resolve against: the master's
 * `p:clrMap` plus its theme's `clrScheme`/`fmtScheme`/`fontScheme` (walked master →
 * theme). Backs the text frames of {@link import('./chrome.js').SlideMaster}'s
 * placeholders — a master placeholder run's own `schemeClr` resolves to a literal
 * hex the same way a slide run's does. `masterRoot` is carried so an inherited
 * placeholder value still resolves against the master's own text styles. The maps
 * are empty when the theme is missing, in which case tokens stay unresolved.
 */
export function resolveMasterColorContext(opc: OpcPackage, masterPartName: string): ThemeContext {
	const themePartName = resolveSingleRel(opc, masterPartName, THEME_REL)
	const masterRoot = documentElement(opc, masterPartName)
	const themeElements = themeElementsOf(opc, themePartName)
	return {
		clrMap: parseClrMap(effectiveClrMap(masterRoot, masterRoot)),
		...themeTier(themeElements),
		masterRoot,
	}
}

/**
 * The colour context a *slide layout*'s own getters resolve against: walk layout →
 * master → theme. The effective `clrMap` is the layout's own
 * `p:clrMapOvr/a:overrideClrMapping` when present (rare), else the master's
 * `p:clrMap` — the layout's usual `a:masterClrMapping` means "inherit the master
 * map". `clrScheme`/`fmtScheme`/`fontScheme` come from the master's theme, and both
 * `layoutRoot`/`masterRoot` are carried for inherited-placeholder resolution. Backs
 * the text frames of {@link import('./chrome.js').SlideLayout}'s placeholders.
 */
export function resolveLayoutColorContext(opc: OpcPackage, layoutPartName: string): ThemeContext {
	const masterPartName = resolveSingleRel(opc, layoutPartName, SLIDE_MASTER_REL)
	const themePartName = masterPartName ? resolveSingleRel(opc, masterPartName, THEME_REL) : null
	const layoutRoot = documentElement(opc, layoutPartName)
	const masterRoot = documentElement(opc, masterPartName)
	const themeElements = themeElementsOf(opc, themePartName)
	return {
		clrMap: parseClrMap(effectiveClrMap(layoutRoot, masterRoot)),
		...themeTier(themeElements),
		layoutRoot,
		masterRoot,
	}
}

/**
 * The three theme tiers every colour context carries, read off one `a:themeElements`.
 *
 * `clrScheme` is parsed into the token map the resolvers read; `fmtScheme` and `fontScheme`
 * ride as elements, because a style ref and a `+mj-*`/`+mn-*` token are resolved against the
 * live DOM rather than a snapshot. Four builders spelled these three lookups out.
 */
function themeTier(themeElements: Element | null): Pick<ThemeContext, 'clrScheme' | 'fmtScheme' | 'fontScheme'> {
	return {
		clrScheme: parseClrScheme(themeElements ? firstChild(themeElements, 'a:clrScheme') : null),
		fmtScheme: themeElements ? firstChild(themeElements, 'a:fmtScheme') : null,
		fontScheme: themeElements ? firstChild(themeElements, 'a:fontScheme') : null,
	}
}

/**
 * The colour map in force for `ownerRoot`: its own `p:clrMapOvr/a:overrideClrMapping` when it
 * has one, else the master's `p:clrMap`.
 *
 * A slide's and a layout's usual `a:masterClrMapping` means "inherit the master map", which is
 * why an absent override falls through rather than mapping nothing. A master is its own owner
 * and has no override, so it passes its root as both.
 */
function effectiveClrMap(ownerRoot: Element | null, masterRoot: Element | null): Element | null {
	const clrMapOvr = ownerRoot ? firstChild(ownerRoot, 'p:clrMapOvr') : null
	const override = clrMapOvr ? firstChild(clrMapOvr, 'a:overrideClrMapping') : null
	return override ?? (masterRoot ? firstChild(masterRoot, 'p:clrMap') : null)
}

/** The `a:themeElements` of a theme part (by partname), or `null` when the part/element is absent. */
function themeElementsOf(opc: OpcPackage, themePartName: string | null): Element | null {
	return themeElementsOfRoot(documentElement(opc, themePartName))
}

/** Identifies a placeholder by its `p:ph` `type`/`idx` for inheritance lookups. */
export interface PlaceholderRef {
	type: string | null
	idx: string
}

/**
 * The colour a run effectively renders when its own `a:rPr` defines none,
 * resolved to a full {@link ResolvedColor}. Walks the inheritance the way
 * PowerPoint does: the paragraph's `a:pPr/a:defRPr` colour, then the slide text
 * body's `a:lstStyle` colour for the run's `level`, then — only for a placeholder
 * run — the placeholder's layout → master → master-`p:txStyles` chain (via
 * {@link placeholderInheritedFill}), then the presentation's `p:defaultTextStyle`
 * for the run's level (`ctx.defaultTextStyle`). The first tier that defines a
 * colour wins. `ph` is `null` for a non-placeholder run, which skips the
 * placeholder chain but still reaches `p:defaultTextStyle`. `null` when nothing in
 * the chain defines a colour, or it cannot be made literal.
 */
export function resolveInheritedRunColor(
	ph: PlaceholderRef | null,
	level: number,
	pPr: Element | null,
	slideLstStyle: Element | null,
	ctx: ThemeContext
): ResolvedColor | null {
	const defRPr = pPr && firstChild(pPr, 'a:defRPr')
	const paraFill = defRPr && firstChild(defRPr, 'a:solidFill')
	if (paraFill) return resolveColorElement(firstChildElement(paraFill), ctx)
	const slideFill = lstStyleLevelFill(slideLstStyle, level)
	if (slideFill) return resolveColorElement(firstChildElement(slideFill), ctx)
	if (ph) {
		const colorEl = placeholderInheritedFill(ph.type, ph.idx, level, ctx)
		if (colorEl) return resolveColorElement(colorEl, ctx)
	}
	const defaultFill = lstStyleLevelFill(ctx.defaultTextStyle ?? null, level)
	return defaultFill ? resolveColorElement(firstChildElement(defaultFill), ctx) : null
}

/**
 * The `a:defRPr` tiers a run resolves an inherited character property against, in
 * priority order: the paragraph's `a:pPr/a:defRPr`, then the slide text body's
 * `a:lstStyle` level `a:defRPr`, then — only for a placeholder run — the
 * placeholder's layout → master → master-`p:txStyles` chain (via
 * {@link placeholderInheritedDefRPrs}), then the presentation's
 * `p:defaultTextStyle` level `a:defRPr` (`ctx.defaultTextStyle`). The first tier
 * that defines the property wins — each property resolves independently, mirroring
 * how the colour resolver walks the same chain. `ph` is `null` for a
 * non-placeholder run, which skips the placeholder chain but still reaches
 * `p:defaultTextStyle`.
 */
function inheritedRunDefRPrs(
	ph: PlaceholderRef | null,
	level: number,
	pPr: Element | null,
	slideLstStyle: Element | null,
	ctx: ThemeContext
): Element[] {
	const tiers: Element[] = []
	const paraDefRPr = pPr && firstChild(pPr, 'a:defRPr')
	if (paraDefRPr) tiers.push(paraDefRPr)
	const slideDefRPr = lstStyleLevelDefRPr(slideLstStyle, level)
	if (slideDefRPr) tiers.push(slideDefRPr)
	if (ph) tiers.push(...placeholderInheritedDefRPrs(ph.type, ph.idx, level, ctx))
	const defaultDefRPr = lstStyleLevelDefRPr(ctx.defaultTextStyle ?? null, level)
	if (defaultDefRPr) tiers.push(defaultDefRPr)
	return tiers
}

/**
 * The point size a run effectively renders when its own `a:rPr` sets no `@sz`,
 * walking the inheritance the way PowerPoint does (see {@link inheritedRunDefRPrs}):
 * the first `a:defRPr/@sz` (hundredths of a point) in the paragraph → slide →
 * (placeholder) layout → master → `p:txStyles` → `p:defaultTextStyle` chain,
 * converted to points. A non-placeholder run skips the placeholder tiers but still
 * reaches `p:defaultTextStyle`. `null` when nothing in the chain defines a size.
 */
export function resolveInheritedRunSize(
	ph: PlaceholderRef | null,
	level: number,
	pPr: Element | null,
	slideLstStyle: Element | null,
	ctx: ThemeContext
): number | null {
	for (const defRPr of inheritedRunDefRPrs(ph, level, pPr, slideLstStyle, ctx)) {
		const sz = numberValue(attr(defRPr, 'sz'))
		if (sz !== null) return ptFromHundredths(sz)
	}
	return null
}

/**
 * The typeface a run effectively renders when its own `a:rPr` sets no `a:latin`,
 * walking the same chain as {@link resolveInheritedRunSize}: the first
 * `a:defRPr/a:latin/@typeface` it finds, then resolving a `+mj-*`/`+mn-*` theme
 * font token to a literal face name through the theme `fontScheme`. A
 * non-placeholder run skips the placeholder tiers but still reaches
 * `p:defaultTextStyle` (whose level `a:latin` is `+mn-lt` in a PowerPoint-written
 * deck). `null` when nothing in the chain names a face, or the token cannot be
 * resolved.
 */
export function resolveInheritedRunFontFace(
	ph: PlaceholderRef | null,
	level: number,
	pPr: Element | null,
	slideLstStyle: Element | null,
	ctx: ThemeContext
): string | null {
	for (const defRPr of inheritedRunDefRPrs(ph, level, pPr, slideLstStyle, ctx)) {
		const latin = firstChild(defRPr, 'a:latin')
		const typeface = latin && attr(latin, 'typeface')
		if (typeface) return resolveThemeFont(typeface, ctx.fontScheme ?? null)
	}
	return null
}

/**
 * Whether a run effectively renders with the boolean character property `name`
 * (`b`/`i`) when its own `a:rPr` sets none, walking the same chain as
 * {@link resolveInheritedRunSize}: the first `a:defRPr/@<name>` in the paragraph →
 * slide → (placeholder) layout → master → `p:txStyles` → `p:defaultTextStyle`
 * chain. The first tier that *states* the attribute wins, and its value is parsed
 * by {@link boolValue} — the same `xsd:boolean` parser the run's own `bold`/`italic`
 * getters use — so an unparseable value reports `null` rather than being read as
 * `false`. `null` when nothing in the chain states it.
 */
function resolveInheritedRunFlag(
	name: string,
	ph: PlaceholderRef | null,
	level: number,
	pPr: Element | null,
	slideLstStyle: Element | null,
	ctx: ThemeContext
): boolean | null {
	for (const defRPr of inheritedRunDefRPrs(ph, level, pPr, slideLstStyle, ctx)) {
		const raw = attr(defRPr, name)
		if (raw !== null) return boolValue(raw)
	}
	return null
}

/**
 * Whether a run effectively renders bold when its own `a:rPr` sets no `@b`: the
 * first `a:defRPr/@b` in the inheritance chain (see {@link resolveInheritedRunFlag}).
 * A non-placeholder run skips the placeholder tiers but still reaches
 * `p:defaultTextStyle` (which sets no `@b` in a PowerPoint-written deck, so bold
 * stays `null` there). `null` when nothing in the chain defines bold.
 */
export function resolveInheritedRunBold(
	ph: PlaceholderRef | null,
	level: number,
	pPr: Element | null,
	slideLstStyle: Element | null,
	ctx: ThemeContext
): boolean | null {
	return resolveInheritedRunFlag('b', ph, level, pPr, slideLstStyle, ctx)
}

/**
 * Whether a run effectively renders italic when its own `a:rPr` sets no `@i`: the
 * first `a:defRPr/@i` in the inheritance chain (see {@link resolveInheritedRunFlag}).
 * `@b` and `@i` are siblings on `CT_TextCharacterProperties` and a master text
 * style states them together, so this is the exact twin of
 * {@link resolveInheritedRunBold} one attribute along. `null` when nothing in the
 * chain defines italic.
 */
export function resolveInheritedRunItalic(
	ph: PlaceholderRef | null,
	level: number,
	pPr: Element | null,
	slideLstStyle: Element | null,
	ctx: ThemeContext
): boolean | null {
	return resolveInheritedRunFlag('i', ph, level, pPr, slideLstStyle, ctx)
}

/**
 * The vertical anchor a placeholder text frame effectively renders when its own
 * `a:bodyPr` sets no `@anchor`: the value inherited from the layout → master
 * placeholder `a:bodyPr` (see {@link placeholderInheritedAnchor}). `null` when the
 * frame is not in a placeholder or nothing in the chain sets an anchor (PowerPoint
 * then defaults to top).
 */
export function resolveInheritedAnchor(ph: PlaceholderRef, ctx: ThemeContext): string | null {
	return placeholderInheritedAnchor(ph.type, ph.idx, ctx)
}

/** Which tier of the slide → layout → master chain a {@link ResolvedFrame} resolved from. */
export type GeometrySource = 'own' | 'layout' | 'master'

/** A shape's effective position and size in EMU, tagged with where it resolved from. */
export interface ResolvedFrame {
	left: number
	top: number
	width: number
	height: number
	source: GeometrySource
}

/**
 * The position/size a placeholder inherits when its own `a:spPr` carries no
 * `a:xfrm`: the matching layout placeholder's geometry, else the master's (see
 * {@link placeholderInheritedXfrm}). `null` when the chain has no matching
 * placeholder geometry to inherit, or the matched `a:xfrm` is missing an
 * `a:off`/`a:ext` coordinate. Callers that also need the "own `a:xfrm` wins"
 * check (the common case) use {@link Shape.resolvedFrame}, which layers that on
 * top of this.
 */
export function resolveInheritedFrame(ph: PlaceholderRef, ctx: ThemeContext): ResolvedFrame | null {
	const found = placeholderInheritedXfrm(ph.type, ph.idx, ctx)
	if (!found) return null
	// `readBox` is the same read `absoluteFrame` and `GroupShape.childFrame` already use, and it
	// returns `null` for an incomplete box, which is what the four unreachable `=== undefined`
	// arms here were guarding against.
	const box = readBox(found.xfrm, 'a:off', 'a:ext')
	if (!box) return null
	return { left: box.x, top: box.y, width: box.cx, height: box.cy, source: found.source }
}

/**
 * A DrawingML colour reference resolved against a slide's theme to a literal hex.
 *
 * `hex` is the **base** token colour and `transforms` reports the colour-transform
 * children (`lumMod`/`lumOff`/`shade`/`tint`/`alpha`/…) in document order as
 * `{ name, value }` pairs — both kept for traceability and for the
 * `theme: 'preserve'` flatten path that re-emits the transforms verbatim.
 *
 * `effectiveHex` is the colour a renderer actually paints: `hex` with its
 * `transforms` applied (see {@link applyColorTransforms}). Read this for the final
 * rendered colour. `alpha` (0–1) is present only when an `alpha*` transform set an
 * opacity.
 */
export interface ResolvedColor {
	hex: string
	transforms: ColorTransform[]
	effectiveHex: string
	alpha?: number
}

/**
 * Resolve a DrawingML colour *element* (every `a:EG_ColorChoice` member except
 * `a:scrgbClr` — see `read/oxml/theme.ts`) against `ctx` into a full
 * {@link ResolvedColor} — base hex, raw transform list,
 * and the `effectiveHex`/`alpha` after applying those transforms. `null` when the
 * element cannot be made literal (unmapped token, or a colour model we do not
 * resolve). Shared by the solid-fill and gradient-stop colour reads.
 */
export function resolveColorElement(colorEl: Element | null, ctx: ColorContext): ResolvedColor | null {
	const resolved = resolveColor(colorEl, ctx)
	if (!resolved) return null
	const transforms = resolved.transforms.map((t) => ({ name: t.localName ?? '', value: attr(t, 'val') }))
	const { hex, alpha } = applyColorTransforms(resolved.hex, transforms)
	return alpha === undefined
		? { hex: resolved.hex, transforms, effectiveHex: hex }
		: { hex: resolved.hex, transforms, effectiveHex: hex, alpha }
}

/**
 * Resolve the `a:solidFill` colour of a properties container (`p:spPr`,
 * `p:grpSpPr`, `a:ln`, or a run's `a:rPr`) to a literal hex through `ctx`.
 * `null` when the container has no solid fill, or the colour cannot be made
 * literal (an unmapped token, or a colour model we do not resolve).
 */
export function resolveSolidFillColor(container: Element | null, ctx: ColorContext): ResolvedColor | null {
	if (!container) return null
	const solidFill = firstChild(container, 'a:solidFill')
	if (!solidFill) return null
	return resolveColorElement(firstChildElement(solidFill), ctx)
}

/**
 * Resolve the fill colour a shape inherits from its `p:style` `a:fillRef`
 * (style-matrix fill) to a literal hex through `ctx`. Used as the fallback for
 * {@link import('./shapes.js').Shape.resolvedFill} when the shape carries no
 * explicit `spPr` fill choice. `null` when there is no `fillRef`, it cannot be
 * resolved, or the indexed style entry is not a solid fill (a gradient style fill
 * has no single colour — read it through `gradientStops` instead).
 */
export function resolveStyleFillColor(shape: Element, ctx: ThemeContext): ResolvedColor | null {
	const style = firstChild(shape, 'p:style')
	const fill = style && styleRefFill(firstChild(style, 'a:fillRef'), ctx)
	if (!fill || fill.localName !== 'solidFill') return null
	return resolveColorElement(firstChildElement(fill), ctx)
}

/**
 * Resolve the line colour a shape inherits from its `p:style` `a:lnRef`
 * (style-matrix line) to a literal hex through `ctx`. Used as the fallback for
 * {@link import('./shapes.js').Shape.resolvedLine} when the shape carries no
 * explicit `spPr/a:ln`. `null` when there is no `lnRef` or it cannot be resolved.
 */
export function resolveStyleLineColor(shape: Element, ctx: ThemeContext): ResolvedColor | null {
	const style = firstChild(shape, 'p:style')
	const ln = style && styleRefLine(firstChild(style, 'a:lnRef'), ctx)
	return ln ? resolveSolidFillColor(ln, ctx) : null
}

/** A shape's resolved `p:style/a:fontRef` text tier: its colour and theme face. */
export interface StyleFontRef {
	/** The `a:fontRef` child colour (`a:schemeClr`/`a:srgbClr`/…) resolved through the theme, or `null` when it names none / cannot be made literal. */
	color: ResolvedColor | null
	/** The face named by `a:fontRef/@idx` (`major`→`+mj-lt`, `minor`→`+mn-lt`) resolved through the theme `fontScheme`, or `null` for `idx="none"` / an unresolvable token. */
	face: string | null
}

/**
 * Resolve the text colour and typeface a shape derives from its `p:style/a:fontRef`
 * (the style-matrix font reference) — the tier PowerPoint applies to a shape's runs
 * just below their own `a:rPr` and above the placeholder/`p:defaultTextStyle` chain.
 * The `a:fontRef` child colour resolves through `ctx` like any other (it can carry
 * `lumMod`/`shade` transforms); `@idx` (`major`|`minor`|`none`) maps to the theme's
 * major/minor Latin font. `null` when the shape has no `p:style/a:fontRef` at all.
 */
export function resolveStyleFontRef(shape: Element, ctx: ThemeContext): StyleFontRef | null {
	const style = firstChild(shape, 'p:style')
	const fontRef = style && firstChild(style, 'a:fontRef')
	if (!fontRef) return null
	const color = resolveColorElement(firstChildElement(fontRef), ctx)
	const idx = attr(fontRef, 'idx')
	const token = idx === 'major' ? '+mj-lt' : idx === 'minor' ? '+mn-lt' : null
	return { color, face: token ? resolveThemeFont(token, ctx.fontScheme ?? null) : null }
}
