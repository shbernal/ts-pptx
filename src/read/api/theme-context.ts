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
import { applyColorTransforms } from '../oxml/color-transform.js'
import { ELEMENT_NODE, attr, firstChild, intValue, type Element } from '../oxml/dom.js'
import { lstStyleLevelDefRPr, lstStyleLevelFill, parseClrMap, parseClrScheme, placeholderInheritedDefRPrs, placeholderInheritedFill, resolveColor, resolveThemeFont, styleRefFill, styleRefLine, type ColorContext, type FlattenContext } from '../oxml/theme.js'
import type { OpcPackage } from '../opc/package.js'

const SLIDE_LAYOUT_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
const SLIDE_MASTER_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
const THEME_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'

/** The resolved theme subgraph a slide depends on, plus its parsed colour maps. */
export interface SlideThemeParts extends ColorContext {
	slideRoot: Element | null
	layoutPartName: string | null
	masterPartName: string | null
	layoutRoot: Element | null
	masterRoot: Element | null
	/** The theme's `a:themeElements`, for callers that also need `a:fmtScheme`. */
	themeElements: Element | null
}

/** Resolve the single relationship of `type` owned by `partName`, or `null`. */
function resolveSingleRel(opc: OpcPackage, partName: string, type: string): string | null {
	const rels = opc.relationshipsFor(partName)
	const rel = rels.byType(type)[0]
	return rel ? rels.resolveTarget(rel.id) : null
}

/** The document element of a part, or `null` when the partname/part is absent. */
function documentElement(opc: OpcPackage, partName: string | null): Element | null {
	return partName ? (opc.part(partName)?.dom.documentElement ?? null) : null
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

	// A slide's clrMapOvr/overrideClrMapping (if present) wins over the master map.
	const masterClrMap = masterRoot ? firstChild(masterRoot, 'p:clrMap') : null
	const clrMapOvr = slideRoot ? firstChild(slideRoot, 'p:clrMapOvr') : null
	const override = clrMapOvr ? firstChild(clrMapOvr, 'a:overrideClrMapping') : null
	const themeElements = themeRoot ? firstChild(themeRoot, 'a:themeElements') : null

	return {
		slideRoot,
		layoutPartName,
		masterPartName,
		layoutRoot,
		masterRoot,
		themeElements,
		clrMap: parseClrMap(override ?? masterClrMap),
		clrScheme: parseClrScheme(themeElements ? firstChild(themeElements, 'a:clrScheme') : null),
	}
}

/**
 * The colour context a slide's read-model getters resolve against: the
 * {@link ColorContext} maps plus the theme's `a:fmtScheme` (so a colour delivered
 * through a shape's `p:style` `fillRef`/`lnRef` can be resolved like the
 * `theme: 'preserve'` flatten path does). The `fmtScheme` is `null` when the
 * slide's theme is missing.
 */
export function resolveSlideColorContext(opc: OpcPackage, slidePartName: string): FlattenContext {
	const { clrMap, clrScheme, themeElements, layoutRoot, masterRoot } = resolveSlideThemeParts(opc, slidePartName)
	return {
		clrMap,
		clrScheme,
		fmtScheme: themeElements ? firstChild(themeElements, 'a:fmtScheme') : null,
		// The fontScheme lets the run font getters resolve a +mj-*/+mn-* token (the
		// placeholder-inherited typeface chain bottoms out in one) to a literal face.
		fontScheme: themeElements ? firstChild(themeElements, 'a:fontScheme') : null,
		// layout/master roots let the read-model run-colour/size/face getters resolve
		// a placeholder-inherited value the same way the flatten path does.
		layoutRoot,
		masterRoot,
	}
}

/** Identifies a placeholder by its `p:ph` `type`/`idx` for inheritance lookups. */
export interface PlaceholderRef {
	type: string | null
	idx: string
}

/**
 * The colour a placeholder run effectively renders when its own `a:rPr` defines
 * none, resolved to a full {@link ResolvedColor}. Walks the inheritance the way
 * PowerPoint does: the paragraph's `a:pPr/a:defRPr` colour, then the slide text
 * body's `a:lstStyle` colour for the run's `level`, then the placeholder's
 * layout → master → master-`p:txStyles` chain (via {@link placeholderInheritedFill}).
 * The first tier that defines a colour wins. `null` when the run is not in a
 * placeholder, nothing in the chain defines a colour, or it cannot be made literal.
 */
export function resolveInheritedRunColor(ph: PlaceholderRef, level: number, pPr: Element | null, slideLstStyle: Element | null, ctx: FlattenContext): ResolvedColor | null {
	const defRPr = pPr && firstChild(pPr, 'a:defRPr')
	const paraFill = defRPr && firstChild(defRPr, 'a:solidFill')
	if (paraFill) return resolveColorElement(firstChildElement(paraFill), ctx)
	const slideFill = lstStyleLevelFill(slideLstStyle, level)
	if (slideFill) return resolveColorElement(firstChildElement(slideFill), ctx)
	const colorEl = placeholderInheritedFill(ph.type, ph.idx, level, ctx)
	return colorEl ? resolveColorElement(colorEl, ctx) : null
}

/**
 * The `a:defRPr` tiers a placeholder run resolves an inherited character property
 * against, in priority order: the paragraph's `a:pPr/a:defRPr`, then the slide
 * text body's `a:lstStyle` level `a:defRPr`, then the placeholder's layout →
 * master → master-`p:txStyles` chain (via {@link placeholderInheritedDefRPrs}).
 * The first tier that defines the property wins — each property resolves
 * independently, mirroring how the colour resolver walks the same chain.
 */
function inheritedRunDefRPrs(ph: PlaceholderRef, level: number, pPr: Element | null, slideLstStyle: Element | null, ctx: FlattenContext): Element[] {
	const tiers: Element[] = []
	const paraDefRPr = pPr && firstChild(pPr, 'a:defRPr')
	if (paraDefRPr) tiers.push(paraDefRPr)
	const slideDefRPr = lstStyleLevelDefRPr(slideLstStyle, level)
	if (slideDefRPr) tiers.push(slideDefRPr)
	tiers.push(...placeholderInheritedDefRPrs(ph.type, ph.idx, level, ctx))
	return tiers
}

/**
 * The point size a placeholder run effectively renders when its own `a:rPr` sets
 * no `@sz`, walking the inheritance the way PowerPoint does (see
 * {@link inheritedRunDefRPrs}): the first `a:defRPr/@sz` (hundredths of a point)
 * in the paragraph → slide → layout → master → `p:txStyles` chain, converted to
 * points. `null` when the run is not in a placeholder or nothing in the chain
 * defines a size.
 */
export function resolveInheritedRunSize(ph: PlaceholderRef, level: number, pPr: Element | null, slideLstStyle: Element | null, ctx: FlattenContext): number | null {
	for (const defRPr of inheritedRunDefRPrs(ph, level, pPr, slideLstStyle, ctx)) {
		const sz = intValue(attr(defRPr, 'sz'))
		if (sz !== null) return sz / 100
	}
	return null
}

/**
 * The typeface a placeholder run effectively renders when its own `a:rPr` sets no
 * `a:latin`, walking the same chain as {@link resolveInheritedRunSize}: the first
 * `a:defRPr/a:latin/@typeface` it finds, then resolving a `+mj-*`/`+mn-*` theme
 * font token to a literal face name through the theme `fontScheme`. `null` when
 * the run is not in a placeholder, nothing in the chain names a face, or the token
 * cannot be resolved.
 */
export function resolveInheritedRunFontFace(ph: PlaceholderRef, level: number, pPr: Element | null, slideLstStyle: Element | null, ctx: FlattenContext): string | null {
	for (const defRPr of inheritedRunDefRPrs(ph, level, pPr, slideLstStyle, ctx)) {
		const latin = firstChild(defRPr, 'a:latin')
		const typeface = latin && attr(latin, 'typeface')
		if (typeface) return resolveThemeFont(typeface, ctx.fontScheme ?? null)
	}
	return null
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
	transforms: { name: string; value: string | null }[]
	effectiveHex: string
	alpha?: number
}

/**
 * Resolve a DrawingML colour *element* (`a:srgbClr`/`a:schemeClr`/`a:sysClr`)
 * against `ctx` into a full {@link ResolvedColor} — base hex, raw transform list,
 * and the `effectiveHex`/`alpha` after applying those transforms. `null` when the
 * element cannot be made literal (unmapped token, or a colour model we do not
 * resolve). Shared by the solid-fill and gradient-stop colour reads.
 */
export function resolveColorElement(colorEl: Element | null, ctx: ColorContext): ResolvedColor | null {
	const resolved = resolveColor(colorEl, ctx)
	if (!resolved) return null
	const transforms = resolved.transforms.map((t) => ({ name: t.localName, value: attr(t, 'val') }))
	const { hex, alpha } = applyColorTransforms(resolved.hex, transforms)
	return alpha === undefined
		? { hex: resolved.hex, transforms, effectiveHex: hex }
		: { hex: resolved.hex, transforms, effectiveHex: hex, alpha }
}

/** First child *element* of `parent` (skipping text/comment nodes), or `null`. */
function firstChildElement(parent: Element): Element | null {
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) return node as Element
	}
	return null
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
export function resolveStyleFillColor(shape: Element, ctx: FlattenContext): ResolvedColor | null {
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
export function resolveStyleLineColor(shape: Element, ctx: FlattenContext): ResolvedColor | null {
	const style = firstChild(shape, 'p:style')
	const ln = style && styleRefLine(firstChild(style, 'a:lnRef'), ctx)
	return ln ? resolveSolidFillColor(ln, ctx) : null
}
