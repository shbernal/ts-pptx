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
import { ELEMENT_NODE, attr, firstChild, type Element } from '../oxml/dom.js'
import { parseClrMap, parseClrScheme, resolveColor, type ColorContext } from '../oxml/theme.js'
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

/** The {@link ColorContext} subset of a slide's theme parts (for the read-model colour getters). */
export function resolveSlideColorContext(opc: OpcPackage, slidePartName: string): ColorContext {
	const { clrMap, clrScheme } = resolveSlideThemeParts(opc, slidePartName)
	return { clrMap, clrScheme }
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
