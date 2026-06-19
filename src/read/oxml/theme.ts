/**
 * Theme resolution for the {@link Presentation.importSlide} `theme: 'preserve'`
 * mode: bake the values a slide's *source* theme would have produced into the
 * slide XML, so the slide no longer depends on which theme it resolves against.
 *
 * Two kinds of theme dependency are flattened (see IMPORT_SLIDE_THEME_PRESERVE):
 *
 * 1. **`a:schemeClr` tokens** — a scheme token (`accent1`, `bg1`, …) routes
 *    through the master's `clrMap` to a `clrScheme` slot, whose literal RGB we
 *    emit as `a:srgbClr`. Child colour transforms (`lumMod`/`shade`/`alpha`/…)
 *    are carried through untouched, so tints/shades render identically — we swap
 *    only the *base* colour reference, never compute the transform math.
 * 2. **Style-matrix refs** — a shape's `p:style` (`lnRef`/`fillRef`/`effectRef`)
 *    indexes the theme's `fmtScheme`; we resolve the indexed entry into an
 *    explicit `spPr` fill/line/effect (substituting its `phClr` placeholder with
 *    the ref's own colour) and neutralize the ref so it cannot re-resolve. The
 *    `fontRef` is left intact so its font (and resolved colour) can re-bind to
 *    the destination theme — the deliberate "normalize fonts on attach" bonus.
 * 3. **Placeholder-inherited run colour** — a run whose `a:rPr` carries no own
 *    fill takes its colour from the source placeholder → layout → master text
 *    style chain (`p:txStyles` / placeholder `a:lstStyle`). Rebinding to the
 *    destination master would replace that chain and flip the colour, so we
 *    resolve each such run's effective colour from the *source* styles and write
 *    it explicitly onto the run.
 * 4. **Placeholder-inherited geometry & run size** — a placeholder shape with no
 *    own `a:xfrm` takes its position/size from the matching source layout/master
 *    placeholder, and a run with no own `sz`/`b`/`i` takes them from the same text
 *    style chain. Rebinding to the destination master replaces both inheritances,
 *    so a title clips off-canvas and type comes out at the wrong size. We bake the
 *    effective `a:xfrm` onto the shape and the effective `sz`/`b`/`i` onto each run.
 *    Typeface (`a:latin`) is deliberately *not* baked — it re-binds to the
 *    destination theme along with `fontRef`.
 *
 * All functions operate purely on DOM elements; the caller gathers the source
 * `clrMap` / `clrScheme` / `fmtScheme` parts (and the source layout/master roots
 * for the placeholder inheritance passes) and owns marking the slide dirty.
 */
import { ELEMENT_NODE, OOXML_NS, attr, createElement, firstChild, getElements, getOrAddChild, insertInOrder, intValue, removeChildrenByQName, setAttr, type Element } from './dom.js'
import { FILL_CHOICES } from './fill.js'

/** The 12 `a:clrScheme` slot names, in schema order. */
const SCHEME_SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
/** Scheme tokens that name a `clrScheme` slot directly, bypassing the `clrMap`. */
const DIRECT_SLOT_TOKENS = new Set(['dk1', 'lt1', 'dk2', 'lt2'])

/** Schema successors for ordered insertion into `p:spPr` (CT_ShapeProperties). */
const SPPR_XFRM_AFTER = ['a:custGeom', 'a:prstGeom', ...FILL_CHOICES, 'a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
const SPPR_FILL_AFTER = ['a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
const SPPR_LN_AFTER = ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
const SPPR_EFFECT_AFTER = ['a:scene3d', 'a:sp3d', 'a:extLst']
const SHAPE_AFTER_SPPR = ['p:style', 'p:txBody']

/**
 * The colour-resolution context: the two maps that turn a DrawingML colour
 * reference into a literal hex — the effective colour map (token → `clrScheme`
 * slot, honouring any slide `clrMapOvr`) and the resolved colour scheme (slot →
 * 6-hex RGB). Shared by the read-model colour getters and {@link FlattenContext}.
 */
export interface ColorContext {
	clrMap: Map<string, string>
	clrScheme: Map<string, string>
}

/**
 * Everything {@link flattenSlide} needs from the slide's source theme subgraph:
 * the {@link ColorContext} maps plus the live `a:fmtScheme` element (for
 * style-matrix resolution).
 */
export interface FlattenContext extends ColorContext {
	fmtScheme: Element | null
	/**
	 * The slide's effective background inherited from the *source*
	 * `slideLayout`/`slideMaster` (the raw `p:bg` element), or `null` when the
	 * slide carries its own. Applied onto the slide before flattening so the
	 * background survives rebinding to the destination master. See
	 * {@link flattenSlide}.
	 */
	inheritedBackground?: Element | null
	/**
	 * The source `slideLayout` root element, for resolving placeholder-inherited
	 * run colours (gap 1). Read-only — never mutated. `null`/absent disables the
	 * placeholder run-colour pass for the layout tier.
	 */
	layoutRoot?: Element | null
	/**
	 * The source `slideMaster` root element, for resolving placeholder-inherited
	 * run colours via its placeholder `a:lstStyle` and `p:txStyles`. Read-only —
	 * never mutated.
	 */
	masterRoot?: Element | null
}

/** Parse an `a:clrScheme` into slot → 6-hex RGB, reading `srgbClr`/`sysClr`. */
export function parseClrScheme(clrScheme: Element | null): Map<string, string> {
	const out = new Map<string, string>()
	if (!clrScheme) return out
	for (const slot of SCHEME_SLOTS) {
		const slotEl = firstChild(clrScheme, `a:${slot}`)
		const hex = slotEl && colorElementHex(firstChildElement(slotEl))
		if (hex) out.set(slot, hex)
	}
	return out
}

/**
 * Parse the `bg1`/`tx1`/`bg2`/`tx2`/`accent*`/`hlink`/`folHlink` attributes of an
 * `a:clrMap` (or `a:overrideClrMapping`) into token → slot.
 */
export function parseClrMap(clrMap: Element | null): Map<string, string> {
	const out = new Map<string, string>()
	if (!clrMap) return out
	for (const token of ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) {
		const slot = attr(clrMap, token)
		if (slot) out.set(token, slot)
	}
	return out
}

/** Resolve a scheme token (`accent1`, `bg1`, `dk1`, …) to a 6-hex RGB, or `null`. */
function resolveSchemeToken(token: string, ctx: ColorContext): string | null {
	if (token === 'phClr') return null
	const slot = DIRECT_SLOT_TOKENS.has(token) ? token : ctx.clrMap.get(token)
	return slot ? (ctx.clrScheme.get(slot) ?? null) : null
}

/** First child *element* of `parent` (skipping text/comment nodes), or `null`. */
function firstChildElement(parent: Element): Element | null {
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) return node as Element
	}
	return null
}

/** Direct child *elements* of `parent`, in order. */
function childElements(parent: Element): Element[] {
	const out: Element[] = []
	for (let node = parent.firstChild; node; node = node.nextSibling) {
		if (node.nodeType === ELEMENT_NODE) out.push(node as Element)
	}
	return out
}

/** Whether `el` is a DrawingML element with the given local name. */
function isA(el: Element | null, local: string): boolean {
	return !!el && el.namespaceURI === OOXML_NS.a && el.localName === local
}

/** The literal RGB of a colour element (`srgbClr`/`sysClr`), or `null` for others. */
function colorElementHex(color: Element | null): string | null {
	if (!color) return null
	if (isA(color, 'srgbClr')) return attr(color, 'val')
	if (isA(color, 'sysClr')) return attr(color, 'lastClr') ?? attr(color, 'val')
	return null
}

/** A colour reference resolved to a literal base RGB plus its transform children. */
export interface ResolvedColor {
	hex: string
	transforms: Element[]
}

/**
 * Resolve a DrawingML colour element to a literal `{ hex, transforms }`, routing
 * `a:schemeClr` through the context. Returns `null` when the base cannot be made
 * literal (unmapped token, or a colour model we do not flatten).
 */
export function resolveColor(color: Element | null, ctx: ColorContext): ResolvedColor | null {
	if (!color) return null
	const transforms = childElements(color)
	if (isA(color, 'srgbClr')) {
		const hex = attr(color, 'val')
		return hex ? { hex, transforms } : null
	}
	if (isA(color, 'sysClr')) {
		const hex = colorElementHex(color)
		return hex ? { hex, transforms } : null
	}
	if (isA(color, 'schemeClr')) {
		const token = attr(color, 'val')
		const hex = token ? resolveSchemeToken(token, ctx) : null
		return hex ? { hex, transforms } : null
	}
	return null
}

/**
 * Flatten one slide's theme dependencies in place:
 *
 * 1. carry the background the slide inherited from its source layout/master onto
 *    the slide (so it survives rebinding to a different master);
 * 2. materialize `p:bgRef` and style-matrix refs into explicit fills/lines/effects;
 * 3. bake each placeholder's effective geometry (inherited `a:xfrm`) onto the
 *    shape so a rebind cannot move or resize it;
 * 4. bake each placeholder run's effective colour and size/weight (inherited from
 *    the source layout/master text styles) explicitly onto the run;
 * 5. rewrite every remaining `a:schemeClr` to its literal `a:srgbClr`.
 *
 * Steps run in this order so the inherited/materialized backgrounds are present
 * before the final scheme-colour sweep resolves the colours they carry. The
 * placeholder geometry/colour/size passes have no data dependency on the others.
 * The caller marks the part dirty.
 */
export function flattenSlide(slideRoot: Element, ctx: FlattenContext): void {
	applyInheritedBackground(slideRoot, ctx)
	materializeBackground(slideRoot, ctx)
	materializeStyleRefs(slideRoot, ctx)
	resolvePlaceholderGeometry(slideRoot, ctx)
	resolvePlaceholderRunColors(slideRoot, ctx)
	resolvePlaceholderRunSizes(slideRoot, ctx)
	resolveSchemeColors(slideRoot, ctx)
}

/**
 * Flatten a single lifted shape's theme dependencies in place — the shape-scoped
 * subset of {@link flattenSlide} used by `Presentation.importShape` `preserve`
 * mode. It runs every pass that resolves a *shape's* theme references against the
 * source theme (style-matrix refs, placeholder-inherited geometry/colour/size,
 * scheme colours) but deliberately **omits** the two slide-scoped background
 * passes (`applyInheritedBackground`/`materializeBackground`): a background
 * belongs to a slide, not to a shape being composed onto a foreign host.
 *
 * `shapeRoot` must be an element whose *descendants* include the lifted shape
 * (the passes match via `getElementsByTagNameNS`, which excludes the root element
 * itself) — the caller wraps the imported `p:sp`/`p:pic`/`p:graphicFrame`/`p:grpSp`
 * in a throwaway container before calling. The caller marks the part dirty.
 */
export function flattenShape(shapeRoot: Element, ctx: FlattenContext): void {
	materializeStyleRefs(shapeRoot, ctx)
	resolvePlaceholderGeometry(shapeRoot, ctx)
	resolvePlaceholderRunColors(shapeRoot, ctx)
	resolvePlaceholderRunSizes(shapeRoot, ctx)
	resolveSchemeColors(shapeRoot, ctx)
}

/**
 * Restyle a rebound slide in place (the `theme: 'restyle'` mode): the exact
 * inverse of {@link flattenSlide}. It bakes *nothing* — leaving every
 * `a:schemeClr`/style-matrix ref and `p:bg` `bgRef` symbolic is the whole point,
 * so they re-resolve against the *destination* master's `clrMap` + theme once the
 * slide is rebound. The single mutation is dropping the slide's own colour-map
 * override (`p:clrMapOvr/a:overrideClrMapping`): a source override would keep
 * mapping the slide's scheme-token names the source way and defeat the re-brand,
 * so it must yield to the destination master's `p:clrMap`.
 *
 * Caveat carried from the plan: only *symbolic* colours re-brand. Anything the
 * source authored as a literal `a:srgbClr` has no theme reference to re-resolve
 * and stays exactly that colour. The caller marks the part dirty.
 */
export function restyleSlide(slideRoot: Element): void {
	removeChildrenByQName(slideRoot, ['p:clrMapOvr'])
}

/**
 * If the slide has no own `p:bg`, insert (a copy of) the background it inherited
 * from its source layout/master as an explicit `p:cSld/p:bg`. The clone is left
 * unresolved here; the later passes flatten its `bgRef`/`schemeClr` in place.
 */
function applyInheritedBackground(slideRoot: Element, ctx: FlattenContext): void {
	const inherited = ctx.inheritedBackground
	if (!inherited) return
	const cSld = firstChild(slideRoot, 'p:cSld')
	if (!cSld || firstChild(cSld, 'p:bg')) return // no cSld, or the slide already owns a background
	const doc = slideRoot.ownerDocument!
	const bg = doc.importNode(inherited, true) as Element
	insertInOrder(cSld, bg, ['p:spTree', 'p:custDataLst', 'p:controls', 'p:extLst'])
}

/**
 * Resolve every `p:bgRef` (theme-indexed background) under the slide into an
 * explicit `p:bgPr` fill, so the background no longer depends on the destination
 * theme's `fmtScheme`. A `bgPr` background is left for the scheme-colour sweep.
 */
function materializeBackground(slideRoot: Element, ctx: FlattenContext): void {
	const doc = slideRoot.ownerDocument!
	for (const bg of elementsByTag(slideRoot, OOXML_NS.p, 'bg')) {
		const bgRef = firstChild(bg, 'p:bgRef')
		if (!bgRef) continue
		const idx = intAttr(bgRef, 'idx')
		const ref = resolveColor(firstChildElement(bgRef), ctx)
		const fill = idx !== null && idx > 0 ? (idx >= 1000 ? fmtEntry(ctx, 'a:bgFillStyleLst', idx - 1000) : fmtEntry(ctx, 'a:fillStyleLst', idx)) : null
		const bgPr = createElement(doc, 'p:bgPr')
		if (fill && ref) {
			substitutePhClr(fill, ref)
			bgPr.appendChild(fill)
		} else {
			bgPr.appendChild(createElement(doc, 'a:noFill')) // idx 0 / unresolved → transparent
		}
		bg.replaceChild(bgPr, bgRef)
	}
}

/** Rewrite every `a:schemeClr` under `root` to a literal `a:srgbClr` when resolvable. */
function resolveSchemeColors(root: Element, ctx: FlattenContext): void {
	const doc = root.ownerDocument!
	for (const schemeClr of elementsByTag(root, OOXML_NS.a, 'schemeClr')) {
		const token = attr(schemeClr, 'val')
		const hex = token ? resolveSchemeToken(token, ctx) : null
		if (!hex) continue // phClr or an unmapped token — leave it for the destination theme.
		const srgb = createElement(doc, 'a:srgbClr')
		setAttr(srgb, 'val', hex)
		while (schemeClr.firstChild) srgb.appendChild(schemeClr.firstChild) // carry transforms
		schemeClr.parentNode!.replaceChild(srgb, schemeClr)
	}
}

/** Schema successors of `a:solidFill` inside `a:rPr` (CT_TextCharacterProperties sequence). */
const RPR_FILL_AFTER = [
	'a:effectLst',
	'a:effectDag',
	'a:highlight',
	'a:uLnTx',
	'a:uLn',
	'a:uFillTx',
	'a:uFill',
	'a:latin',
	'a:ea',
	'a:cs',
	'a:sym',
	'a:hlinkClick',
	'a:hlinkMouseOver',
	'a:rtl',
	'a:extLst',
]

/** The master `p:txStyles` style element name for a placeholder category. */
const TX_STYLE_NAME: Record<'title' | 'body' | 'other', string> = {
	title: 'p:titleStyle',
	body: 'p:bodyStyle',
	other: 'p:otherStyle',
}

/**
 * Bake placeholder-inherited run colours onto the slide (gap 1). For each
 * placeholder run that defines no colour of its own (nor at paragraph/text-body
 * level on the *slide*), resolve the colour it would inherit from the source
 * `slideLayout`/`slideMaster` text styles and write it explicitly onto the run's
 * `a:rPr`. After this the run's colour cannot change when the slide is rebound to
 * the destination master. `a:fld` runs (dates, slide numbers) are treated like
 * `a:r`. Only colour is resolved; other inheritable run properties are left to
 * re-bind to the destination styles.
 */
function resolvePlaceholderRunColors(slideRoot: Element, ctx: FlattenContext): void {
	if (!ctx.layoutRoot && !ctx.masterRoot) return
	for (const sp of elementsByTag(slideRoot, OOXML_NS.p, 'sp')) {
		const ph = placeholderOf(sp)
		if (!ph) continue
		const txBody = firstChild(sp, 'p:txBody')
		if (!txBody) continue
		const type = attr(ph, 'type')
		const idx = attr(ph, 'idx') ?? '0'
		const slideLst = firstChild(txBody, 'a:lstStyle')
		const byLevel = new Map<number, ResolvedColor | null>()
		for (const p of getElements(txBody, 'a:p')) {
			const pPr = firstChild(p, 'a:pPr')
			const level = (pPr && intValue(attr(pPr, 'lvl'))) ?? 0
			const runs = [...getElements(p, 'a:r'), ...getElements(p, 'a:fld')]
			if (runs.length === 0) continue
			let color = byLevel.get(level)
			if (color === undefined) {
				color = placeholderInheritedColor(type, idx, level, ctx)
				byLevel.set(level, color)
			}
			if (!color) continue
			for (const run of runs) {
				if (slideDefinesColor(run, pPr, slideLst, level)) continue
				writeRunColor(run, color)
			}
		}
	}
}

/** The `p:ph` element of a shape (`p:sp/p:nvSpPr/p:nvPr/p:ph`), or `null`. */
function placeholderOf(sp: Element): Element | null {
	const nvSpPr = firstChild(sp, 'p:nvSpPr')
	const nvPr = nvSpPr && firstChild(nvSpPr, 'p:nvPr')
	return nvPr ? firstChild(nvPr, 'p:ph') : null
}

/** The master text-style category a placeholder type resolves against (absent ⇒ `obj` ⇒ body). */
function phCategory(type: string | null): 'title' | 'body' | 'other' {
	if (type === 'title' || type === 'ctrTitle') return 'title'
	if (type === null || type === 'body' || type === 'subTitle' || type === 'obj') return 'body'
	return 'other'
}

/**
 * The placeholder shape in `root` (a layout/master) that the given slide
 * placeholder inherits from: prefer a same-`idx` placeholder of the same
 * category, then any same-`idx`, then any same-category. Returns `null` when none
 * match.
 */
function findPlaceholder(root: Element, slideType: string | null, slideIdx: string): Element | null {
	const cat = phCategory(slideType)
	let idxMatch: Element | null = null
	let catMatch: Element | null = null
	for (const sp of elementsByTag(root, OOXML_NS.p, 'sp')) {
		const ph = placeholderOf(sp)
		if (!ph) continue
		const i = attr(ph, 'idx') ?? '0'
		const sameCat = phCategory(attr(ph, 'type')) === cat
		if (i === slideIdx && sameCat) return sp
		if (i === slideIdx && !idxMatch) idxMatch = sp
		if (sameCat && !catMatch) catMatch = sp
	}
	return idxMatch ?? catMatch
}

/** The `a:lstStyle` of a placeholder shape's `p:txBody`, or `null`. */
function placeholderLstStyle(sp: Element): Element | null {
	const txBody = firstChild(sp, 'p:txBody')
	return txBody ? firstChild(txBody, 'a:lstStyle') : null
}

/**
 * The `a:defRPr` for `level` (0-based) in a `CT_TextListStyle` (`a:lstStyle` or a
 * `p:txStyles` style): the level-specific `a:lvlNpPr/a:defRPr`, else the
 * `a:defPPr/a:defRPr` fallback. The shared root for colour and size resolution.
 */
function lstStyleLevelDefRPr(listStyle: Element | null, level: number): Element | null {
	if (!listStyle) return null
	const lvl = firstChild(listStyle, `a:lvl${level + 1}pPr`) ?? firstChild(listStyle, 'a:defPPr')
	return lvl ? firstChild(lvl, 'a:defRPr') : null
}

/** The `a:solidFill` of a level's `a:defRPr`, or `null` when none is defined there. */
export function lstStyleLevelFill(listStyle: Element | null, level: number): Element | null {
	const defRPr = lstStyleLevelDefRPr(listStyle, level)
	return defRPr ? firstChild(defRPr, 'a:solidFill') : null
}

/**
 * The colour *element* a placeholder run inherits from the source style chain: the
 * layout placeholder's `a:lstStyle`, then the master placeholder's, then the
 * master `p:txStyles` category style. Returns the first tier's colour element that
 * resolves against `ctx` (the `a:srgbClr`/`a:schemeClr`/… inside its
 * `a:solidFill`), or `null` when nothing in the chain defines a resolvable colour.
 * The read-model `Run.resolvedColor` getter feeds this element to
 * `resolveColorElement` for the `effectiveHex`; the flatten path resolves it
 * directly via {@link placeholderInheritedColor}.
 */
export function placeholderInheritedFill(type: string | null, idx: string, level: number, ctx: FlattenContext): Element | null {
	const tiers: (Element | null)[] = []
	if (ctx.layoutRoot) {
		const layoutPh = findPlaceholder(ctx.layoutRoot, type, idx)
		tiers.push(layoutPh && lstStyleLevelFill(placeholderLstStyle(layoutPh), level))
	}
	if (ctx.masterRoot) {
		const masterPh = findPlaceholder(ctx.masterRoot, type, idx)
		tiers.push(masterPh && lstStyleLevelFill(placeholderLstStyle(masterPh), level))
		const txStyles = firstChild(ctx.masterRoot, 'p:txStyles')
		const styleEl = txStyles && firstChild(txStyles, TX_STYLE_NAME[phCategory(type)])
		tiers.push(styleEl && lstStyleLevelFill(styleEl, level))
	}
	for (const fill of tiers) {
		const colorEl = fill && firstChildElement(fill)
		if (colorEl && resolveColor(colorEl, ctx)) return colorEl
	}
	return null
}

/**
 * The colour a placeholder run inherits from the source style chain, resolved to a
 * literal `{ hex, transforms }`. Thin wrapper over {@link placeholderInheritedFill}
 * for the flatten path, which re-emits the transforms verbatim. Returns `null`
 * when nothing in the chain defines a resolvable colour (the run then re-binds to
 * the destination).
 */
function placeholderInheritedColor(type: string | null, idx: string, level: number, ctx: FlattenContext): ResolvedColor | null {
	const colorEl = placeholderInheritedFill(type, idx, level, ctx)
	return colorEl ? resolveColor(colorEl, ctx) : null
}

/** Whether the *slide itself* already fixes this run's colour (so a rebind cannot change it). */
function slideDefinesColor(run: Element, pPr: Element | null, slideLst: Element | null, level: number): boolean {
	const rPr = firstChild(run, 'a:rPr')
	if (rPr && FILL_CHOICES.some((q) => firstChild(rPr, q))) return true
	const defRPr = pPr && firstChild(pPr, 'a:defRPr')
	if (defRPr && firstChild(defRPr, 'a:solidFill')) return true
	return !!lstStyleLevelFill(slideLst, level)
}

/** Write a resolved colour as an explicit `a:solidFill` (with carried transforms) onto a run's `a:rPr`. */
function writeRunColor(run: Element, color: ResolvedColor): void {
	const doc = run.ownerDocument!
	const rPr = getOrAddChild(run, 'a:rPr', ['a:t'])
	const fill = createElement(doc, 'a:solidFill')
	const srgb = createElement(doc, 'a:srgbClr')
	setAttr(srgb, 'val', color.hex)
	for (const t of color.transforms) srgb.appendChild(t.cloneNode(true) as Element)
	fill.appendChild(srgb)
	insertInOrder(rPr, fill, RPR_FILL_AFTER)
}

/**
 * Bake placeholder-inherited geometry onto the slide (gap 3). A placeholder shape
 * that carries no own `p:spPr/a:xfrm` takes its position/size from the matching
 * source `slideLayout` placeholder, else the `slideMaster` placeholder. Rebinding
 * to the destination master replaces that inheritance, so the placeholder would
 * snap to the destination default (often clipping off-canvas). We deep-clone the
 * effective source `a:xfrm` and write it explicitly onto the shape. Shapes with
 * their own `a:xfrm` are left untouched (explicit geometry is not inherited), and
 * an orphan placeholder with no source match keeps the current fall-back behaviour.
 */
function resolvePlaceholderGeometry(slideRoot: Element, ctx: FlattenContext): void {
	if (!ctx.layoutRoot && !ctx.masterRoot) return
	for (const sp of elementsByTag(slideRoot, OOXML_NS.p, 'sp')) {
		const ph = placeholderOf(sp)
		if (!ph) continue
		const spPr = firstChild(sp, 'p:spPr')
		if (spPr && firstChild(spPr, 'a:xfrm')) continue // explicit geometry is not inherited — leave it
		const xfrm = placeholderInheritedXfrm(attr(ph, 'type'), attr(ph, 'idx') ?? '0', ctx)
		if (!xfrm) continue
		const target = getOrAddChild(sp, 'p:spPr', SHAPE_AFTER_SPPR)
		insertInOrder(target, xfrm.cloneNode(true) as Element, SPPR_XFRM_AFTER)
	}
}

/** The `a:xfrm` a placeholder inherits from the source layout, then master, or `null`. */
function placeholderInheritedXfrm(type: string | null, idx: string, ctx: FlattenContext): Element | null {
	for (const root of [ctx.layoutRoot, ctx.masterRoot]) {
		if (!root) continue
		const ph = findPlaceholder(root, type, idx)
		const spPr = ph && firstChild(ph, 'p:spPr')
		const xfrm = spPr && firstChild(spPr, 'a:xfrm')
		if (xfrm) return xfrm
	}
	return null
}

/** Inheritable run properties baked under `preserve`: size and weight/slant (not typeface). */
const RUN_PROP_NAMES = ['sz', 'b', 'i'] as const
type RunProps = Record<(typeof RUN_PROP_NAMES)[number], string | null>

/**
 * Bake placeholder-inherited run size/weight onto the slide (gap 3). Mirrors
 * {@link resolvePlaceholderRunColors}: for each placeholder run that sets no
 * `sz`/`b`/`i` of its own (nor at paragraph/text-body level on the *slide*),
 * resolve the value it would inherit from the source `slideLayout`/`slideMaster`
 * text styles — per paragraph list level — and write it explicitly onto the run's
 * `a:rPr`. Each property resolves independently up the chain. Typeface (`a:latin`)
 * is left to re-bind to the destination theme, as gap 1 does for the `fontRef`.
 */
function resolvePlaceholderRunSizes(slideRoot: Element, ctx: FlattenContext): void {
	if (!ctx.layoutRoot && !ctx.masterRoot) return
	for (const sp of elementsByTag(slideRoot, OOXML_NS.p, 'sp')) {
		const ph = placeholderOf(sp)
		if (!ph) continue
		const txBody = firstChild(sp, 'p:txBody')
		if (!txBody) continue
		const type = attr(ph, 'type')
		const idx = attr(ph, 'idx') ?? '0'
		const slideLst = firstChild(txBody, 'a:lstStyle')
		const byLevel = new Map<number, RunProps | null>()
		for (const p of getElements(txBody, 'a:p')) {
			const pPr = firstChild(p, 'a:pPr')
			const level = (pPr && intValue(attr(pPr, 'lvl'))) ?? 0
			const runs = [...getElements(p, 'a:r'), ...getElements(p, 'a:fld')]
			if (runs.length === 0) continue
			let props = byLevel.get(level)
			if (props === undefined) {
				props = placeholderInheritedRunProps(type, idx, level, ctx)
				byLevel.set(level, props)
			}
			if (!props) continue
			for (const run of runs) writeRunProps(run, props, pPr, slideLst, level)
		}
	}
}

/**
 * The run size/weight a placeholder run inherits from the source style chain:
 * layout placeholder `a:lstStyle` → master placeholder `a:lstStyle` → master
 * `p:txStyles` category style, per list level. Each of `sz`/`b`/`i` is taken from
 * the first tier that defines it (properties resolve independently). Returns
 * `null` when no tier defines any of them.
 */
function placeholderInheritedRunProps(type: string | null, idx: string, level: number, ctx: FlattenContext): RunProps | null {
	const tiers: (Element | null)[] = []
	if (ctx.layoutRoot) {
		const layoutPh = findPlaceholder(ctx.layoutRoot, type, idx)
		tiers.push(layoutPh && lstStyleLevelDefRPr(placeholderLstStyle(layoutPh), level))
	}
	if (ctx.masterRoot) {
		const masterPh = findPlaceholder(ctx.masterRoot, type, idx)
		tiers.push(masterPh && lstStyleLevelDefRPr(placeholderLstStyle(masterPh), level))
		const txStyles = firstChild(ctx.masterRoot, 'p:txStyles')
		const styleEl = txStyles && firstChild(txStyles, TX_STYLE_NAME[phCategory(type)])
		tiers.push(styleEl && lstStyleLevelDefRPr(styleEl, level))
	}
	const props = {} as RunProps
	let any = false
	for (const name of RUN_PROP_NAMES) {
		let value: string | null = null
		for (const tier of tiers) {
			value = tier && attr(tier, name)
			if (value != null) break
		}
		props[name] = value
		if (value != null) any = true
	}
	return any ? props : null
}

/** Whether the *slide itself* already fixes a run property (so a rebind cannot change it). */
function slideDefinesProp(name: string, run: Element, pPr: Element | null, slideLst: Element | null, level: number): boolean {
	const rPr = firstChild(run, 'a:rPr')
	if (rPr && attr(rPr, name) != null) return true
	const defRPr = pPr && firstChild(pPr, 'a:defRPr')
	if (defRPr && attr(defRPr, name) != null) return true
	const slideDefRPr = lstStyleLevelDefRPr(slideLst, level)
	return !!(slideDefRPr && attr(slideDefRPr, name) != null)
}

/** Write each resolved run property onto a run's `a:rPr`, skipping ones the slide already fixes. */
function writeRunProps(run: Element, props: RunProps, pPr: Element | null, slideLst: Element | null, level: number): void {
	let rPr: Element | null = null
	for (const name of RUN_PROP_NAMES) {
		const value = props[name]
		if (value == null) continue
		if (slideDefinesProp(name, run, pPr, slideLst, level)) continue
		rPr ??= getOrAddChild(run, 'a:rPr', ['a:t'])
		setAttr(rPr, name, value)
	}
}

/** Snapshot all descendant elements of `root` matching a namespace + local name. */
function elementsByTag(root: Element, ns: string, local: string): Element[] {
	const live = root.getElementsByTagNameNS(ns, local)
	const out: Element[] = []
	for (let i = 0; i < live.length; i++) out.push(live[i])
	return out
}

/**
 * Resolve each shape's `p:style` `lnRef`/`fillRef`/`effectRef` into explicit
 * `spPr` children (using the theme `fmtScheme`), then neutralize the ref. The
 * `fontRef` is intentionally left for the destination theme to re-resolve.
 */
function materializeStyleRefs(root: Element, ctx: FlattenContext): void {
	if (!ctx.fmtScheme) return
	for (const style of elementsByTag(root, OOXML_NS.p, 'style')) {
		const shape = style.parentNode as Element | null
		if (!shape) continue
		const spPr = getOrAddChild(shape, 'p:spPr', SHAPE_AFTER_SPPR)
		materializeFill(spPr, firstChild(style, 'a:fillRef'), ctx)
		materializeLine(spPr, firstChild(style, 'a:lnRef'), ctx)
		materializeEffect(spPr, firstChild(style, 'a:effectRef'), ctx)
	}
}

/** Replace every `phClr` under `el` with the ref colour (ref transforms first, then the `phClr`'s own). */
function substitutePhClr(el: Element, ref: ResolvedColor): void {
	const doc = el.ownerDocument!
	for (const phClr of elementsByTag(el, OOXML_NS.a, 'schemeClr')) {
		if (attr(phClr, 'val') !== 'phClr') continue
		const srgb = createElement(doc, 'a:srgbClr')
		setAttr(srgb, 'val', ref.hex)
		for (const t of ref.transforms) srgb.appendChild(t.cloneNode(true))
		while (phClr.firstChild) srgb.appendChild(phClr.firstChild)
		phClr.parentNode!.replaceChild(srgb, phClr)
	}
}

/** The `idx`-th entry (1-based) of a `fmtScheme` style list, deep-cloned, or `null`. */
function fmtEntry(ctx: FlattenContext, listName: string, idx: number): Element | null {
	const list = ctx.fmtScheme && firstChild(ctx.fmtScheme, listName)
	if (!list || idx < 1) return null
	const entry = childElements(list)[idx - 1]
	return entry ? (entry.cloneNode(true) as Element) : null
}

/**
 * Build the explicit fill element a `p:style` `a:fillRef` resolves to — the indexed
 * `fmtScheme` `fillStyleLst`/`bgFillStyleLst` entry (deep-cloned) with its `phClr`
 * replaced by the ref's resolved colour. Pure: mutates neither the ref nor the
 * theme. `null` when the ref is absent, `idx` is 0/unset, or the entry or its
 * colour cannot be resolved. Shared by the flatten path ({@link materializeFill})
 * and the read-model `resolveStyleFillColor` getter so both see the same fill.
 */
export function styleRefFill(fillRef: Element | null, ctx: FlattenContext): Element | null {
	if (!fillRef) return null
	const idx = intAttr(fillRef, 'idx')
	if (idx === null || idx <= 0) return null
	// idx >= 1000 selects bgFillStyleLst (offset by 1000); otherwise fillStyleLst.
	const fill = idx >= 1000 ? fmtEntry(ctx, 'a:bgFillStyleLst', idx - 1000) : fmtEntry(ctx, 'a:fillStyleLst', idx)
	const ref = resolveColor(firstChildElement(fillRef), ctx)
	if (!fill || !ref) return null
	substitutePhClr(fill, ref)
	return fill
}

/**
 * Build the explicit `a:ln` element a `p:style` `a:lnRef` resolves to — the indexed
 * `fmtScheme` `lnStyleLst` entry (deep-cloned) with its `phClr` replaced by the
 * ref's resolved colour. Pure; the line/read counterpart of {@link styleRefFill}.
 */
export function styleRefLine(lnRef: Element | null, ctx: FlattenContext): Element | null {
	if (!lnRef) return null
	const idx = intAttr(lnRef, 'idx')
	if (idx === null || idx <= 0) return null
	const ln = fmtEntry(ctx, 'a:lnStyleLst', idx)
	const ref = resolveColor(firstChildElement(lnRef), ctx)
	if (!ln || !ref) return null
	substitutePhClr(ln, ref)
	return ln
}

function materializeFill(spPr: Element, fillRef: Element | null, ctx: FlattenContext): void {
	if (!fillRef) return
	if (!FILL_CHOICES.some((q) => firstChild(spPr, q))) {
		const fill = styleRefFill(fillRef, ctx)
		if (fill) insertInOrder(spPr, fill, SPPR_FILL_AFTER)
	}
	neutralizeRef(fillRef)
}

function materializeLine(spPr: Element, lnRef: Element | null, ctx: FlattenContext): void {
	if (!lnRef) return
	if (!firstChild(spPr, 'a:ln')) {
		const ln = styleRefLine(lnRef, ctx)
		if (ln) insertInOrder(spPr, ln, SPPR_LN_AFTER)
	}
	neutralizeRef(lnRef)
}

function materializeEffect(spPr: Element, effectRef: Element | null, ctx: FlattenContext): void {
	if (!effectRef) return
	const idx = intAttr(effectRef, 'idx')
	if (idx !== null && idx > 0 && !firstChild(spPr, 'a:effectLst') && !firstChild(spPr, 'a:effectDag')) {
		const style = fmtEntry(ctx, 'a:effectStyleLst', idx) // a:effectStyle (effectLst?, scene3d?, sp3d?)
		const ref = resolveColor(firstChildElement(effectRef), ctx)
		if (style && ref) {
			substitutePhClr(style, ref)
			// Lift the effectStyle's children (effectLst/scene3d/sp3d) into spPr, in order.
			for (const child of childElements(style)) {
				if (isA(child, 'effectLst') || isA(child, 'effectDag')) insertInOrder(spPr, child, SPPR_EFFECT_AFTER)
				else if (isA(child, 'scene3d')) insertInOrder(spPr, child, ['a:sp3d', 'a:extLst'])
				else if (isA(child, 'sp3d')) insertInOrder(spPr, child, ['a:extLst'])
			}
		}
	}
	neutralizeRef(effectRef)
}

/** Strip a style-matrix ref to `idx="0"` with no colour child so it contributes nothing. */
function neutralizeRef(ref: Element): void {
	setAttr(ref, 'idx', '0')
	for (const child of childElements(ref)) ref.removeChild(child)
}

function intAttr(el: Element, name: string): number | null {
	const value = attr(el, name)
	if (value === null || value === '') return null
	const n = Number(value)
	return Number.isFinite(n) ? n : null
}
