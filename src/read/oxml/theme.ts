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
 *
 * All functions operate purely on DOM elements; the caller gathers the source
 * `clrMap` / `clrScheme` / `fmtScheme` parts and owns marking the slide dirty.
 */
import { ELEMENT_NODE, OOXML_NS, attr, createElement, firstChild, getOrAddChild, insertInOrder, setAttr, type Element } from './dom.js'
import { FILL_CHOICES } from './fill.js'

/** The 12 `a:clrScheme` slot names, in schema order. */
const SCHEME_SLOTS = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
/** Scheme tokens that name a `clrScheme` slot directly, bypassing the `clrMap`. */
const DIRECT_SLOT_TOKENS = new Set(['dk1', 'lt1', 'dk2', 'lt2'])

/** Schema successors for ordered insertion into `p:spPr` (CT_ShapeProperties). */
const SPPR_FILL_AFTER = ['a:ln', 'a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
const SPPR_LN_AFTER = ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst']
const SPPR_EFFECT_AFTER = ['a:scene3d', 'a:sp3d', 'a:extLst']
const SHAPE_AFTER_SPPR = ['p:style', 'p:txBody']

/**
 * Everything {@link flattenSlide} needs from the slide's source theme subgraph:
 * the effective colour map (token → slot), the resolved colour scheme (slot →
 * 6-hex RGB), and the live `a:fmtScheme` element (for style-matrix resolution).
 */
export interface FlattenContext {
	clrMap: Map<string, string>
	clrScheme: Map<string, string>
	fmtScheme: Element | null
	/**
	 * The slide's effective background inherited from the *source*
	 * `slideLayout`/`slideMaster` (the raw `p:bg` element), or `null` when the
	 * slide carries its own. Applied onto the slide before flattening so the
	 * background survives rebinding to the destination master. See
	 * {@link flattenSlide}.
	 */
	inheritedBackground?: Element | null
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
function resolveSchemeToken(token: string, ctx: FlattenContext): string | null {
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
interface ResolvedColor {
	hex: string
	transforms: Element[]
}

/**
 * Resolve a DrawingML colour element to a literal `{ hex, transforms }`, routing
 * `a:schemeClr` through the context. Returns `null` when the base cannot be made
 * literal (unmapped token, or a colour model we do not flatten).
 */
function resolveColor(color: Element | null, ctx: FlattenContext): ResolvedColor | null {
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
 * 3. rewrite every remaining `a:schemeClr` to its literal `a:srgbClr`.
 *
 * Steps run in this order so the inherited/materialized backgrounds are present
 * before the final scheme-colour sweep resolves the colours they carry. The
 * caller marks the part dirty.
 */
export function flattenSlide(slideRoot: Element, ctx: FlattenContext): void {
	applyInheritedBackground(slideRoot, ctx)
	materializeBackground(slideRoot, ctx)
	materializeStyleRefs(slideRoot, ctx)
	resolveSchemeColors(slideRoot, ctx)
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

function materializeFill(spPr: Element, fillRef: Element | null, ctx: FlattenContext): void {
	if (!fillRef) return
	const idx = intAttr(fillRef, 'idx')
	if (idx !== null && idx > 0 && !FILL_CHOICES.some((q) => firstChild(spPr, q))) {
		// idx >= 1000 selects bgFillStyleLst (offset by 1000); otherwise fillStyleLst.
		const fill = idx >= 1000 ? fmtEntry(ctx, 'a:bgFillStyleLst', idx - 1000) : fmtEntry(ctx, 'a:fillStyleLst', idx)
		const ref = resolveColor(firstChildElement(fillRef), ctx)
		if (fill && ref) {
			substitutePhClr(fill, ref)
			insertInOrder(spPr, fill, SPPR_FILL_AFTER)
		}
	}
	neutralizeRef(fillRef)
}

function materializeLine(spPr: Element, lnRef: Element | null, ctx: FlattenContext): void {
	if (!lnRef) return
	const idx = intAttr(lnRef, 'idx')
	if (idx !== null && idx > 0 && !firstChild(spPr, 'a:ln')) {
		const ln = fmtEntry(ctx, 'a:lnStyleLst', idx)
		const ref = resolveColor(firstChildElement(lnRef), ctx)
		if (ln && ref) {
			substitutePhClr(ln, ref)
			insertInOrder(spPr, ln, SPPR_LN_AFTER)
		}
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
