/**
 * Baking a slide's theme dependencies into the slide itself, for `Presentation.importSlide` /
 * `importShape`.
 *
 * A slide imported into another deck is rebound to the *destination* master, so everything it
 * resolved symbolically against its own theme silently re-resolves against a different one — its
 * scheme colours change, its style-matrix fills change, its placeholders snap to different
 * geometry and type sizes. The `theme: 'preserve'` mode exists to stop that: each pass here
 * resolves a dependency against the **source** theme (via `read/oxml/theme.ts` and
 * `read/oxml/placeholder-inherit.ts`) and writes the answer explicitly onto the slide, so after
 * the rebind there is nothing left to re-resolve.
 *
 * What gets flattened, and why each one matters:
 *
 * 1. **`a:schemeClr` tokens** — rewritten to the literal `a:srgbClr` the source `clrMap` +
 *    `clrScheme` produce. Child transforms (`lumMod`/`shade`/`alpha`/…) are carried through
 *    untouched, so tints and shades render identically; only the *base* reference is swapped,
 *    never the transform math.
 * 2. **Style-matrix refs** — a shape's `p:style` (`lnRef`/`fillRef`/`effectRef`) indexes the
 *    theme `fmtScheme`; the indexed entry is resolved into an explicit `spPr` fill/line/effect
 *    (with its `phClr` substituted) and the ref neutralized so it cannot re-resolve. The
 *    `fontRef` is deliberately left intact so its font can re-bind to the destination theme —
 *    the "normalize fonts on attach" bonus.
 * 3. **Placeholder-inherited run colour, size and weight** — a run that sets none of its own
 *    takes them from the source placeholder → layout → master text-style chain. Rebinding
 *    replaces that chain, so each is resolved from the source styles and written onto the run.
 *    Typeface (`a:latin`) is deliberately *not* baked — it re-binds along with `fontRef`.
 * 4. **Placeholder-inherited geometry** — a placeholder with no own `a:xfrm` takes position and
 *    size from the matching source layout/master placeholder; without baking it, a title snaps
 *    to the destination default and often clips off-canvas.
 *
 * The `restyle` mode is the deliberate inverse ({@link restyleSlide}): it bakes *nothing*, so
 * everything re-resolves against the destination and the slide takes on the new brand.
 *
 * **Why this lives in `api/ops/` and not in `read/oxml/`.** Everything here *mutates* a live
 * part. The modules it calls into resolve and build detached elements but never write, which is
 * what lets the read model's getters share them; keeping the writes on this side of the line is
 * what keeps that true. Callers own marking the part dirty.
 */
import {
	OOXML_NS,
	attr,
	childElements,
	createElement,
	descendantsByTag,
	firstChild,
	firstChildElement,
	getElements,
	getOrAddChild,
	insertInOrder,
	numberAttr,
	numberValue,
	ownerDocumentOf,
	removeChildrenByQName,
	setAttr,
	type Element,
} from '../../oxml/dom.js'
import { FILL_CHOICES } from '../../oxml/fill.js'
import {
	LST_STYLE_LEVELS,
	RUN_PROP_NAMES,
	placeholderInheritedAnchor,
	placeholderInheritedColor,
	placeholderInheritedListStyles,
	placeholderInheritedRunProps,
	placeholderInheritedXfrm,
	placeholderOf,
	slideDefinesColor,
	slideDefinesProp,
	type RunProps,
} from '../../oxml/placeholder-inherit.js'
import {
	type ColorContext,
	fmtEntry,
	isA,
	replaceColorElement,
	resolveColor,
	type ResolvedColorRef,
	resolveSchemeToken,
	SCHEME_SLOTS,
	styleRefFill,
	styleRefLine,
	substitutePhClr,
	type ThemeContext,
} from '../../oxml/theme.js'
import {
	RPR_FILL_AFTER,
	SHAPE_AFTER_SPPR,
	SPPR_AFTER_XFRM,
	SPPR_EFFECT_AFTER,
	SPPR_FILL_AFTER,
	SPPR_LN_AFTER,
	SPPR_SCENE3D_AFTER,
	SPPR_SP3D_AFTER,
} from '../../../ooxml/sequence.js'
import { cSldOf, nvPrOf } from '../../oxml/slide-dom.js'

/**
 * A {@link ThemeContext} plus the one thing only the flatten pass needs to write from.
 *
 * It is separate from `ThemeContext` because every *other* consumer of a theme context — the
 * read model's colour, font and geometry getters — resolves against a source subgraph without
 * ever carrying a background to apply. Keeping the field here means those callers cannot be
 * handed one and quietly ignore it.
 */
export interface FlattenContext extends ThemeContext {
	/**
	 * The slide's effective background inherited from the *source* `slideLayout`/`slideMaster`
	 * (the raw `p:bg` element), or `null` when the slide carries its own. Applied onto the slide
	 * before flattening so the background survives rebinding to the destination master.
	 */
	inheritedBackground?: Element | null
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
 * source theme (style-matrix refs, placeholder-inherited geometry/colour/size/
 * anchor/list-style, scheme colours) but deliberately **omits** the two
 * slide-scoped background passes (`applyInheritedBackground`/`materializeBackground`):
 * a background belongs to a slide, not to a shape being composed onto a foreign host.
 *
 * Unlike {@link flattenSlide}, this also **demotes** a lifted placeholder to a
 * plain shape ({@link demotePlaceholders}) once everything it inherited is baked.
 * A placeholder makes sense on its own slide, where it resolves against that deck's
 * master/layout; lifted onto a foreign host its surviving `p:ph` would re-resolve
 * against the *host* placeholder of the same type/idx (wrong inheritance, or a
 * fallback when absent) and could collide with the host's own placeholder. The
 * extra bakes here (anchor + list style, on top of the geometry/colour/size the
 * shared passes cover) make the shape self-contained so demotion loses nothing
 * visible. This is why the demotion is scoped to `flattenShape`: `flattenSlide`
 * keeps placeholders as placeholders by design.
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
	resolvePlaceholderBodyPr(shapeRoot, ctx)
	resolvePlaceholderListStyle(shapeRoot, ctx) // before resolveSchemeColors: cloned levels carry schemeClr
	resolveSchemeColors(shapeRoot, ctx)
	demotePlaceholders(shapeRoot) // last: the passes above key on p:ph
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
 * The optional inverse of {@link resolveSchemeColors}, for `restyle`'s
 * force-remap-literals mode (the `Presentation.importSlide` `remapLiterals`
 * option). Every literal `a:srgbClr` whose value equals a *source* `clrScheme`
 * slot is rewritten back to a symbolic `a:schemeClr` so it re-resolves against the
 * destination theme; a literal matching no slot (the common case) is left exactly
 * as authored. Transform children (`lumMod`/`shade`/`alpha`/…) are carried across.
 *
 * Plain `restyle` cannot re-brand a literal — it carries no theme reference. This
 * re-introduces one by matching the literal's RGB against the slots the *source*
 * theme defined, emitting the scheme token that routes through the source `clrMap`
 * (so a literal equal to the source `accent1` slot becomes `schemeClr accent1`,
 * which the destination `clrMap`/theme then re-resolve). Slots are matched in
 * `clrScheme` order, so the first slot defining a given RGB wins when several share
 * it. Opt-in because it deliberately reinterprets authored literals as theme
 * colours — visual QA territory, the same caveat as the rest of `restyle`.
 */
export function remapLiteralColors(slideRoot: Element, ctx: ColorContext): void {
	const slotByHex = reverseClrScheme(ctx.clrScheme)
	if (slotByHex.size === 0) return
	const tokenBySlot = reverseClrMap(ctx.clrMap)
	for (const srgb of descendantsByTag(slideRoot, OOXML_NS.a, 'srgbClr')) {
		const hex = attr(srgb, 'val')
		const slot = hex ? slotByHex.get(hex.toUpperCase()) : undefined
		if (!slot) continue
		// Routed through the source clrMap; dk1/lt1/… are themselves valid tokens.
		replaceColorElement(srgb, 'a:schemeClr', tokenBySlot.get(slot) ?? slot)
	}
}

/** Reverse a slot → RGB `clrScheme` into RGB → slot, in `SCHEME_SLOTS` order (first slot wins on a shared RGB). */
function reverseClrScheme(clrScheme: Map<string, string>): Map<string, string> {
	const out = new Map<string, string>()
	for (const slot of SCHEME_SLOTS) {
		const hex = clrScheme.get(slot)
		if (hex && !out.has(hex.toUpperCase())) out.set(hex.toUpperCase(), slot)
	}
	return out
}

/** Reverse a token → slot `clrMap` into slot → token (first token wins on a shared slot). */
function reverseClrMap(clrMap: Map<string, string>): Map<string, string> {
	const out = new Map<string, string>()
	for (const [token, slot] of clrMap) if (!out.has(slot)) out.set(slot, token)
	return out
}

/**
 * If the slide has no own `p:bg`, insert (a copy of) the background it inherited
 * from its source layout/master as an explicit `p:cSld/p:bg`. The clone is left
 * unresolved here; the later passes flatten its `bgRef`/`schemeClr` in place.
 */
function applyInheritedBackground(slideRoot: Element, ctx: FlattenContext): void {
	const inherited = ctx.inheritedBackground
	if (!inherited) return
	const cSld = cSldOf(slideRoot)
	if (!cSld || firstChild(cSld, 'p:bg')) return // no cSld, or the slide already owns a background
	const doc = ownerDocumentOf(slideRoot)
	const bg = doc.importNode(inherited, true)
	insertInOrder(cSld, bg, ['p:spTree', 'p:custDataLst', 'p:controls', 'p:extLst'])
}

/**
 * Resolve every `p:bgRef` (theme-indexed background) under the slide into an
 * explicit `p:bgPr` fill, so the background no longer depends on the destination
 * theme's `fmtScheme`. A `bgPr` background is left for the scheme-colour sweep.
 */
function materializeBackground(slideRoot: Element, ctx: FlattenContext): void {
	const doc = ownerDocumentOf(slideRoot)
	for (const bg of descendantsByTag(slideRoot, OOXML_NS.p, 'bg')) {
		const bgRef = firstChild(bg, 'p:bgRef')
		if (!bgRef) continue
		// A `p:bgRef` is a `CT_StyleMatrixReference`, the same shape as a shape's
		// `a:fillRef`, so `styleRefFill` builds its resolved `fmtScheme` fill (phClr
		// already substituted). `null` (idx 0 / unresolved colour) → transparent.
		const fill = styleRefFill(bgRef, ctx)
		const bgPr = createElement(doc, 'p:bgPr')
		bgPr.appendChild(fill ?? createElement(doc, 'a:noFill'))
		bg.replaceChild(bgPr, bgRef)
	}
}

/** Rewrite every `a:schemeClr` under `root` to a literal `a:srgbClr` when resolvable. */
function resolveSchemeColors(root: Element, ctx: FlattenContext): void {
	for (const schemeClr of descendantsByTag(root, OOXML_NS.a, 'schemeClr')) {
		const token = attr(schemeClr, 'val')
		const hex = token ? resolveSchemeToken(token, ctx) : null
		if (!hex) continue // phClr or an unmapped token — leave it for the destination theme.
		replaceColorElement(schemeClr, 'a:srgbClr', hex)
	}
}

/**
 * Bake placeholder-inherited run colours onto the slide. For each placeholder run
 * that defines no colour of its own (nor at paragraph/text-body level on the
 * *slide*), resolve the colour it would inherit from the source
 * `slideLayout`/`slideMaster` text styles and write it explicitly onto the run's
 * `a:rPr`. After this the run's colour cannot change when the slide is rebound to
 * the destination master. `a:fld` runs (dates, slide numbers) are treated like
 * `a:r`. Only colour is resolved; other inheritable run properties are left to
 * re-bind to the destination styles.
 */
function resolvePlaceholderRunColors(slideRoot: Element, ctx: FlattenContext): void {
	bakePlaceholderRunProperty(slideRoot, ctx, placeholderInheritedColor, (run, color, pPr, slideLst, level) => {
		if (slideDefinesColor(run, pPr, slideLst, level)) return
		writeRunColor(run, color)
	})
}

/**
 * The traversal both run-property passes share: every placeholder shape's text body, every
 * paragraph in it, every `a:r` and `a:fld` in that paragraph.
 *
 * `resolve` is asked once per *list level* per shape and memoized, because walking the
 * layout/master style chain is the expensive half and every paragraph at the same level gets
 * the same answer. `write` then runs per run, and owns the decision to skip a run the slide
 * already fixes — the two passes ask that question differently (one colour check versus three
 * independent property checks), so it stays on their side of the boundary rather than becoming
 * a predicate parameter only one caller could use.
 *
 * @param slideRoot - the slide part's root element
 * @param ctx - the flatten context, carrying the source layout/master roots
 * @param resolve - the inherited value for a placeholder type/idx at a list level, or `null`
 * @param write - apply a resolved value to one run
 */
/**
 * Visit every placeholder shape in a lifted subtree that has a text body, with the two
 * identifiers each pass resolves inheritance against.
 *
 * The preamble is the same for all three of them: nothing to inherit *from* when neither a
 * layout nor a master came across, a `p:sp` without a `p:ph` is not a placeholder, and one
 * without a `p:txBody` has no text to inherit anything for. It was written out three times,
 * which is one more than the file's own {@link bakePlaceholderRunProperty} was already
 * generalising over.
 *
 * @param shapeRoot - the root of the lifted subtree
 * @param ctx - the flatten context, carrying the source layout/master roots
 * @param visit - called per placeholder with its text body, `p:ph/@type` and `@idx`
 */
function forEachPlaceholderTextBody(
	shapeRoot: Element,
	ctx: FlattenContext,
	visit: (txBody: Element, type: string | null, idx: string) => void
): void {
	if (!ctx.layoutRoot && !ctx.masterRoot) return
	for (const sp of descendantsByTag(shapeRoot, OOXML_NS.p, 'sp')) {
		const ph = placeholderOf(sp)
		if (!ph) continue
		const txBody = firstChild(sp, 'p:txBody')
		if (!txBody) continue
		visit(txBody, attr(ph, 'type'), attr(ph, 'idx') ?? '0')
	}
}

function bakePlaceholderRunProperty<T>(
	slideRoot: Element,
	ctx: FlattenContext,
	resolve: (type: string | null, idx: string, level: number, ctx: FlattenContext) => T | null,
	write: (run: Element, value: T, pPr: Element | null, slideLst: Element | null, level: number) => void
): void {
	forEachPlaceholderTextBody(slideRoot, ctx, (txBody, type, idx) => {
		const slideLst = firstChild(txBody, 'a:lstStyle')
		const byLevel = new Map<number, T | null>()
		for (const p of getElements(txBody, 'a:p')) {
			const pPr = firstChild(p, 'a:pPr')
			const level = (pPr && numberValue(attr(pPr, 'lvl'))) ?? 0
			const runs = [...getElements(p, 'a:r'), ...getElements(p, 'a:fld')]
			if (runs.length === 0) continue
			let value = byLevel.get(level)
			if (value === undefined) {
				value = resolve(type, idx, level, ctx)
				byLevel.set(level, value)
			}
			if (!value) continue
			for (const run of runs) write(run, value, pPr, slideLst, level)
		}
	})
}

/** Write a resolved colour as an explicit `a:solidFill` (with carried transforms) onto a run's `a:rPr`. */
function writeRunColor(run: Element, color: ResolvedColorRef): void {
	const doc = ownerDocumentOf(run)
	const rPr = getOrAddChild(run, 'a:rPr', ['a:t'])
	const fill = createElement(doc, 'a:solidFill')
	const srgb = createElement(doc, 'a:srgbClr')
	setAttr(srgb, 'val', color.hex)
	for (const t of color.transforms) srgb.appendChild(t.cloneNode(true))
	fill.appendChild(srgb)
	insertInOrder(rPr, fill, RPR_FILL_AFTER)
}

/**
 * Bake placeholder-inherited geometry onto the slide. A placeholder shape
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
	for (const sp of descendantsByTag(slideRoot, OOXML_NS.p, 'sp')) {
		const ph = placeholderOf(sp)
		if (!ph) continue
		const spPr = firstChild(sp, 'p:spPr')
		if (spPr && firstChild(spPr, 'a:xfrm')) continue // explicit geometry is not inherited — leave it
		const found = placeholderInheritedXfrm(attr(ph, 'type'), attr(ph, 'idx') ?? '0', ctx)
		if (!found) continue
		const target = getOrAddChild(sp, 'p:spPr', SHAPE_AFTER_SPPR)
		insertInOrder(target, found.xfrm.cloneNode(true), SPPR_AFTER_XFRM)
	}
}

/**
 * Bake the placeholder-inherited vertical anchor onto a lifted shape's text body
 * (the {@link flattenShape} placeholder-demotion path). A placeholder whose own
 * `a:bodyPr` sets no `@anchor` inherits it from the source layout/master placeholder;
 * once {@link demotePlaceholders} strips the `p:ph` the shape would lose that
 * inheritance and fall back to top-anchored, so a vertically-centred title jumps.
 * We resolve the effective anchor from the source chain and write it explicitly.
 * Shapes whose `a:bodyPr` already fixes `@anchor`, and placeholders with no
 * resolvable inherited anchor, are left untouched. Other `a:bodyPr` knobs (insets,
 * autofit) are left as authored — they have sane defaults a plain shape keeps.
 */
function resolvePlaceholderBodyPr(shapeRoot: Element, ctx: FlattenContext): void {
	forEachPlaceholderTextBody(shapeRoot, ctx, (txBody, type, idx) => {
		const bodyPr = firstChild(txBody, 'a:bodyPr')
		if (bodyPr && attr(bodyPr, 'anchor') != null) return // slide fixes it — not inherited
		const anchor = placeholderInheritedAnchor(type, idx, ctx)
		if (!anchor) return
		setAttr(bodyPr ?? getOrAddChild(txBody, 'a:bodyPr', ['a:lstStyle', 'a:p']), 'anchor', anchor)
	})
}

/**
 * Bake the placeholder-inherited list style onto a lifted shape's text body (the
 * {@link flattenShape} placeholder-demotion path). A placeholder's per-level
 * *paragraph* formatting — indent (`marL`/`indent`), bullets (`a:buChar`/
 * `a:buAutoNum`/`a:buNone`), alignment, and the level `a:defRPr` — is inherited
 * from the source layout/master placeholder `a:lstStyle` and the master
 * `p:txStyles` category style. Run colour/size are baked onto each run by the
 * sibling passes, but paragraph-level defaults live here; once
 * {@link demotePlaceholders} strips the `p:ph` the shape stops inheriting them, so
 * bullets and indents would vanish. We materialize the effective list style onto
 * the shape's `p:txBody/a:lstStyle`.
 *
 * Resolution is per *level*, most specific tier wins whole: the slide's own level
 * (if it defines one) is kept verbatim, else the first source tier that defines it
 * (layout placeholder, then master placeholder, then master category style) is
 * cloned. This is a whole-element overlay, not the per-attribute merge PowerPoint
 * does — but explicit paragraph `a:pPr` on the slide's own runs travels with the
 * shape and still wins, so this only supplies defaults for the inherited case
 * (paragraphs that set no `a:pPr` of their own). Scheme colours in the cloned
 * levels are resolved by the later `resolveSchemeColors` pass.
 */
function resolvePlaceholderListStyle(shapeRoot: Element, ctx: FlattenContext): void {
	forEachPlaceholderTextBody(shapeRoot, ctx, (txBody, type, idx) => {
		const tiers = placeholderInheritedListStyles(type, idx, ctx)
		if (tiers.length === 0) return
		const slideLst = firstChild(txBody, 'a:lstStyle')
		const merged = createElement(ownerDocumentOf(txBody), 'a:lstStyle')
		let any = false
		for (const level of LST_STYLE_LEVELS) {
			// Slide's own level wins; otherwise the most-specific source tier that defines it.
			const own = slideLst && firstChild(slideLst, level)
			const src = own ?? tiers.map((t) => firstChild(t, level)).find((e): e is Element => !!e) ?? null
			if (!src) continue
			merged.appendChild(src.cloneNode(true))
			any = true
		}
		if (!any) return
		if (slideLst) txBody.replaceChild(merged, slideLst)
		else insertInOrder(txBody, merged, ['a:p'])
	})
}

/**
 * Strip the `p:ph` marker from every placeholder shape in a lifted subtree so it
 * becomes a self-contained ordinary shape (the {@link flattenShape} path only). By
 * the time this runs, the shape's inherited geometry, colour, run size, anchor, and
 * list style are all baked explicitly, so it no longer needs — and must not keep —
 * its placeholder identity: a surviving `p:ph` would re-resolve against the *host*
 * deck's layout/master placeholder of the same type/idx (wrong inheritance, or a
 * fallback when the host has none) and could collide with the host slide's own
 * placeholder of that type. Removing the marker severs both. The `p:nvPr` itself is
 * kept (it can carry media/custom-data children); only the `p:ph` child goes.
 */
function demotePlaceholders(shapeRoot: Element): void {
	for (const sp of descendantsByTag(shapeRoot, OOXML_NS.p, 'sp')) {
		const nvPr = nvPrOf(sp)
		if (nvPr) removeChildrenByQName(nvPr, ['p:ph'])
	}
}

/**
 * Bake placeholder-inherited run size/weight onto the slide. Mirrors
 * {@link resolvePlaceholderRunColors}: for each placeholder run that sets no
 * `sz`/`b`/`i` of its own (nor at paragraph/text-body level on the *slide*),
 * resolve the value it would inherit from the source `slideLayout`/`slideMaster`
 * text styles — per paragraph list level — and write it explicitly onto the run's
 * `a:rPr`. Each property resolves independently up the chain. Typeface (`a:latin`)
 * is left to re-bind to the destination theme, as the colour pass does for the `fontRef`.
 */
function resolvePlaceholderRunSizes(slideRoot: Element, ctx: FlattenContext): void {
	bakePlaceholderRunProperty(slideRoot, ctx, placeholderInheritedRunProps, writeRunProps)
}

/** Write each resolved run property onto a run's `a:rPr`, skipping ones the slide already fixes. */
function writeRunProps(
	run: Element,
	props: RunProps,
	pPr: Element | null,
	slideLst: Element | null,
	level: number
): void {
	let rPr: Element | null = null
	for (const name of RUN_PROP_NAMES) {
		const value = props[name]
		if (value == null) continue
		if (slideDefinesProp(name, run, pPr, slideLst, level)) continue
		rPr ??= getOrAddChild(run, 'a:rPr', ['a:t'])
		setAttr(rPr, name, value)
	}
}

/**
 * Resolve each shape's `p:style` `lnRef`/`fillRef`/`effectRef` into explicit
 * `spPr` children (using the theme `fmtScheme`), then neutralize the ref. The
 * `fontRef` is intentionally left for the destination theme to re-resolve.
 */
function materializeStyleRefs(root: Element, ctx: FlattenContext): void {
	if (!ctx.fmtScheme) return
	for (const style of descendantsByTag(root, OOXML_NS.p, 'style')) {
		const shape = style.parentNode as Element | null
		if (!shape) continue
		const spPr = getOrAddChild(shape, 'p:spPr', SHAPE_AFTER_SPPR)
		materializeFill(spPr, firstChild(style, 'a:fillRef'), ctx)
		materializeLine(spPr, firstChild(style, 'a:lnRef'), ctx)
		materializeEffect(spPr, firstChild(style, 'a:effectRef'), ctx)
	}
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
	const idx = numberAttr(effectRef, 'idx')
	if (idx !== null && idx > 0 && !firstChild(spPr, 'a:effectLst') && !firstChild(spPr, 'a:effectDag')) {
		const style = fmtEntry(ctx, 'a:effectStyleLst', idx) // a:effectStyle (effectLst?, scene3d?, sp3d?)
		const ref = resolveColor(firstChildElement(effectRef), ctx)
		if (style && ref) {
			substitutePhClr(style, ref)
			// Lift the effectStyle's children (effectLst/scene3d/sp3d) into spPr, in order.
			for (const child of childElements(style)) {
				if (isA(child, 'effectLst') || isA(child, 'effectDag')) insertInOrder(spPr, child, SPPR_EFFECT_AFTER)
				else if (isA(child, 'scene3d')) insertInOrder(spPr, child, SPPR_SCENE3D_AFTER)
				else if (isA(child, 'sp3d')) insertInOrder(spPr, child, SPPR_SP3D_AFTER)
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
