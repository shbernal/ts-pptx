/**
 * Resolving what a placeholder shape *inherits* from its source layout/master chain.
 *
 * A placeholder that sets none of its own colour, size, geometry, anchor or list style takes
 * each from a chain: the layout placeholder's `a:lstStyle`, then the master placeholder's, then
 * the master's `p:txStyles` category style (and, on a notes context, the notesMaster
 * `p:notesStyle`). Every function here answers "what would this placeholder get?" — none of them
 * writes anything.
 *
 * **Two callers, one answer.** The read model's getters (`Run.resolvedColor`,
 * `TextFrame.resolvedAnchor`, the font getters) and the import-time flatten passes
 * (`read/api/ops/flatten.ts`, which bakes the inherited values onto a slide so they survive a
 * rebind to a different master) both ask the same questions. They must not answer them
 * differently: if the getter and the bake disagreed, a caller would see one colour before export
 * and another in the file. Sharing this module is what makes that structurally impossible.
 *
 * This module is pure — mutation lives in the flatten pass, which is also why it sits in
 * `read/oxml/` (substrate) rather than `read/api/ops/`.
 */

import { attr, descendantsByTag, firstChild, firstChildElement, OOXML_NS, type Element } from './dom.js'
import { FILL_CHOICES } from './fill.js'
import { resolveColor, type ResolvedColor, type ThemeContext } from './theme.js'

/** The master `p:txStyles` style element name for a placeholder category. */
export const TX_STYLE_NAME: Record<'title' | 'body' | 'other', string> = {
	title: 'p:titleStyle',
	body: 'p:bodyStyle',
	other: 'p:otherStyle',
}

/** The ordered children of a `CT_TextListStyle` (`a:lstStyle` / a `p:txStyles` style). */
export const LST_STYLE_LEVELS = [
	'a:defPPr',
	'a:lvl1pPr',
	'a:lvl2pPr',
	'a:lvl3pPr',
	'a:lvl4pPr',
	'a:lvl5pPr',
	'a:lvl6pPr',
	'a:lvl7pPr',
	'a:lvl8pPr',
	'a:lvl9pPr',
]

/** Inheritable run properties baked under `preserve`: size and weight/slant (not typeface). */
export const RUN_PROP_NAMES = ['sz', 'b', 'i'] as const
export type RunProps = Record<(typeof RUN_PROP_NAMES)[number], string | null>

/** The `p:ph` element of a shape (`p:sp/p:nvSpPr/p:nvPr/p:ph`), or `null`. */
export function placeholderOf(sp: Element): Element | null {
	const nvSpPr = firstChild(sp, 'p:nvSpPr')
	const nvPr = nvSpPr && firstChild(nvSpPr, 'p:nvPr')
	return nvPr ? firstChild(nvPr, 'p:ph') : null
}

/**
 * The master text-style *category* a placeholder type resolves against
 * (absent ⇒ `obj` ⇒ body). This is a `p:txStyles` selector — it picks
 * `p:titleStyle`/`p:bodyStyle`/`p:otherStyle` (see {@link TX_STYLE_NAME}) and
 * deliberately collapses `dt`/`ftr`/`sldNum`/`hdr` (and anything else) into
 * `'other'`. It is **not** a placeholder *identity* predicate: it must not be
 * used to decide which source placeholder a slide placeholder inherits geometry
 * from, or the footer trio become mutually interchangeable — use the exact
 * `type` match in {@link findPlaceholder} for that.
 */
export function phCategory(type: string | null): 'title' | 'body' | 'other' {
	if (type === 'title' || type === 'ctrTitle') return 'title'
	if (type === null || type === 'body' || type === 'subTitle' || type === 'obj') return 'body'
	return 'other'
}

/**
 * Placeholder types of which a layout/master holds at most one, and whose
 * identity is their `type` alone: `dt`, `ftr`, `sldNum`, `hdr`. A slide
 * placeholder of one of these types must inherit **only** from a source
 * placeholder of the *same type* — never via `idx` or the txStyles *category*,
 * both of which lump the whole trio together and would let `sldNum` borrow the
 * footer's box (and vice versa).
 */
const SINGLETON_PH = new Set(['dt', 'ftr', 'sldNum', 'hdr'])

/**
 * The placeholder shape in `root` (a layout/master) that the given slide
 * placeholder inherits from.
 *
 * A singleton-type slide placeholder (`dt`/`ftr`/`sldNum`/`hdr`) matches only an
 * exact same-`type` source placeholder — PowerPoint gives the trio different
 * `idx` on the layout (dt=10/ftr=11/sldNum=12) than the master (dt=2/ftr=3/
 * sldNum=4), so an `idx`- or category-based fallback would silently pick the
 * wrong member of the trio. Every other type prefers a same-`idx` placeholder of
 * the same category, then any same-`idx`, then any same-category, and never
 * lands on a singleton placeholder. Returns `null` when nothing matches.
 */
export function findPlaceholder(root: Element, slideType: string | null, slideIdx: string): Element | null {
	if (slideType && SINGLETON_PH.has(slideType)) {
		for (const sp of descendantsByTag(root, OOXML_NS.p, 'sp')) {
			const ph = placeholderOf(sp)
			if (ph && attr(ph, 'type') === slideType) return sp
		}
		return null
	}
	const cat = phCategory(slideType)
	let idxMatch: Element | null = null
	let catMatch: Element | null = null
	for (const sp of descendantsByTag(root, OOXML_NS.p, 'sp')) {
		const ph = placeholderOf(sp)
		if (!ph) continue
		const type = attr(ph, 'type')
		if (type && SINGLETON_PH.has(type)) continue // a non-singleton must never inherit from the footer trio
		const i = attr(ph, 'idx') ?? '0'
		const sameCat = phCategory(type) === cat
		if (i === slideIdx && sameCat) return sp
		if (i === slideIdx && !idxMatch) idxMatch = sp
		if (sameCat && !catMatch) catMatch = sp
	}
	return idxMatch ?? catMatch
}

/** The `a:lstStyle` of a placeholder shape's `p:txBody`, or `null`. */
export function placeholderLstStyle(sp: Element): Element | null {
	const txBody = firstChild(sp, 'p:txBody')
	return txBody ? firstChild(txBody, 'a:lstStyle') : null
}

/**
 * The `a:defRPr` for `level` (0-based) in a `CT_TextListStyle` (`a:lstStyle` or a
 * `p:txStyles` style): the level-specific `a:lvlNpPr/a:defRPr`, else the
 * `a:defPPr/a:defRPr` fallback. The shared root for colour and size resolution.
 */
export function lstStyleLevelDefRPr(listStyle: Element | null, level: number): Element | null {
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
 * The source style tiers a placeholder resolves against, most specific first: the layout
 * placeholder's `a:lstStyle`, the master placeholder's, then the master `p:txStyles` category
 * style — and, on a notes context only, the notesMaster `p:notesStyle`.
 *
 * The one place the chain is spelled out. `placeholderInherited*` differ only in what they pull
 * from each tier, so they all walk this.
 */
function styleTiers(type: string | null, idx: string, ctx: ThemeContext): Element[] {
	const tiers: (Element | null)[] = []
	if (ctx.layoutRoot) {
		const layoutPh = findPlaceholder(ctx.layoutRoot, type, idx)
		tiers.push(layoutPh && placeholderLstStyle(layoutPh))
	}
	if (ctx.masterRoot) {
		const masterPh = findPlaceholder(ctx.masterRoot, type, idx)
		tiers.push(masterPh && placeholderLstStyle(masterPh))
		const txStyles = firstChild(ctx.masterRoot, 'p:txStyles')
		tiers.push(txStyles && firstChild(txStyles, TX_STYLE_NAME[phCategory(type)]))
	}
	// Notes body runs inherit from the notesMaster's `p:notesStyle` (keyed by level, not
	// placeholder type); it is the bottom tier of the notes chain — see the field note on
	// `ThemeContext.notesStyle`. Never set on a slide context.
	if (ctx.notesStyle) tiers.push(ctx.notesStyle)
	return tiers.filter((t): t is Element => t !== null)
}

/**
 * The colour *element* a placeholder run inherits from the source style chain. Returns the
 * first tier's colour element that resolves against `ctx` (the `a:srgbClr`/`a:schemeClr`/…
 * inside its `a:solidFill`), or `null` when nothing in the chain defines a resolvable colour.
 * The read-model `Run.resolvedColor` getter feeds this element to `resolveColorElement` for the
 * `effectiveHex`; the flatten path resolves it directly via {@link placeholderInheritedColor}.
 */
export function placeholderInheritedFill(
	type: string | null,
	idx: string,
	level: number,
	ctx: ThemeContext
): Element | null {
	for (const tier of styleTiers(type, idx, ctx)) {
		const fill = lstStyleLevelFill(tier, level)
		const colorEl = fill && firstChildElement(fill)
		if (colorEl && resolveColor(colorEl, ctx)) return colorEl
	}
	return null
}

/**
 * The level `a:defRPr` elements a placeholder inherits from the source style chain, in
 * resolution order. Tiers with no `a:defRPr` for `level` are dropped. The shared root for
 * inherited run *size* and *typeface* resolution (the size/face sibling of
 * {@link placeholderInheritedFill}), read directly by the flatten path and the read-model font
 * getters.
 */
export function placeholderInheritedDefRPrs(
	type: string | null,
	idx: string,
	level: number,
	ctx: ThemeContext
): Element[] {
	return styleTiers(type, idx, ctx)
		.map((tier) => lstStyleLevelDefRPr(tier, level))
		.filter((t): t is Element => t !== null)
}

/**
 * The source list-style tiers a placeholder inherits paragraph formatting from, most specific
 * first. The paragraph-level sibling of {@link placeholderInheritedDefRPrs} — it hands back the
 * tier elements whole rather than pulling one level's properties out of each.
 *
 * The notes tier is included for consistency with the rest of the chain, though it cannot arise
 * in practice: the only caller is the list-style bake in `flattenShape`, which runs on slide and
 * shape contexts, and `ctx.notesStyle` is set only on a notes context.
 */
export function placeholderInheritedListStyles(type: string | null, idx: string, ctx: ThemeContext): Element[] {
	return styleTiers(type, idx, ctx)
}

/**
 * The colour a placeholder run inherits from the source style chain, resolved to a
 * literal `{ hex, transforms }`. Thin wrapper over {@link placeholderInheritedFill}
 * for the flatten path, which re-emits the transforms verbatim. Returns `null`
 * when nothing in the chain defines a resolvable colour (the run then re-binds to
 * the destination).
 */
export function placeholderInheritedColor(
	type: string | null,
	idx: string,
	level: number,
	ctx: ThemeContext
): ResolvedColor | null {
	const colorEl = placeholderInheritedFill(type, idx, level, ctx)
	return colorEl ? resolveColor(colorEl, ctx) : null
}

/**
 * The run size/weight a placeholder run inherits from the source style chain, per list level.
 * Each of `sz`/`b`/`i` is taken from the first tier that defines it (properties resolve
 * independently). Returns `null` when no tier defines any of them.
 */
export function placeholderInheritedRunProps(
	type: string | null,
	idx: string,
	level: number,
	ctx: ThemeContext
): RunProps | null {
	const tiers = placeholderInheritedDefRPrs(type, idx, level, ctx)
	const props = {} as RunProps
	let any = false
	for (const name of RUN_PROP_NAMES) {
		let value: string | null = null
		for (const tier of tiers) {
			value = attr(tier, name)
			if (value != null) break
		}
		props[name] = value
		if (value != null) any = true
	}
	return any ? props : null
}

/**
 * The `a:xfrm` a placeholder inherits from the source layout, then master,
 * tagged with which tier it came from — `null` when neither defines one. The
 * read-model `resolveInheritedFrame` (`theme-context.ts`) sibling of
 * {@link placeholderInheritedAnchor}; also the bake source for the flatten pass'
 * `resolvePlaceholderGeometry`.
 */
export function placeholderInheritedXfrm(
	type: string | null,
	idx: string,
	ctx: ThemeContext
): { xfrm: Element; source: 'layout' | 'master' } | null {
	for (const [root, source] of [
		[ctx.layoutRoot, 'layout'],
		[ctx.masterRoot, 'master'],
	] as const) {
		if (!root) continue
		const ph = findPlaceholder(root, type, idx)
		const spPr = ph && firstChild(ph, 'p:spPr')
		const xfrm = spPr && firstChild(spPr, 'a:xfrm')
		if (xfrm) return { xfrm, source }
	}
	return null
}

/**
 * The vertical anchor (`a:bodyPr/@anchor`) a placeholder inherits from the source
 * layout, then master, or `null`. Resolves per-attribute: the first tier whose
 * placeholder `a:bodyPr` actually sets `@anchor` wins, so a layout `a:bodyPr`
 * present but without `@anchor` does not mask the master's. The read-model
 * `TextFrame.resolvedAnchor` sibling of {@link placeholderInheritedXfrm}.
 */
export function placeholderInheritedAnchor(type: string | null, idx: string, ctx: ThemeContext): string | null {
	for (const root of [ctx.layoutRoot, ctx.masterRoot]) {
		if (!root) continue
		const ph = findPlaceholder(root, type, idx)
		const txBody = ph && firstChild(ph, 'p:txBody')
		const bodyPr = txBody && firstChild(txBody, 'a:bodyPr')
		const anchor = bodyPr && attr(bodyPr, 'anchor')
		if (anchor) return anchor
	}
	return null
}

/** Whether the *slide itself* already fixes this run's colour (so a rebind cannot change it). */
export function slideDefinesColor(run: Element, pPr: Element | null, slideLst: Element | null, level: number): boolean {
	const rPr = firstChild(run, 'a:rPr')
	if (rPr && FILL_CHOICES.some((q) => firstChild(rPr, q))) return true
	const defRPr = pPr && firstChild(pPr, 'a:defRPr')
	if (defRPr && firstChild(defRPr, 'a:solidFill')) return true
	return !!lstStyleLevelFill(slideLst, level)
}

/** Whether the *slide itself* already fixes a run property (so a rebind cannot change it). */
export function slideDefinesProp(
	name: string,
	run: Element,
	pPr: Element | null,
	slideLst: Element | null,
	level: number
): boolean {
	const rPr = firstChild(run, 'a:rPr')
	if (rPr && attr(rPr, name) != null) return true
	const defRPr = pPr && firstChild(pPr, 'a:defRPr')
	if (defRPr && attr(defRPr, name) != null) return true
	const slideDefRPr = lstStyleLevelDefRPr(slideLst, level)
	return !!(slideDefRPr && attr(slideDefRPr, name) != null)
}
