/**
 * Theme colour resolution: turning a DrawingML colour reference or a style-matrix reference into
 * something literal.
 *
 * Two maps do the colour work. A scheme token (`accent1`, `bg1`, …) routes through the master's
 * `clrMap` to a `clrScheme` slot, whose literal RGB is the answer; child colour transforms
 * (`lumMod`/`shade`/`alpha`/…) are carried through untouched rather than computed, so a caller
 * that re-emits them renders identically. A style-matrix reference (`a:fillRef`/`a:lnRef`) instead
 * indexes the theme's `fmtScheme`, and resolves to a deep-cloned entry with its `phClr`
 * placeholder substituted.
 *
 * Everything here is **pure**: it reads a DOM and builds detached elements, and mutates nothing
 * the caller passes in. Two very different consumers depend on that:
 *
 * - the read model's colour getters, which resolve a colour to report it;
 * - the import-time flatten passes (`read/api/ops/flatten.ts`), which resolve the same colour in
 *   order to *bake* it into a slide.
 *
 * They must agree, so they share this. The passes that write results back into a part live with
 * the other import operations, not here — this module deliberately has no mutating exports.
 * Placeholder-inheritance resolution (what a placeholder gets from its layout/master chain) is
 * its own concern and lives in `placeholder-inherit.ts`.
 */
import {
	OOXML_NS,
	attr,
	childElements,
	createElement,
	descendantsByTag,
	firstChild,
	firstChildElement,
	ownerDocumentOf,
	replaceInParent,
	setAttr,
	type Element,
} from './dom.js'

/** The 12 `a:clrScheme` slot names, in schema order. */
export const SCHEME_SLOTS = [
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

/** Scheme tokens that name a `clrScheme` slot directly, bypassing the `clrMap`. */
const DIRECT_SLOT_TOKENS = new Set(['dk1', 'lt1', 'dk2', 'lt2'])

/**
 * The colour-resolution context: the two maps that turn a DrawingML colour
 * reference into a literal hex — the effective colour map (token → `clrScheme`
 * slot, honouring any slide `clrMapOvr`) and the resolved colour scheme (slot →
 * 6-hex RGB). Shared by the read-model colour getters and {@link ThemeContext}.
 */
export interface ColorContext {
	clrMap: Map<string, string>
	clrScheme: Map<string, string>
}

/**
 * The source theme subgraph a slide, notes slide, layout or master resolves against: the
 * {@link ColorContext} maps, the live `a:fmtScheme` for style-matrix resolution, and the roots
 * of the inheritance chain a placeholder reads through.
 *
 * Every root here is **read-only** — nothing in this module or in `placeholder-inherit.ts`
 * mutates one. The import-time flatten pass extends this with the one field it needs to write
 * from (see `FlattenContext` in `read/api/ops/flatten.ts`).
 */
export interface ThemeContext extends ColorContext {
	fmtScheme: Element | null
	/**
	 * The source `slideLayout` root element, for resolving what a placeholder inherits.
	 * `null`/absent disables the layout tier.
	 */
	layoutRoot?: Element | null
	/**
	 * The source `slideMaster` root element, for resolving placeholder inheritance via its
	 * placeholder `a:lstStyle` and `p:txStyles`.
	 */
	masterRoot?: Element | null
	/**
	 * The theme's `a:fontScheme`, for resolving a `+mj-*`/`+mn-*` major/minor font
	 * token (the placeholder-inherited typeface chain bottoms out in one) to a
	 * literal face name via {@link resolveThemeFont}. `null`/absent leaves such a
	 * token unresolved.
	 */
	fontScheme?: Element | null
	/**
	 * The notesMaster's `p:notesStyle` (a `CT_TextListStyle`), for resolving the
	 * character properties a *notes-body* run inherits when it sets none of its own.
	 * Only ever set on a notes context (see `resolveNotesColorContext`); on a slide
	 * context it is absent, so the slide placeholder chain (layout/master
	 * `a:lstStyle` → `p:txStyles`) is unaffected. It is the notes analogue of the
	 * master `p:txStyles` category style, keyed by level rather than placeholder
	 * type. Read-only — never mutated.
	 */
	notesStyle?: Element | null
	/**
	 * The presentation's `p:defaultTextStyle` (a `CT_TextListStyle` from
	 * `presentation.xml`), keyed by level — PowerPoint's lowest-priority text
	 * fallback, applying to any run (placeholder or not) that resolves nothing
	 * above it. Only ever set on a *slide* context (see `resolveSlideColorContext`);
	 * absent on notes/master/layout contexts, so their chains are unaffected.
	 * Read-only — never mutated.
	 */
	defaultTextStyle?: Element | null
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
	for (const token of [
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
	]) {
		const slot = attr(clrMap, token)
		if (slot) out.set(token, slot)
	}
	return out
}

/** Resolve a scheme token (`accent1`, `bg1`, `dk1`, …) to a 6-hex RGB, or `null`. */
export function resolveSchemeToken(token: string, ctx: ColorContext): string | null {
	if (token === 'phClr') return null
	const slot = DIRECT_SLOT_TOKENS.has(token) ? token : ctx.clrMap.get(token)
	return slot ? (ctx.clrScheme.get(slot) ?? null) : null
}

/** Whether `el` is a DrawingML element with the given local name. */
export function isA(el: Element | null, local: string): boolean {
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
 * Resolve a `+mj-*`/`+mn-*` major/minor font token against a theme `a:fontScheme`
 * to a literal typeface name. A non-token `typeface` (an already-literal face) is
 * returned verbatim. `null` when `typeface` is absent, or when a token cannot be
 * resolved (no `fontScheme`, or the indexed major/minor script slot is empty).
 * The script suffix selects the scheme child: `lt`→`a:latin`, `ea`→`a:ea`,
 * `cs`→`a:cs`.
 */
export function resolveThemeFont(typeface: string | null, fontScheme: Element | null): string | null {
	if (!typeface) return null
	const match = /^\+(mj|mn)-(lt|ea|cs)$/.exec(typeface)
	if (!match) return typeface
	if (!fontScheme) return null
	const font = firstChild(fontScheme, match[1] === 'mj' ? 'a:majorFont' : 'a:minorFont')
	const childName = match[2] === 'lt' ? 'a:latin' : match[2] === 'ea' ? 'a:ea' : 'a:cs'
	const child = font && firstChild(font, childName)
	const resolved = child && attr(child, 'typeface')
	return resolved || null
}

/** Replace every `phClr` under `el` with the ref colour (ref transforms first, then the `phClr`'s own). */
export function substitutePhClr(el: Element, ref: ResolvedColor): void {
	const doc = ownerDocumentOf(el)
	for (const phClr of descendantsByTag(el, OOXML_NS.a, 'schemeClr')) {
		if (attr(phClr, 'val') !== 'phClr') continue
		const srgb = createElement(doc, 'a:srgbClr')
		setAttr(srgb, 'val', ref.hex)
		for (const t of ref.transforms) srgb.appendChild(t.cloneNode(true))
		while (phClr.firstChild) srgb.appendChild(phClr.firstChild)
		replaceInParent(phClr, srgb)
	}
}

/** The `idx`-th entry (1-based) of a `fmtScheme` style list, deep-cloned, or `null`. */
export function fmtEntry(ctx: ThemeContext, listName: string, idx: number): Element | null {
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
 * colour cannot be resolved. Shared by the flatten path and the read-model
 * `resolveStyleFillColor` getter so both see the same fill.
 */
export function styleRefFill(fillRef: Element | null, ctx: ThemeContext): Element | null {
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
export function styleRefLine(lnRef: Element | null, ctx: ThemeContext): Element | null {
	if (!lnRef) return null
	const idx = intAttr(lnRef, 'idx')
	if (idx === null || idx <= 0) return null
	const ln = fmtEntry(ctx, 'a:lnStyleLst', idx)
	const ref = resolveColor(firstChildElement(lnRef), ctx)
	if (!ln || !ref) return null
	substitutePhClr(ln, ref)
	return ln
}

/** Read an integer attribute; `null`/empty/non-finite → `null`. */
export function intAttr(el: Element, name: string): number | null {
	const value = attr(el, name)
	if (value === null || value === '') return null
	const n = Number(value)
	return Number.isFinite(n) ? n : null
}
