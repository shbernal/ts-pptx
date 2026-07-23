/**
 * Read a slide's effective background (`p:cSld/p:bg`), decoded into a typed union.
 *
 * A background lives on the slide itself, or is inherited from its slideLayout,
 * else its slideMaster — the {@link SlideBackground.source} field records which.
 * The writer authors solid-colour, gradient, and image backgrounds (all FAITHFUL);
 * pattern and theme-indexed (`p:bgRef`) backgrounds are read-only for imported
 * decks — surface them so a codegen pass can carry the part rather than fake it.
 */
import { attr, firstChild, firstChildElement, intValue, type Element } from '../oxml/dom.js'
import type { ColorContext } from '../oxml/theme.js'
import type { Relationships } from '../opc/relationships.js'
import { resolveColorElement, type ResolvedColor } from './theme-context.js'
import { readGradientFill, type GradientFill } from './gradient.js'

/** Where a slide's effective background comes from in the slide → layout → master chain. */
export type BackgroundSource = 'slide' | 'layout' | 'master'

/**
 * A slide's effective background (`p:cSld/p:bg`), as a discriminated union on
 * `type`. Every variant carries {@link BackgroundSource}. `solid`/`gradient`/`image`
 * are what the writer authors and round-trip faithfully; `pattern`/`themeRef` are
 * read-only (imported decks) and `none` is an explicit `a:noFill`.
 */
export type SlideBackground =
	| { type: 'solid'; source: BackgroundSource; color: ResolvedColor | null }
	| { type: 'gradient'; source: BackgroundSource; gradient: GradientFill }
	| { type: 'image'; source: BackgroundSource; relId: string | null; partName: string | null }
	| {
			type: 'pattern'
			source: BackgroundSource
			preset: string | null
			foreground: ResolvedColor | null
			background: ResolvedColor | null
	  }
	| { type: 'themeRef'; source: BackgroundSource; idx: number | null; color: ResolvedColor | null }
	| { type: 'none'; source: BackgroundSource }

/** The `p:cSld/p:bg` element of a slide/layout/master root, or `null`. */
export function backgroundElementOf(root: Element | null): Element | null {
	const cSld = root && firstChild(root, 'p:cSld')
	return cSld ? firstChild(cSld, 'p:bg') : null
}

/**
 * Decode a `p:bg` element into a {@link SlideBackground}. `ctx` resolves colour
 * tokens; `rels` (the *owning* part's relationships) resolves an image
 * background's `r:embed` to an absolute part name.
 */
export function readSlideBackground(
	bg: Element,
	source: BackgroundSource,
	ctx: ColorContext,
	rels: Relationships | null
): SlideBackground {
	// A background is either a `p:bgPr` (explicit fill) or a `p:bgRef` (theme-indexed).
	const bgRef = firstChild(bg, 'p:bgRef')
	if (bgRef) {
		return {
			type: 'themeRef',
			source,
			idx: intValue(attr(bgRef, 'idx')),
			color: resolveColorElement(firstChildElement(bgRef), ctx),
		}
	}

	const bgPr = firstChild(bg, 'p:bgPr')
	if (!bgPr) return { type: 'none', source }

	const solid = firstChild(bgPr, 'a:solidFill')
	if (solid) return { type: 'solid', source, color: resolveColorElement(firstChildElement(solid), ctx) }

	const grad = firstChild(bgPr, 'a:gradFill')
	if (grad) return { type: 'gradient', source, gradient: readGradientFill(bgPr, ctx) as GradientFill }

	const blip = firstChild(bgPr, 'a:blipFill')
	if (blip) {
		const blipEl = firstChild(blip, 'a:blip')
		const relId = blipEl ? attr(blipEl, 'r:embed') : null
		return { type: 'image', source, relId, partName: relId && rels ? rels.resolveTarget(relId) : null }
	}

	const patt = firstChild(bgPr, 'a:pattFill')
	if (patt) {
		const wrapColor = (qname: string): ResolvedColor | null => {
			const wrap = firstChild(patt, qname)
			const colorEl = wrap && firstChildElement(wrap)
			return colorEl ? resolveColorElement(colorEl, ctx) : null
		}
		return {
			type: 'pattern',
			source,
			preset: attr(patt, 'prst') ?? null,
			foreground: wrapColor('a:fgClr'),
			background: wrapColor('a:bgClr'),
		}
	}

	// `a:noFill` (or an empty/unrecognized bgPr) is an explicit transparent background.
	return { type: 'none', source }
}
