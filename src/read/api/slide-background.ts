/**
 * Read a slide's effective background (`p:cSld/p:bg`), decoded into a typed union.
 *
 * A background lives on the slide itself, or is inherited from its slideLayout,
 * else its slideMaster — the {@link SlideBackground.source} field records which.
 * The writer authors solid-colour, gradient, and image backgrounds (all FAITHFUL);
 * pattern and theme-indexed (`p:bgRef`) backgrounds are read-only for imported
 * decks — surface them so a codegen pass can carry the part rather than fake it. A
 * `themeRef` additionally resolves its `idx` through the slide theme's `fmtScheme`
 * to a concrete {@link BackgroundFill} in `resolvedFill`, so a consumer can ask what
 * colour/gradient a theme-indexed background actually renders as.
 */
import {
	attr,
	createElement,
	firstChild,
	firstChildElement,
	intValue,
	ownerDocumentOf,
	type Element,
} from '../oxml/dom.js'
import { styleRefFill, type ThemeContext } from '../oxml/theme.js'
import type { Relationships } from '../opc/relationships.js'
import { resolveColorElement, type ResolvedColor } from './theme-context.js'
import { readGradientFill, type GradientFill } from './gradient.js'
import { readPictureFill, type PictureFill } from './picture-fill.js'
import { readPatternFill } from './pattern-fill.js'

/** Where a slide's effective background comes from in the slide → layout → master chain. */
export type BackgroundSource = 'slide' | 'layout' | 'master'

/**
 * An explicit background fill, decoded from a `p:bgPr` (or from a `p:bgRef`'s
 * resolved `fmtScheme` entry). The source-less core of {@link SlideBackground}: the
 * top-level variants add {@link BackgroundSource}, and it is what a `themeRef`'s
 * {@link SlideBackground.resolvedFill} carries.
 *
 * The `image` variant keeps `relId`/`partName` at the top level — the fields
 * callers already read — and adds the full {@link PictureFill} (stretch/tile
 * geometry, crop, alpha) under `picture`.
 */
export type BackgroundFill =
	| { type: 'solid'; color: ResolvedColor | null }
	| { type: 'gradient'; gradient: GradientFill }
	| { type: 'image'; picture: PictureFill; relId: string | null; partName: string | null }
	| { type: 'pattern'; preset: string | null; foreground: ResolvedColor | null; background: ResolvedColor | null }
	| { type: 'none' }

/**
 * A slide's effective background (`p:cSld/p:bg`), as a discriminated union on
 * `type`. Every variant carries {@link BackgroundSource}. `solid`/`gradient`/`image`
 * are what the writer authors and round-trip faithfully; `pattern`/`themeRef` are
 * read-only (imported decks) and `none` is an explicit `a:noFill`. A `themeRef`
 * keeps its raw `idx` for fidelity and also exposes `resolvedFill` — the concrete
 * fill that `idx` resolves to through the slide theme's `fmtScheme` (`null` when the
 * theme/entry cannot be resolved).
 */
export type SlideBackground =
	| { type: 'solid'; source: BackgroundSource; color: ResolvedColor | null }
	| { type: 'gradient'; source: BackgroundSource; gradient: GradientFill }
	| { type: 'image'; source: BackgroundSource; picture: PictureFill; relId: string | null; partName: string | null }
	| {
			type: 'pattern'
			source: BackgroundSource
			preset: string | null
			foreground: ResolvedColor | null
			background: ResolvedColor | null
	  }
	| {
			type: 'themeRef'
			source: BackgroundSource
			idx: number | null
			color: ResolvedColor | null
			resolvedFill: BackgroundFill | null
	  }
	| { type: 'none'; source: BackgroundSource }

/** The `p:cSld/p:bg` element of a slide/layout/master root, or `null`. */
export function backgroundElementOf(root: Element | null): Element | null {
	const cSld = root && firstChild(root, 'p:cSld')
	return cSld ? firstChild(cSld, 'p:bg') : null
}

/**
 * Decode a colour-bearing container (`p:bgPr`, or a synthetic wrapper around a
 * `fmtScheme` fill entry) into a {@link BackgroundFill}. `a:noFill`, an empty
 * container, or an unrecognized fill all read as `{ type: 'none' }`.
 */
function decodeBackgroundFill(container: Element, ctx: ThemeContext, rels: Relationships | null): BackgroundFill {
	const solid = firstChild(container, 'a:solidFill')
	if (solid) return { type: 'solid', color: resolveColorElement(firstChildElement(solid), ctx) }

	const grad = firstChild(container, 'a:gradFill')
	if (grad) return { type: 'gradient', gradient: readGradientFill(container, ctx) as GradientFill }

	const picture = readPictureFill(container, rels)
	if (picture) return { type: 'image', picture, relId: picture.relId, partName: picture.partName }

	const patt = readPatternFill(container, ctx)
	if (patt) return { type: 'pattern', preset: patt.preset, foreground: patt.foreground, background: patt.background }

	return { type: 'none' }
}

/**
 * Resolve a `p:bgRef`'s `idx` into the concrete {@link BackgroundFill} it renders
 * as, through the slide theme's `fmtScheme`. A `p:bgRef` is a
 * `CT_StyleMatrixReference` — the same shape as a shape's `a:fillRef` — so
 * {@link styleRefFill} builds the indexed `bgFillStyleLst`/`fillStyleLst` entry with
 * its `phClr` substituted by the bgRef's colour, and the resulting fill element is
 * decoded like a `p:bgPr`. `null` when the theme has no `fmtScheme`, the entry is
 * absent, or the colour cannot be resolved.
 *
 * `themeRels` — not the owning slide/layout/master's rels — is what an `a:blipFill`
 * entry's `r:embed` resolves against: the fill element is lifted out of the *theme*
 * part, so its relationship id is scoped to that part. Several stock Office themes
 * (Ion, Facet, …) put a picture in the third `bgFillStyleLst` slot, so a deck using
 * `p:bgRef idx="1003"` hits this.
 */
function resolveThemeRefFill(
	bgRef: Element,
	ctx: ThemeContext,
	themeRels: Relationships | null
): BackgroundFill | null {
	const fill = styleRefFill(bgRef, ctx)
	if (!fill) return null
	// Wrap the bare fill element (e.g. `a:gradFill`) in a synthetic container so the
	// same child-dispatch decode used for `p:bgPr` applies unchanged.
	const wrapper = createElement(ownerDocumentOf(bgRef), 'p:bgPr')
	wrapper.appendChild(fill)
	return decodeBackgroundFill(wrapper, ctx, themeRels)
}

/**
 * Decode a `p:bg` element into a {@link SlideBackground}. `ctx` resolves colour
 * tokens and (for a `themeRef`) the theme `fmtScheme`; `rels` (the *owning* part's
 * relationships) resolves an image background's `r:embed` to an absolute part name.
 *
 * `themeRels` is the theme part's own relationships, used only for a `themeRef`'s
 * {@link SlideBackground.resolvedFill}: that fill comes out of the theme part, so an
 * image entry's `r:embed` is scoped to the theme, not to `rels`.
 */
export function readSlideBackground(
	bg: Element,
	source: BackgroundSource,
	ctx: ThemeContext,
	rels: Relationships | null,
	themeRels: Relationships | null
): SlideBackground {
	// A background is either a `p:bgPr` (explicit fill) or a `p:bgRef` (theme-indexed).
	const bgRef = firstChild(bg, 'p:bgRef')
	if (bgRef) {
		return {
			type: 'themeRef',
			source,
			idx: intValue(attr(bgRef, 'idx')),
			color: resolveColorElement(firstChildElement(bgRef), ctx),
			resolvedFill: resolveThemeRefFill(bgRef, ctx, themeRels),
		}
	}

	const bgPr = firstChild(bg, 'p:bgPr')
	if (!bgPr) return { type: 'none', source }

	// `a:noFill` (or an empty/unrecognized bgPr) is an explicit transparent background.
	const fill = decodeBackgroundFill(bgPr, ctx, rels)
	return { ...fill, source }
}
