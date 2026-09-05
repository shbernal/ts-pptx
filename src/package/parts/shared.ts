/**
 * ts-pptx: the part-contributor seam
 *
 * What a construct family adds to a written package: the parts it puts in the zip, and the
 * `[Content_Types].xml` entries those parts need in order to resolve. The packager
 * (`package/assemble.ts`) owns the skeleton every deck has — content types, the rels graph,
 * docProps, theme, slides, layouts, the master, media — and calls a list of contributors at fixed
 * points in it, so it never names a family.
 *
 * It has to be a list rather than imports, for the same reason the shape walk is handed its
 * renderers (`gen/slide/objects/shared.ts`): a named import inside a reachable function body is
 * retained unconditionally, so a packager that calls `createExcelWorksheet` itself links all of
 * `gen/chart/` — the plot modules, the axes, the chartEx sidecars — into every program that writes
 * a deck, text-only ones included. Same argument, the other axis: that one is about which XML
 * lands in a slide, this one about which parts land in the zip.
 *
 * Nothing else belongs in this module. A contributor's emitters live with the family they belong
 * to under `gen/`; the contributors themselves are one small module each, in this directory.
 */

import type { ZipWriter } from '../../zip.js'
import type { ContentTypeContributions, ContentTypeDefault, ContentTypeOverride } from '../../gen/opc/content-types.js'
import type { PresentationPropsInternal, PresSlideInternal, SlideLayoutInternal } from '../../types/internal.js'

/** A slide, a layout, or the master — anything carrying the rel arrays a part can be derived from. */
export type PartTarget = PresSlideInternal | SlideLayoutInternal

/**
 * One construct family's contribution to the package.
 *
 * Every hook is optional and every hook is called at one fixed point in the emission order, which
 * is what keeps zip part order and `[Content_Types].xml` byte-identical no matter which families a
 * program links.
 */
export interface PartContributor {
	/**
	 * Rank among contributors that write into the same slot. Fixed here, with the family, rather
	 * than taken from the order a caller happens to list contributors in: the list is a tier's
	 * business, and the same deck has to come out byte-identical from every tier that can write it.
	 * Spaced by ten so a family can be slotted between two others without renumbering.
	 */
	readonly order: number
	/** Parts this family writes into the package. */
	readonly parts?: ContributedParts
	/** Content-type entries this family's parts need to resolve. */
	readonly contentTypes?: ContributedContentTypes
}

/** Where a family's parts go into the zip. Order within each hook is the order they are added. */
export interface ContributedParts {
	/** Written right after slide `slideNumber`'s own part and its rels, inside the slide loop. */
	readonly withEachSlide?: (slide: PresSlideInternal, slideNumber: number, zip: ZipWriter) => void
	/** Written once, after the slide master and its rels. */
	readonly afterMaster?: (pres: PresentationPropsInternal, zip: ZipWriter) => void
	/**
	 * Written from one target's rel arrays, for every layout, then every slide, then the master —
	 * last, alongside the media parts. Any promises returned join the single pool the packager
	 * awaits, so a family with async part work does not serialize the write.
	 */
	readonly fromTargetRels?: (target: PartTarget, zip: ZipWriter) => Promise<unknown>[] | void
}

/**
 * Which `[Content_Types].xml` entries a family needs. Each hook names the slot its entries land
 * in; see {@link ContentTypeContributions} for where those slots sit in the emitted order.
 */
export interface ContributedContentTypes {
	readonly defaults?: (pres: PresentationPropsInternal) => ContentTypeDefault[]
	readonly presentation?: (pres: PresentationPropsInternal) => ContentTypeOverride[]
	readonly perSlide?: (slide: PresSlideInternal, slideNumber: number) => ContentTypeOverride[]
	readonly perLayout?: (layout: SlideLayoutInternal, layoutNumber: number) => ContentTypeOverride[]
	readonly trailing?: (pres: PresentationPropsInternal) => ContentTypeOverride[]
}

/**
 * The contributors in emission order. Sorting here rather than trusting the caller's array is the
 * whole point of {@link PartContributor.order}; the sort is stable, so two families that share a
 * rank keep their listed order rather than swapping unpredictably.
 */
export function orderedContributors(contributors: readonly PartContributor[]): PartContributor[] {
	return [...contributors].sort((a, b) => a.order - b.order)
}

/**
 * Ask every contributor for its content-type entries and flatten them into the slot-keyed data
 * `makeXmlContTypes` emits. Takes the contributors already ordered — within a slot the entries
 * come out in the order they are collected here.
 */
export function collectContentTypes(
	contributors: readonly PartContributor[],
	pres: PresentationPropsInternal
): ContentTypeContributions {
	const defaults: ContentTypeDefault[] = []
	const presentation: ContentTypeOverride[] = []
	const trailing: ContentTypeOverride[] = []
	const perSlide: ContentTypeOverride[][] = pres.slides.map(() => [])
	const perLayout: ContentTypeOverride[][] = pres.slideLayouts.map(() => [])

	for (const contributor of contributors) {
		const hooks = contributor.contentTypes
		if (!hooks) continue
		if (hooks.defaults) defaults.push(...hooks.defaults(pres))
		if (hooks.presentation) presentation.push(...hooks.presentation(pres))
		if (hooks.perSlide) {
			const perSlideHook = hooks.perSlide
			pres.slides.forEach((slide, idx) => perSlide[idx]?.push(...perSlideHook(slide, idx + 1)))
		}
		if (hooks.perLayout) {
			const perLayoutHook = hooks.perLayout
			pres.slideLayouts.forEach((layout, idx) => perLayout[idx]?.push(...perLayoutHook(layout, idx + 1)))
		}
		if (hooks.trailing) trailing.push(...hooks.trailing(pres))
	}

	return { defaults, presentation, perSlide, perLayout, trailing }
}
