/**
 * ts-pptx: the pre-serialization pass
 *
 * Everything that must happen to authored slide state *before* any XML is built. Two callers
 * reach serialization and both need it:
 *
 *   - `package/assemble.ts` — the normal `write`/`stream`/`writeFile` path, building a
 *     complete `.pptx`.
 *   - `PresentationCore.extractSlides()` — serializes slide bodies only, for splicing into
 *     an already-loaded deck via `Presentation.appendSlides()` (`ts-pptx/read`).
 *
 * The second used to reimplement the first inline, kept in step by a comment that said
 * "exactly as ... does". Nothing failed when they drifted: `extractSlides` would simply emit
 * slides that did not match what a normal write produces, and no test compares the two.
 *
 * Two primitives rather than one combined pass, because the package path has work of its own
 * that must land *between* them: it de-duplicates identical media across the deck (which
 * needs the encoded bytes {@link encodeMediaForTargets} produces) and assigns package-unique
 * chart part filenames, both before any slide content is baked. Exposing the halves lets
 * `assemble` keep its established order exactly while still sharing the code.
 *
 * The order across the two is load-bearing:
 *
 * 1. **Media first.** It is the only asynchronous step, and the sync XML pass reads the
 *    `rel.data` it populates.
 * 2. **Backfill placeholders**, so a slide inherits its layout's placeholder objects before
 *    anything measures or serializes them.
 * 3. **Bake measured fit last**, because it measures the text step 2 may have just added and
 *    writes the `fontScale` the sync XML pass then reads.
 *
 * The two functions take different target sets on purpose: media lives on layouts and the
 * master as well as on slides, while placeholder backfill and measured fit apply only to
 * slides. `extractSlides` passes only its slides to both, because it emits no layout or
 * master parts.
 */

import { applyMeasuredFit } from '../measure/fit.js'
import type { FontMetricsRegistry } from '../measure/font-metrics.js'
import type { RuntimeAdapter } from '../runtime/types.js'
import type { PresSlideInternal, SlideLayoutInternal } from '../types/internal.js'
import { addPlaceholdersToSlideLayouts } from './define/placeholder.js'
import { encodeSlideMediaRels } from './media.js'

/**
 * Step 1: load and base64-encode every media rel on the given targets, populating `rel.data`.
 * @param {ReadonlyArray<PresSlideInternal | SlideLayoutInternal>} targets - media-bearing targets (slides, layouts, master)
 * @param {RuntimeAdapter} runtime - runtime adapter used to load media
 * @param {'throw' | 'placeholder'} onMediaError - failure policy for unloadable media
 */
export async function encodeMediaForTargets(
	targets: ReadonlyArray<PresSlideInternal | SlideLayoutInternal>,
	runtime: RuntimeAdapter,
	onMediaError: 'throw' | 'placeholder'
): Promise<void> {
	const promises: Array<Promise<string>> = []
	for (const target of targets) {
		promises.push(...encodeSlideMediaRels(target, runtime, onMediaError))
	}
	await Promise.all(promises)
}

/**
 * Steps 2 and 3: backfill inherited layout placeholders, then bake measured text fit.
 * @param {PresSlideInternal[]} slides - the slides to prepare
 * @param {FontMetricsRegistry} fontMetrics - registered font metrics; an empty registry skips the fit bake
 */
export function bakeSlideContent(slides: PresSlideInternal[], fontMetrics: FontMetricsRegistry): void {
	for (const slide of slides) {
		if (slide._slideLayout) addPlaceholdersToSlideLayouts(slide)
	}

	applyMeasuredFit(slides, fontMetrics)
}
