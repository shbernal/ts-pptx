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
 * Three steps rather than one combined pass, because the package path has work of its own that
 * must land between the last two: it de-duplicates identical media across the deck (which needs
 * the encoded bytes {@link encodeMediaForTargets} produces) and assigns package-unique chart part
 * filenames, both before any text is measured. Exposing the steps lets `assemble` keep that work
 * in place while still sharing the code.
 *
 * The order is load-bearing:
 *
 * 1. **Backfill placeholders first**, so a slide carries its layout's placeholder objects before
 *    anything reads its media. A seeded placeholder can register media of its own, an image fill,
 *    and when backfill ran after encoding that image was never loaded and its rel named a part
 *    that was never written.
 * 2. **Media.** It is the only asynchronous step, and the sync XML pass reads the `rel.data` it
 *    populates.
 * 3. **Bake measured fit last**, because it measures the text step 1 may have just added and
 *    writes the `fontScale` the sync XML pass then reads.
 *
 * The functions take different target sets on purpose: media lives on layouts and the master as
 * well as on slides, while placeholder backfill and measured fit apply only to slides.
 * `extractSlides` passes only its slides to all three, because it emits no layout or master
 * parts.
 */

import { InvalidOptionError } from '../errors.js'
import { applyMeasuredFit } from '../measure/fit.js'
import type { FontMetricsRegistry } from '../measure/font-metrics.js'
import type { RuntimeAdapter } from '../runtime/types.js'
import type { PresSlideInternal, SlideLayoutInternal, SlideMasterInternal } from '../types/internal.js'
import { addPlaceholdersToSlideLayouts } from './define/placeholder.js'
import { encodeSlideMediaRels } from './media.js'
import { isHyperlinkRel } from './utils.js'

/**
 * Before any step: refuse a link to a slide the deck does not have.
 *
 * A `hyperlink: { slide }` and a Slide Zoom's numeric `target` are checked for being slide numbers
 * when they are authored, but whether the slide exists is only known once the deck is complete: a
 * link to slide 3 can be authored before slide 3 is added. Both serializing callers check here,
 * before anything is built. Otherwise the link is written as a relationship to a part the package
 * never has.
 * @param slides - every slide in the deck, in order
 */
export function requireSlideLinksInDeck(slides: readonly PresSlideInternal[]): void {
	for (const slide of slides) {
		for (const rel of slide._rels) {
			if (!isHyperlinkRel(rel) || rel.data !== 'slide') continue
			// Negated so a target that is not a number at all passes on to the internal assertion
			// `extractSlides` makes, rather than being blamed on the caller here.
			if (!(Number(rel.Target) > slides.length)) continue
			throw new InvalidOptionError(
				'slide/link-past-last-slide',
				`Slide ${slide._slideNum} links to slide ${String(rel.Target)}, but the deck has ${slides.length} slide${slides.length === 1 ? '' : 's'}.`
			)
		}
	}
}

/**
 * Step 2: load and base64-encode every media rel on the given targets, populating `rel.data`.
 * @param {ReadonlyArray<PresSlideInternal | SlideLayoutInternal | SlideMasterInternal>} targets - media-bearing targets (slides, layouts, master)
 * @param {RuntimeAdapter} runtime - runtime adapter used to load media
 * @param {'throw' | 'placeholder'} onMediaError - failure policy for unloadable media
 */
export async function encodeMediaForTargets(
	targets: ReadonlyArray<PresSlideInternal | SlideLayoutInternal | SlideMasterInternal>,
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
 * Step 1: seed each slide with the layout placeholders it leaves empty.
 * @param {PresSlideInternal[]} slides - the slides to prepare
 */
export function backfillPlaceholders(slides: PresSlideInternal[]): void {
	for (const slide of slides) {
		if (slide._slideLayout) addPlaceholdersToSlideLayouts(slide)
	}
}

/**
 * Step 3: bake measured text fit.
 * @param {PresSlideInternal[]} slides - the slides to prepare
 * @param {FontMetricsRegistry} fontMetrics - registered font metrics; an empty registry skips the fit bake
 */
export function bakeMeasuredFit(slides: PresSlideInternal[], fontMetrics: FontMetricsRegistry): void {
	applyMeasuredFit(slides, fontMetrics)
}
