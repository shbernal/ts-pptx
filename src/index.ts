/**
 * The runtime-agnostic authoring entry, which the bare `pptx-ts` specifier resolves to when
 * neither the `node` nor the `browser` export condition applies.
 *
 * @module
 */
import PresentationCore from './presentation.js'
import { composePresentation, type ComposeOptions, type Composed } from './entry-compose.js'
import type { ConstructFamily } from './families/shared.js'
import { ALL_CONSTRUCT_FAMILIES } from './entry-families.js'
import { createNeutralRuntime } from './runtime/neutral.js'

/**
 * The runtime-agnostic entry: what a consumer gets from the bare `pptx-ts` specifier
 * when neither the `node` nor the `browser` export condition resolves — Deno, Bun, edge workers.
 * Node and browser consumers reach `pptx-ts/node` and `pptx-ts/browser` through those conditions
 * without naming them, and get the same class backed by an adapter that can reach their host.
 *
 * Authoring is identical on all three. The difference is only where the finished deck can go:
 * `write`, `stream` and `toParts` hand the bytes back here as everywhere, while `writeFile`
 * throws `runtime/file-output-unavailable` because there is no filesystem and no DOM to write
 * to. Live-DOM `tableToSlides` is likewise absent — it is defined on the browser entry, and the
 * DOM-agnostic form is the free `tableToSlides` on `pptx-ts/html`.
 */
export class TsPptx extends PresentationCore {
	constructor() {
		super(createNeutralRuntime(), ALL_CONSTRUCT_FAMILIES)
	}
}

/**
 * Compose a presentation from the construct families it needs, on the runtime-agnostic adapter.
 *
 * The counterpart to {@link TsPptx}, which is composed with every family there is. This one costs
 * what it carries: the core tier (text, shapes, images, groups, speaker notes) plus whatever
 * `use` names, and nothing else reaches the bundle.
 * @param {ComposeOptions} opts - the families to compose in, from `pptx-ts/families`
 * @example
 * import { createPresentation } from 'pptx-ts'
 * import { charts } from 'pptx-ts/families'
 * const pres = createPresentation({ use: [charts] })
 */
export function createPresentation<const Fs extends readonly ConstructFamily[] = []>(
	opts?: ComposeOptions<Fs>
): Composed<Fs> {
	return composePresentation(createNeutralRuntime(), opts)
}

export { TsPptx as default }
export * from './entry-surface.js'
