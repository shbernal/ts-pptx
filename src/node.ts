import PresentationCore from './presentation.js'
import { composePresentation, type ComposeOptions, type Composed } from './entry-compose.js'
import type { ConstructFamily } from './families/shared.js'
import { ALL_CONSTRUCT_FAMILIES } from './entry-families.js'
import { createNodeRuntime } from './runtime/node.js'

/**
 * The Node entry, reached through the `node` export condition — a consumer importing the bare
 * `pptx-ts` specifier under Node lands here without naming this subpath.
 *
 * Same authoring API as every other entry (see `entry-surface.ts`); the runtime adapter is what
 * differs. This one can reach the filesystem, so `writeFile` writes a real file and media may be
 * loaded from a path — both of which throw on the runtime-agnostic entry.
 */
export class TsPptx extends PresentationCore {
	constructor() {
		super(createNodeRuntime(), ALL_CONSTRUCT_FAMILIES)
	}
}

/**
 * Compose a presentation from the construct families it needs, on the Node adapter.
 *
 * The counterpart to {@link TsPptx}, which is composed with every family there is. This one costs
 * what it carries: the core tier (text, shapes, images, groups, speaker notes) plus whatever
 * `use` names, and nothing else reaches the bundle.
 * @param {ComposeOptions} opts - the families to compose in, from `pptx-ts/families`
 * @example
 * import { createPresentation } from 'pptx-ts/node'
 * import { charts } from 'pptx-ts/families'
 * const pres = createPresentation({ use: [charts] })
 */
export function createPresentation<const Fs extends readonly ConstructFamily[] = []>(
	opts?: ComposeOptions<Fs>
): Composed<Fs> {
	return composePresentation(createNodeRuntime(), opts)
}

export { TsPptx as default }
export * from './entry-surface.js'
