import PresentationCore from './presentation.js'
import { createNeutralRuntime } from './runtime/neutral.js'

/**
 * The runtime-agnostic entry: what a consumer gets from the bare `@shbernal/ts-pptx` specifier
 * when neither the `node` nor the `browser` export condition resolves — Deno, Bun, edge workers.
 * Node and browser consumers reach `ts-pptx/node` and `ts-pptx/browser` through those conditions
 * without naming them, and get the same class backed by an adapter that can reach their host.
 *
 * Authoring is identical on all three. The difference is only where the finished deck can go:
 * `write`, `stream` and `toParts` hand the bytes back here as everywhere, while `writeFile`
 * throws `runtime/file-output-unavailable` because there is no filesystem and no DOM to write
 * to. Live-DOM `tableToSlides` is likewise absent — it is defined on the browser entry, and the
 * DOM-agnostic form is the free `tableToSlides` on `ts-pptx/html`.
 */
export class TsPptx extends PresentationCore {
	constructor() {
		super(createNeutralRuntime())
	}
}

export { TsPptx as default }
export * from './entry-surface.js'
