import PresentationCore from './presentation.js'
import { createNodeRuntime } from './runtime/node.js'

/**
 * The Node entry, reached through the `node` export condition — a consumer importing the bare
 * `@shbernal/ts-pptx` specifier under Node lands here without naming this subpath.
 *
 * Same authoring API as every other entry (see `entry-surface.ts`); the runtime adapter is what
 * differs. This one can reach the filesystem, so `writeFile` writes a real file and media may be
 * loaded from a path — both of which throw on the runtime-agnostic entry.
 */
export class TsPptx extends PresentationCore {
	constructor() {
		super(createNodeRuntime())
	}
}

export { TsPptx as default }
export * from './entry-surface.js'
