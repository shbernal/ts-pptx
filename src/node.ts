import PresentationCore from './pptxgen.js'
import { createNodeRuntime } from './runtime/node.js'

export class TsPptx extends PresentationCore {
	constructor() {
		super(createNodeRuntime())
	}
}

export { TsPptx as default }
export * from './core-enums.js'
export * from './units.js'
// Use `export *` (not `export type *`) so the value exports `textRun`/`textRuns`
// reach this entry; `export type *` would drop them and crash any Node consumer
// that imports them, while TypeScript (reading index.d.ts) stays green.
export * from './core-interfaces.js'
