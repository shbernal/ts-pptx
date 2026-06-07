import PresentationCore from './pptxgen.js'
import { createNodeRuntime } from './runtime/node.js'

export class PptxGenJS extends PresentationCore {
	constructor() {
		super(createNodeRuntime())
	}
}

export { PptxGenJS as Presentation, PptxGenJS as default }
export * from './core-enums.js'
export * from './units.js'
export type * from './core-interfaces.js'
export type { PresSlide as Slide } from './core-interfaces.js'
