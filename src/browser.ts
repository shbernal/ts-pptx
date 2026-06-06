import PresentationCore from './pptxgen.js'
import { createBrowserRuntime } from './runtime/browser.js'

export class PptxGenJS extends PresentationCore {
	constructor() {
		super(createBrowserRuntime())
	}
}

export { PptxGenJS as Presentation, PptxGenJS as default }
export * from './core-enums.js'
export type * from './core-interfaces.js'
export type { PresSlide as Slide } from './core-interfaces.js'
