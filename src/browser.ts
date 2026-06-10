import PresentationCore from './pptxgen.js'
import { createBrowserRuntime } from './runtime/browser.js'

export class PptxGenJS extends PresentationCore {
	constructor() {
		super(createBrowserRuntime())
	}
}

export { PptxGenJS as Presentation, PptxGenJS as default }
export * from './core-enums.js'
export * from './units.js'
// Use `export *` (not `export type *`) so the value exports `textRun`/`textRuns`
// reach this entry; `export type *` would drop them and crash any consumer that
// imports them, while TypeScript (reading index.d.ts) stays green.
export * from './core-interfaces.js'
export type { PresSlide as Slide } from './core-interfaces.js'
