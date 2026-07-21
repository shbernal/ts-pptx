import PresentationCore from './pptxgen.js'
import { createBrowserRuntime } from './runtime/browser.js'
import { genTableToSlides } from './gen/table/html-dom.js'
import type { TableToSlidesProps } from './core-interfaces.js'

export class PptxGenJS extends PresentationCore {
	constructor() {
		super(createBrowserRuntime())
	}

	/**
	 * Reproduces a rendered HTML `<table>` as a PowerPoint table — including column widths,
	 * style, etc. — creating one or more slides as needed. Reads the live DOM
	 * (`getComputedStyle`, `offsetWidth`), so it exists only on the browser/standalone build,
	 * not the Node build. The in-memory `slide.addTable(rows, opts)` path is the supported,
	 * platform-agnostic way to build tables.
	 * @param {string} eleId - table HTML element ID
	 * @param {TableToSlidesProps} options - generation options
	 */
	tableToSlides(eleId: string, options: TableToSlidesProps = {}): void {
		// @note `verbose` option is undocumented; used for verbose output of layout process
		genTableToSlides(
			this,
			eleId,
			options,
			options?.masterSlideName
				? this._slideLayouts.find((layout) => layout._name === options.masterSlideName)
				: undefined
		)
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
