import PresentationCore from './pptxgen.js'
import { createBrowserRuntime } from './runtime/browser.js'
import { genTableToSlides } from './gen/table/html-dom.js'
import type { TableToSlidesProps } from './core-interfaces.js'

export class TsPptx extends PresentationCore {
	constructor() {
		super(createBrowserRuntime())
	}

	/**
	 * Reproduces a rendered HTML `<table>` as a PowerPoint table — including column widths,
	 * style, etc. — creating one or more slides as needed. Resolves `eleId` against the global
	 * `document`, so it exists only on the browser/standalone build; the same conversion is
	 * available anywhere there is a DOM as the free `tableToSlides` on `ts-pptx/html`, which
	 * also takes the element directly. The in-memory `slide.addTable(rows, opts)` path remains
	 * the platform-agnostic way to build a table from data you already hold.
	 * @param {string} eleId - table HTML element ID
	 * @param {TableToSlidesProps} options - generation options
	 */
	tableToSlides(eleId: string, options: TableToSlidesProps = {}): void {
		// @note `options.verbose` (a documented dev-only flag on TableToSlidesProps) is read
		// inside genTableToSlides to trace the auto-paging layout process.
		genTableToSlides(this, eleId, options)
	}
}

export { TsPptx as default }
export * from './core-enums.js'
export * from './units.js'
// Use `export *` (not `export type *`) so the value exports `textRun`/`textRuns`
// reach this entry; `export type *` would drop them and crash any consumer that
// imports them, while TypeScript (reading index.d.ts) stays green.
export * from './core-interfaces.js'
export { setDiagnosticHandler, type Diagnostic, type DiagnosticCode, type DiagnosticHandler } from './diagnostics.js'
