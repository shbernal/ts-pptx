import PresentationCore from './presentation.js'
import { createBrowserRuntime } from './runtime/browser.js'
import { genTableToSlides } from './gen/table/html-dom.js'
import type { TableToSlidesProps } from './types/index.js'

/**
 * The browser entry, reached through the `browser` export condition — a bundler resolving the
 * bare `@shbernal/ts-pptx` specifier for the web lands here without naming this subpath.
 *
 * Same authoring API as every other entry (see `entry-surface.ts`), plus one method that only
 * makes sense with a live DOM: {@link TsPptx.tableToSlides}. `writeFile` triggers a download
 * rather than touching a filesystem.
 */
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
export * from './entry-surface.js'
