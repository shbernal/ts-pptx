import PresentationCore from './presentation.js'
import { composePresentation, type ComposeOptions, type Composed } from './entry-compose.js'
import type { ConstructFamily } from './families/shared.js'
import { ALL_CONSTRUCT_FAMILIES } from './entry-families.js'
import { createBrowserRuntime } from './runtime/browser.js'
import { tableDomFamily } from './families/table-dom.js'
import type { TableToSlidesProps } from './types/index.js'

/**
 * The browser entry, reached through the `browser` export condition — a bundler resolving the
 * bare `pptx-ts` specifier for the web lands here without naming this subpath.
 *
 * Same authoring API as every other entry (see `entry-surface.ts`), plus one method that only
 * makes sense with a live DOM: {@link TsPptx.tableToSlides}. `writeFile` triggers a download
 * rather than touching a filesystem.
 */
export class TsPptx extends PresentationCore {
	constructor() {
		// The table family's live-DOM half is composed here and nowhere else: it is the one part of
		// the authoring surface that needs a `document`.
		super(createBrowserRuntime(), [...ALL_CONSTRUCT_FAMILIES, tableDomFamily])
	}

	/**
	 * Reproduces a rendered HTML `<table>` as a PowerPoint table — including column widths,
	 * style, etc. — creating one or more slides as needed. Resolves `eleId` against the global
	 * `document`, so it exists only on the browser/standalone build; the same conversion is
	 * available anywhere there is a DOM as the free `tableToSlides` on `ts-pptx/html`, which
	 * also takes the element directly. The in-memory `slide.addTable(rows, opts)` path remains
	 * the platform-agnostic way to build a table from data you already hold.
	 *
	 * That global lookup is why this stays rather than folding into the free function: reaching
	 * `document` without being handed one is the affordance the browser entry exists to give, and
	 * it is the one thing the platform-agnostic form deliberately will not do. The conversion
	 * itself lives in one place either way — both spellings call `genTableToSlides`.
	 * @param {string} eleId - table HTML element ID
	 * @param {TableToSlidesProps} options - generation options
	 */
	tableToSlides(eleId: string, options: TableToSlidesProps = {}): void {
		this.familyAuthor('tableToSlides')(eleId, options)
	}
}

/**
 * Compose a presentation from the construct families it needs, on the browser adapter.
 *
 * The counterpart to {@link TsPptx}, which is composed with every family there is. This one costs
 * what it carries: the core tier (text, shapes, images, groups, speaker notes) plus whatever
 * `use` names, and nothing else reaches the bundle. `tableToSlides` is one of those families
 * (`domTables`), so a deck that wants it composes it.
 * @param {ComposeOptions} opts - the families to compose in, from `pptx-ts/families`
 * @example
 * import { createPresentation } from 'pptx-ts/browser'
 * import { charts, domTables } from 'pptx-ts/families'
 * const pres = createPresentation({ use: [charts, domTables] })
 */
export function createPresentation<const Fs extends readonly ConstructFamily[] = []>(
	opts?: ComposeOptions<Fs>
): Composed<Fs> {
	return composePresentation(createBrowserRuntime(), opts)
}

export { TsPptx as default }
export * from './entry-surface.js'
