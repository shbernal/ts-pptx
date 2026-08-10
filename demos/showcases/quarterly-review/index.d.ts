/**
 * Types for the one showcase that is exported to other workspace packages.
 *
 * The site's demos page (`www/demos/`) builds this deck in the browser, and the site is
 * typechecked, so the import needs declarations. Only the quarterly review is exported:
 * the Field Notes deck loads photographs from disk by path and cannot run outside Node.
 *
 * Hand-written rather than generated — the deck is plain `.mjs` and the exported surface is
 * three values, so a build step to produce a handful of lines would cost more than it saves.
 */
import type TsPptx from '@shbernal/ts-pptx'

/**
 * Assemble the deck and return the presentation, having written nothing.
 *
 * What a caller that needs the *bytes* uses — the preview feeds them to a reader. `build`
 * is the same deck with a destination attached.
 */
export function compose(): Promise<TsPptx>

/** Build the deck and write it to `outFile`. Returns the written file name. */
export function build(outFile: string): Promise<string>

export const showcase: {
	slug: string
	title: string
	description: string
	fileName: string
	build: typeof build
}
