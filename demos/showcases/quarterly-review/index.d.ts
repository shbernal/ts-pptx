/**
 * Types for the one showcase that is exported to other workspace packages.
 *
 * `demos/vite-demo` builds this deck in the browser, and its `build` script runs `tsc -b`, so
 * the import needs declarations. Only the quarterly review is exported: the Field Notes deck
 * loads photographs from disk by path and cannot run outside Node.
 *
 * Hand-written rather than generated — the deck is plain `.mjs` and the exported surface is
 * two values, so a build step to produce four lines would cost more than it saves.
 */

/** Build the deck and write it to `outFile`. Returns the written file name. */
export function build(outFile: string): Promise<string>;

export const showcase: {
	slug: string;
	title: string;
	description: string;
	fileName: string;
	build: typeof build;
};
