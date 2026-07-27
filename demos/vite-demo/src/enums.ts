/**
 * The snippet shown on the page.
 *
 * Kept as a literal rather than read back from the deck module: what the page illustrates is
 * how *little* the browser has to do, and that is three lines. The deck itself is four hundred
 * and is not the point.
 */
export const DECK_SOURCE = `import { build, showcase } from 'ts-pptx-demos-showcases/quarterly-review'

// The same module Node builds. In a browser, writeFile() downloads
// the package instead of writing it, so a file name is all it needs.
await build(showcase.fileName)`;
