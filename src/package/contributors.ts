/**
 * ts-pptx: every construct family's part contributor, in one table
 *
 * The whole mapping from a construct family to the parts it puts in the package. It lives apart
 * from the packager that reads it (`package/assemble.ts`) for one reason: a named import inside a
 * reachable function body is retained unconditionally, so a packager that names the chart, comment
 * and notes emitters itself links all three into every program that writes a deck — a deck of text
 * boxes would carry the chart part builders. Naming them here instead makes "which families does
 * this program package" a question about which list is handed to the write.
 *
 * The list is unordered as far as callers are concerned: the packager sorts it by each
 * contributor's declared `order`, so a tier that ships a subset writes the same bytes for the
 * parts it does ship. Nothing else belongs here — this module stays a list of names.
 */

import { notesContributor } from './parts/notes.js'
import type { PartContributor } from './parts/shared.js'

/** Every construct family that adds parts to a package — the list the full authoring surface writes with. */
export const ALL_PART_CONTRIBUTORS: readonly PartContributor[] = Object.freeze([notesContributor])
