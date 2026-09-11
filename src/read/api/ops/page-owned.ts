/**
 * What a page copy may share with the page it was copied from, and what it must own.
 *
 * Duplicating a page — `cloneSlide`, `importSlide` of the same page twice,
 * `importSlides` naming it twice — copies the page part and then has to decide,
 * for every part hanging off it, whether the copy points at the same part or gets
 * one of its own. Sharing is the cheap answer and is right for most of the graph:
 * a layout, a master, a theme, an image are deck-wide assets PowerPoint itself
 * stores once and points many shapes at.
 *
 * It is not right for the parts a page *owns*. A chart and its embedded workbook,
 * a SmartArt diagram's five parts, an OLE embedding, a notes slide: PowerPoint
 * treats each as belonging to exactly one page, and a package where two slides
 * resolve to one of them is one it refuses to open at all — the "the file or
 * directory is corrupted and unreadable" (`0x80070570`) dialog, with the whole
 * deck rejected rather than the duplicate page. The schema validator accepts such
 * a package: nothing in ECMA-376 says a chart part has one referrer, so this rule
 * is only visible against the application.
 *
 * Measured against desktop PowerPoint on the `mixed` fixture (a chart page and a
 * SmartArt page), duplicating a page that owns one of these:
 *   - sharing the chart part → refused;
 *   - a copy of the chart part sharing its embedded workbook and its
 *     `chartUserShapes` drawing → still refused;
 *   - the chart subtree copied whole → opens.
 * So ownership is transitive: a part a page owns owns its own subtree in turn,
 * down to the media blobs at the leaves.
 *
 * Hence the list below is of what may be **shared**, and everything else is
 * copied. The asymmetry is deliberate: a wrongly shared part is a deck nobody can
 * open, a wrongly copied one is some duplicated bytes. A new relationship type
 * that nobody has classified therefore lands on the safe side by default.
 */

import { SHARED_PARTS } from '../../../ooxml/rel-types.js'

/**
 * Relationship types whose target two copies of one page may point at together: every kind in
 * {@link SHARED_PARTS}. That is deck furniture a page reaches and the deck owns, including the
 * author registry a comments part names; media blobs, which PowerPoint itself stores once and
 * points every shape at; and another page or an external link, where a jump points at whatever
 * copy of the target the import decided on.
 */
const SHARED_BY_PAGE_COPIES: ReadonlySet<string> = new Set(SHARED_PARTS.map((kind) => kind.relType))

/**
 * Whether a page copy may point at this relationship's target rather than taking
 * a copy of it. Everything not listed is owned by the page and copied — see the
 * module comment for why that is the safe default.
 */
export function isSharedByPageCopies(relType: string): boolean {
	return SHARED_BY_PAGE_COPIES.has(relType)
}
