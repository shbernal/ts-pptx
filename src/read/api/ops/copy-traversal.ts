/**
 * ts-pptx: the one rule three copy traversals follow.
 *
 * Copying a page into another deck means walking its part graph, and two places walk it:
 * `rebuildRels`, the loop under every copy (and, run as a `CopyPlan`, under the dry run that has
 * to reach exactly the same parts before anything moves); and `identicalSubgraph`, which asks
 * whether the destination already holds what the copy would have produced. Each used to write
 * the rule out, and a separate dry run did too, with a comment asking the reader to keep the
 * copies in step by hand — it said the guarantee was "only as good as the drift between them".
 *
 * A leaf module rather than an export from `part-copy.ts`, because `part-copy.ts` already
 * imports from `part-reuse.ts` and the rule has to be reachable from both.
 */

import type { Part } from '../../opc/part.js'
import type { Relationship } from '../../opc/relationships.js'
import { NOTES_SLIDE_REL, SLIDE_LAYOUT_REL, SLIDE_MASTER_CONTENT_TYPE } from '../../../ooxml/rel-types.js'

/**
 * What a copy traversal does with one relationship: skip it, carry it as an external link, or
 * follow it into the part it names.
 *
 * Notes pull in a notesMaster and its own theme, which an imported page does not need, so the
 * notes relationship is never followed — `carryNotes` picks it up afterwards for the pages
 * that asked. A master's relationships to its layouts are skipped so the copy stays lean:
 * copied layouts re-link themselves, and skipping the list is also what keeps one edited
 * layout from disqualifying every other layout that shares its master.
 *
 * Each caller still owns its own extra arms — a copy's overrides and its jump-link rule, the reuse
 * check's byte comparison — which is why this answers only the part they share.
 * @param part - the part the relationship belongs to
 * @param rel - the relationship being considered
 */
export function copyTraversalStep(part: Part, rel: Relationship): 'skip' | 'external' | 'recurse' {
	if (rel.type === NOTES_SLIDE_REL) return 'skip'
	if (part.contentType === SLIDE_MASTER_CONTENT_TYPE && rel.type === SLIDE_LAYOUT_REL) return 'skip'
	return rel.targetMode === 'External' ? 'external' : 'recurse'
}
