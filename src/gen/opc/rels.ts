/**
 * ts-pptx: shared `Relationship` writers.
 *
 * A `.rels` part is mostly one-off entries, written where the part that needs
 * them is built. This module holds the ones **more than one part** writes, so
 * that the attribute set stays a single fact — a `Relationship` whose `Type` and
 * `TargetMode` disagree between two writers does not throw, it produces a link
 * PowerPoint silently drops.
 */

import { voidEl } from '../oxml/el.js'
import { OFFICE_REL } from '../../ooxml/rel-types.js'

/**
 * A `Relationship` to a URL outside the package: the shape of every external
 * hyperlink, whether it hangs off a slide (`gen/slide/object.ts`) or off a notes
 * slide (`gen/slide/notes.ts`). `TargetMode="External"` is what makes the target
 * a URI rather than a partname, so it is not optional.
 * @param {number} rId - the relationship id, without its `rId` prefix
 * @param {string} target - the URL, already escaped for an attribute value
 * @returns {string} the `Relationship` element
 */
export function externalHyperlinkRel(rId: number, target: string): string {
	return voidEl('Relationship', {
		Id: `rId${rId}`,
		Type: OFFICE_REL + 'hyperlink',
		Target: target,
		TargetMode: 'External',
	})
}
