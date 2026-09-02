/**
 * ts-pptx: shared `Relationship` writers.
 *
 * A `.rels` part is mostly one-off entries, written where the part that needs
 * them is built. This module holds the ones **more than one part** writes, so
 * that the attribute set stays a single fact — a `Relationship` whose `Type` and
 * `TargetMode` disagree between two writers does not throw, it produces a link
 * PowerPoint silently drops.
 */

import { el, raw, voidEl, type XmlFmt } from '../oxml/el.js'
import { CRLF, XML_DECL } from '../../constants-internal.js'
import { HYPERLINK_REL, PACKAGE_REL_NS } from '../../ooxml/rel-types.js'

/**
 * One `Relationship` entry.
 *
 * Three private copies of this used to sit in `gen/opc/root-rels.ts`, `gen/pres/
 * presentation-rels.ts` and `gen/chart/embed-xlsx.ts`, identical but for whether the caller
 * passed `rId3` or `3`, with seven more written inline in `gen/slide/`. The attribute set and
 * its order are byte-significant, so one writer is the only way they stay one fact.
 *
 * `target` is passed **unescaped** — `voidEl` escapes it on the way out. `gen/slide/object.ts`
 * depends on that: it computes the escaped form separately to compare a rel against the bytes
 * already emitted, and the two escapers must not drift.
 * @param {string | number} id - the relationship id; a number gets the `rId` prefix
 * @param {string} type - the relationship type URI
 * @param {string} target - the target, unescaped (a partname, or a URL with `targetMode`)
 * @param {object} [opts] - `targetMode` for an external target; `fmt` for byte-layout prefixes
 * @returns {string} the `Relationship` element
 */
export function relationshipEl(
	id: string | number,
	type: string,
	target: string,
	// Read-only argument bag: both keys are consulted with `?.` below and nothing spreads it, so a
	// caller assembling one out of a value it may not have (`{ targetMode }`) is saying the same
	// thing as omitting the key, and the declaration says so.
	opts?: { targetMode?: string | undefined; fmt?: XmlFmt | undefined }
): string {
	return voidEl(
		'Relationship',
		{ Id: typeof id === 'number' ? `rId${id}` : id, Type: type, Target: target, TargetMode: opts?.targetMode },
		opts?.fmt
	)
}

/**
 * The `<Relationships>` wrapper around already-serialized entries.
 *
 * Only the tag name and its namespace are fixed here; `fmt` is passed straight through, because
 * the six `.rels` writers each indent differently and every one of those differences is in the
 * emitted bytes. Normalizing them would be a rewrite of five parts disguised as a cleanup.
 */
export function relationshipsEl(rels: readonly string[], fmt?: XmlFmt): string {
	return el('Relationships', { xmlns: PACKAGE_REL_NS }, rels.map(raw), fmt)
}

/**
 * A complete `.rels` part: the XML declaration, a CRLF, then {@link relationshipsEl}.
 *
 * The embedded workbook's parts are the exception and use {@link relationshipsEl} directly —
 * they follow the declaration with nothing, and two of them end with a bare newline.
 */
export function relationshipsPart(rels: readonly string[], fmt?: XmlFmt): string {
	return XML_DECL + CRLF + relationshipsEl(rels, fmt)
}

/**
 * A `Relationship` to a URL outside the package: the shape of every external
 * hyperlink, whether it hangs off a slide (`gen/slide/object.ts`) or off a notes
 * slide (`gen/slide/notes.ts`). `TargetMode="External"` is what makes the target
 * a URI rather than a partname, so it is not optional.
 * @param {number} rId - the relationship id, without its `rId` prefix
 * @param {string} target - the URL, unescaped; `voidEl` escapes it for the attribute
 * @returns {string} the `Relationship` element
 */
export function externalHyperlinkRel(rId: number, target: string): string {
	return relationshipEl(rId, HYPERLINK_REL, target, { targetMode: 'External' })
}
