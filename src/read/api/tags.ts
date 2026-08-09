/**
 * Read the programmatic **tags** attached to an owner part — a slide or the
 * presentation itself. A tag is a simple `@name`/`@val` string pair PowerPoint
 * (or an add-in) stores out-of-band from the visible content, referenced from the
 * owner via `p:custDataLst/p:tags@r:id` and materialized in a `ppt/tags/tagN.xml`
 * part (`p:tagLst` → `p:tag`, CT_StringTag).
 *
 * There is no writer for tags, so this is a read-only accessor; a deck's tag parts
 * are preserved byte-for-byte on round-trip. An owner may reference more than one
 * tag part (rare but legal); {@link readTagsForPart} concatenates them in
 * relationship order.
 */
import { OpcPackage } from '../opc/package.js'
import { attr, getElements } from '../oxml/dom.js'

/** The owner → tags-part relationship type. */
const TAGS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags'

/** One programmatic tag (`p:tag`) — a name/value string pair. */
export interface Tag {
	/** The tag's `@name`. */
	name: string
	/** The tag's `@val`. */
	val: string
}

/**
 * Read every tag referenced by `ownerPartName` (a slide part or the presentation
 * part), resolving the owner's `tags` relationship(s) to their `ppt/tags/tagN.xml`
 * part(s) and flattening each part's `p:tag` children into `{ name, val }`. `[]`
 * when the owner references no tag part. Rel-driven rather than walking
 * `p:custDataLst`, so it is robust to element ordering in the owner.
 */
export function readTagsForPart(opc: OpcPackage, ownerPartName: string): Tag[] {
	const rels = opc.relationshipsFor(ownerPartName)
	const out: Tag[] = []
	for (const rel of rels.byType(TAGS_REL_TYPE)) {
		const root = opc.part(rels.resolveTarget(rel.id))?.dom.documentElement
		if (!root) continue
		for (const tag of getElements(root, 'p:tag')) {
			out.push({ name: attr(tag, 'name') ?? '', val: attr(tag, 'val') ?? '' })
		}
	}
	return out
}
