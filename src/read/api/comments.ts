/**
 * Read a slide's **legacy** review comments (`p:cmLst/p:cm` in
 * `ppt/comments/commentN.xml`) plus the deck-wide comment-author registry
 * (`p:cmAuthorLst/p:cmAuthor` in `ppt/commentAuthors.xml`), the read counterparts
 * of the write-side `slide.addComment(...)` / `commentAuthors.xml` emitter.
 *
 * This models the ISO/IEC 29500 §13 *legacy* comment schema — the one the writer
 * authors and older PowerPoint decks carry. The 2018 **modern** comment parts
 * (`p188:cm` / `ppt/comments/modernComment_*` + `authors.xml`) are a different
 * schema with no writer, so they round-trip byte-perfect but are not decoded here.
 */
import { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { attr, firstChild, getElements, intValue, type Element } from '../oxml/dom.js'

/** The slide → comments-part relationship type (legacy comments). */
export const COMMENTS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
/** The presentation → commentAuthors-part relationship type. */
export const COMMENT_AUTHORS_REL_TYPE =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors'

/** One entry of the deck-wide comment-author registry (`p:cmAuthor`). */
export interface CommentAuthor {
	/** Author id (`@id`), referenced by each comment's `@authorId`. */
	id: number | null
	/** Display name (`@name`). */
	name: string
	/** Initials shown in the comment marker (`@initials`). */
	initials: string
	/** Highest per-author comment index used (`@lastIdx`). */
	lastIdx: number | null
	/** Author colour slot (`@clrIdx`). */
	clrIdx: number | null
}

/** One legacy review comment (`p:cm`), with its author resolved via the registry. */
export interface Comment {
	/** Resolved author display name (via `@authorId` → `commentAuthors.xml`), or `null` when unresolved. */
	author: string | null
	/** Resolved author initials, or `null` when unresolved. */
	authorInitials: string | null
	/** The comment's `@authorId`, or `null` when absent/unparseable. */
	authorId: number | null
	/** The comment's per-author 1-based index (`@idx`), or `null` when absent. */
	idx: number | null
	/** The comment body text (`p:text`). */
	text: string
	/** Marker X position in EMU (`p:pos/@x`), or `null` when absent. */
	x: number | null
	/** Marker Y position in EMU (`p:pos/@y`), or `null` when absent. */
	y: number | null
	/** Authored timestamp (`@dt`, ISO-8601 as written), or `null` when absent. */
	date: string | null
}

/** Parse a `p:cmAuthorLst` root element into the author registry. */
function parseCommentAuthors(root: Element | null): CommentAuthor[] {
	if (!root) return []
	return getElements(root, 'p:cmAuthor').map((el) => ({
		id: intValue(attr(el, 'id')),
		name: attr(el, 'name') ?? '',
		initials: attr(el, 'initials') ?? '',
		lastIdx: intValue(attr(el, 'lastIdx')),
		clrIdx: intValue(attr(el, 'clrIdx')),
	}))
}

/**
 * Read the deck-wide comment authors from `ppt/commentAuthors.xml`, resolved via the
 * presentation part's `commentAuthors` relationship. `[]` when the deck has no comments.
 */
export function readCommentAuthors(opc: OpcPackage, presentationPartName: string): CommentAuthor[] {
	const rels = opc.relationshipsFor(presentationPartName)
	const rel = rels.byType(COMMENT_AUTHORS_REL_TYPE)[0]
	if (!rel) return []
	const part = opc.part(rels.resolveTarget(rel.id))
	return parseCommentAuthors(part?.dom.documentElement ?? null)
}

/**
 * Read one slide's legacy comments from its `comments/commentN.xml` part (resolved
 * via the slide's `comments` relationship), resolving each comment's `@authorId`
 * against `authors` (typically {@link readCommentAuthors}). `[]` when the slide has
 * no comments part.
 */
export function readSlideComments(opc: OpcPackage, slidePart: Part, authors: CommentAuthor[]): Comment[] {
	const rels = opc.relationshipsFor(slidePart.partName)
	const rel = rels.byType(COMMENTS_REL_TYPE)[0]
	if (!rel) return []
	const part = opc.part(rels.resolveTarget(rel.id))
	const root = part?.dom.documentElement
	if (!root) return []

	const byId = new Map(authors.map((a) => [a.id, a]))
	return getElements(root, 'p:cm').map((cm) => {
		const authorId = intValue(attr(cm, 'authorId'))
		const author = authorId !== null ? (byId.get(authorId) ?? null) : null
		const pos = firstChild(cm, 'p:pos')
		const textEl = firstChild(cm, 'p:text')
		return {
			author: author ? author.name : null,
			authorInitials: author ? author.initials : null,
			authorId,
			idx: intValue(attr(cm, 'idx')),
			text: textEl?.textContent ?? '',
			x: pos ? intValue(attr(pos, 'x')) : null,
			y: pos ? intValue(attr(pos, 'y')) : null,
			date: attr(cm, 'dt') ?? null,
		}
	})
}
