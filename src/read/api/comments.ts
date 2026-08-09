/**
 * Read a slide's **legacy** review comments (`p:cmLst/p:cm` in
 * `ppt/comments/commentN.xml`) plus the deck-wide comment-author registry
 * (`p:cmAuthorLst/p:cmAuthor` in `ppt/commentAuthors.xml`), the read counterparts
 * of the write-side `slide.addComment(...)` / `commentAuthors.xml` emitter.
 *
 * This models the ISO/IEC 29500 §13 *legacy* comment schema — the one the writer
 * authors and older PowerPoint decks carry. The 2018 **modern** comment parts
 * (`p188:cm` / `ppt/comments/modernComment_*` + `authors.xml`) are a separate,
 * GUID-keyed, thread-capable schema; they are decoded read-only by the
 * `readModern*` functions below, and have no writer (preserved byte-perfect on
 * round-trip). The two schemas do not coexist in a single deck in practice —
 * {@link commentSchema} reports which one a deck uses.
 */
import { OpcPackage } from '../opc/package.js'
import type { Part } from '../opc/part.js'
import { attr, firstChild, getElements, intValue, type Element } from '../oxml/dom.js'

/** The slide → comments-part relationship type (legacy comments). */
const COMMENTS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments'
/** The presentation → commentAuthors-part relationship type. */
const COMMENT_AUTHORS_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors'

/** Content type of a legacy per-slide comments part (`ppt/comments/commentN.xml`). */
const LEGACY_COMMENTS_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml'
/** The slide → modern-comments-part relationship type (2018 schema). */
const MODERN_COMMENTS_REL_TYPE = 'http://schemas.microsoft.com/office/2018/10/relationships/comments'
/** The presentation → modern-authors-part relationship type (2018 schema). */
const MODERN_AUTHORS_REL_TYPE = 'http://schemas.microsoft.com/office/2018/10/relationships/authors'
/** Content type of the modern per-slide comments part (`ppt/comments/modernComment_*.xml`). */
const MODERN_COMMENTS_CONTENT_TYPE = 'application/vnd.ms-powerpoint.comments+xml'

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

// --- Modern (2018) comments ------------------------------------------------
//
// PowerPoint's post-2018 comment schema (`p188`, uri
// `http://schemas.microsoft.com/office/powerpoint/2018/8/main`). Authors move to
// their own deck-wide `ppt/authors.xml` keyed by **GUID**, and each slide's
// comments live in a `ppt/comments/modernComment_*.xml` part that supports reply
// threads. This is read-only; the writer never emits these, and a modern part is
// preserved byte-for-byte on round-trip.

/** One entry of the modern deck-wide author registry (`p188:author`). */
export interface ModernCommentAuthor {
	/** Author id (`@id`) — a GUID string like `{E8A64ABA-…}`, referenced by each comment's `@authorId`. */
	id: string | null
	/** Display name (`@name`). */
	name: string
	/** Initials (`@initials`). */
	initials: string
	/** Account identifier (`@userId`, e.g. an email), or `null`. */
	userId: string | null
	/** Identity provider (`@providerId`, e.g. `Windows Live`), or `null`. */
	providerId: string | null
}

/** One modern comment or reply (`p188:cm` / `p188:reply`), with its author resolved via the registry. */
export interface ModernComment {
	/** The comment's own GUID id (`@id`), or `null` when absent. */
	id: string | null
	/** Resolved author display name (via `@authorId` → `authors.xml`), or `null` when unresolved. */
	author: string | null
	/** Resolved author initials, or `null` when unresolved. */
	authorInitials: string | null
	/** The comment's `@authorId` GUID, or `null` when absent. */
	authorId: string | null
	/** Creation timestamp (`@created`, ISO-8601 as written), or `null` when absent. */
	created: string | null
	/** The comment body text; multiple paragraphs are joined with `\n`. */
	text: string
	/** Marker X position in EMU (`p188:pos/@x`), or `null` (replies carry no position). */
	x: number | null
	/** Marker Y position in EMU (`p188:pos/@y`), or `null` (replies carry no position). */
	y: number | null
	/** Nested replies in thread order; each is itself a {@link ModernComment} with no further replies. */
	replies: ModernComment[]
}

/** Which comment schema a deck carries. `'none'` when it has neither part. */
export type CommentSchema = 'legacy' | 'modern' | 'none'

/** Concatenate a `p188:txBody`'s paragraphs into plain text, joining paragraphs with `\n`. */
function modernCommentText(el: Element): string {
	const txBody = firstChild(el, 'p188:txBody')
	if (!txBody) return ''
	const paras = getElements(txBody, 'a:p')
	if (paras.length === 0) return txBody.textContent ?? ''
	return paras.map((p) => p.textContent ?? '').join('\n')
}

/**
 * Decode one `p188:cm` (or `p188:reply`) element. Only the top-level comment
 * carries a position and a reply list; replies pass `withThread: false` so they
 * report `x`/`y` as `null` and never recurse (the schema forbids nested replies).
 */
function parseModernComment(
	el: Element,
	byId: Map<string | null, ModernCommentAuthor>,
	withThread: boolean
): ModernComment {
	const authorId = attr(el, 'authorId')
	const author = authorId !== null ? (byId.get(authorId) ?? null) : null
	const pos = withThread ? firstChild(el, 'p188:pos') : null
	const replyLst = withThread ? firstChild(el, 'p188:replyLst') : null
	return {
		id: attr(el, 'id'),
		author: author ? author.name : null,
		authorInitials: author ? author.initials : null,
		authorId,
		created: attr(el, 'created'),
		text: modernCommentText(el),
		x: pos ? intValue(attr(pos, 'x')) : null,
		y: pos ? intValue(attr(pos, 'y')) : null,
		replies: replyLst ? getElements(replyLst, 'p188:reply').map((r) => parseModernComment(r, byId, false)) : [],
	}
}

/**
 * Read the deck-wide modern comment authors from `ppt/authors.xml`, resolved via
 * the presentation part's 2018 `authors` relationship. `[]` when the deck has no
 * modern authors part.
 */
export function readModernCommentAuthors(opc: OpcPackage, presentationPartName: string): ModernCommentAuthor[] {
	const rels = opc.relationshipsFor(presentationPartName)
	const rel = rels.byType(MODERN_AUTHORS_REL_TYPE)[0]
	if (!rel) return []
	const part = opc.part(rels.resolveTarget(rel.id))
	const root = part?.dom.documentElement
	if (!root) return []
	return getElements(root, 'p188:author').map((el) => ({
		id: attr(el, 'id'),
		name: attr(el, 'name') ?? '',
		initials: attr(el, 'initials') ?? '',
		userId: attr(el, 'userId'),
		providerId: attr(el, 'providerId'),
	}))
}

/**
 * Read one slide's modern comments from its `modernComment_*.xml` part (resolved
 * via the slide's 2018 `comments` relationship), resolving each comment's and
 * reply's `@authorId` GUID against `authors` (typically
 * {@link readModernCommentAuthors}). Replies are nested, not flattened. `[]` when
 * the slide has no modern comments part.
 */
export function readModernSlideComments(
	opc: OpcPackage,
	slidePart: Part,
	authors: ModernCommentAuthor[]
): ModernComment[] {
	const rels = opc.relationshipsFor(slidePart.partName)
	const rel = rels.byType(MODERN_COMMENTS_REL_TYPE)[0]
	if (!rel) return []
	const root = opc.part(rels.resolveTarget(rel.id))?.dom.documentElement
	if (!root) return []
	const byId = new Map(authors.map((a) => [a.id, a]))
	return getElements(root, 'p188:cm').map((cm) => parseModernComment(cm, byId, true))
}

/**
 * Report which comment schema a deck uses: `'modern'` when it carries any 2018
 * `modernComment_*` part, else `'legacy'` when it carries any classic
 * `commentN.xml` part, else `'none'`. Lets a consumer pick the right accessor
 * ({@link readModernSlideComments} vs {@link readSlideComments}) without probing
 * both.
 */
export function commentSchema(opc: OpcPackage): CommentSchema {
	if (opc.partsByContentType(MODERN_COMMENTS_CONTENT_TYPE).length > 0) return 'modern'
	if (opc.partsByContentType(LEGACY_COMMENTS_CONTENT_TYPE).length > 0) return 'legacy'
	return 'none'
}
