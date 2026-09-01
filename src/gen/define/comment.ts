/**
 * ts-pptx: Comment Definition
 *
 * Validates `addComment()` author / text / position and stashes a review comment on the slide
 * (`_comments`); the `<p:cm>` XML is emitted later by `gen/slide/comments.ts`.
 */
import { warn } from '../../diagnostics.js'
import type { CommentProps } from '../../types/index.js'
import type { PresSlideInternal } from '../../types/internal.js'

/**
 * Derive 1-2 letter initials from an author display name (e.g. "Ada Lovelace" -> "AL").
 * Falls back to the first character when the name is a single word.
 */
function deriveAuthorInitials(author: string): string {
	const words = author.trim().split(/\s+/).filter(Boolean)
	const first = words[0]
	if (!first) return '?'
	if (words.length === 1) return first.charAt(0).toUpperCase()
	const last = words[words.length - 1] ?? first
	return (first.charAt(0) + last.charAt(0)).toUpperCase()
}

/**
 * Adds a review comment to a slide (legacy ISO/IEC 29500 §13 comment).
 * @param {PresSlideInternal} target slide object the comment is attached to
 * @param {CommentProps} opts comment author/text/position options
 */
export function addCommentDefinition(target: PresSlideInternal, opts: CommentProps): void {
	const author = typeof opts?.author === 'string' ? opts.author.trim() : ''
	const text = typeof opts?.text === 'string' ? opts.text : ''
	// Don't silently coerce: a comment with no author or no body is meaningless, so warn + skip
	// rather than emit a degenerate <p:cm> (API policy: warn over silent coercion).
	if (!author) {
		warn('comment/missing-author', 'addComment() requires a non-empty `author`; comment ignored.')
		return
	}
	if (!text) {
		warn('comment/missing-text', 'addComment() requires non-empty `text`; comment ignored.')
		return
	}

	const initials =
		typeof opts.initials === 'string' && opts.initials.trim() ? opts.initials.trim() : deriveAuthorInitials(author)
	const x = typeof opts.x === 'number' && Number.isFinite(opts.x) ? opts.x : 0.5
	const y = typeof opts.y === 'number' && Number.isFinite(opts.y) ? opts.y : 0.5
	let date: string | undefined
	if (opts.date instanceof Date) date = opts.date.toISOString()
	else if (typeof opts.date === 'string' && opts.date) date = opts.date

	if (!target._comments) target._comments = []
	// `date` is omitted rather than written as `undefined` when the caller stated none: the
	// comments emitter falls back to its own timestamp on an absent one, and the two spellings
	// would otherwise be two ways of asking for that.
	target._comments.push(date === undefined ? { author, initials, text, x, y } : { author, initials, text, x, y, date })
}
