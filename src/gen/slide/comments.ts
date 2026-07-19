/**
 * PptxGenJS: comment parts
 *
 * Resolve every slide's comments into a deck-wide author registry with per-author
 * numbering, then emit the presentation-level `commentAuthors.xml` and each
 * slide's `comments/commentN.xml` part.
 */

import { CRLF, XML_DECL } from '../../core-enums.js'
import type { PresSlideInternal, ResolvedCommentAuthor, SlideComment } from '../../core-interfaces.js'
import { inch2Emu } from '../../gen-utils.js'
import { el, raw, voidEl } from '../oxml/el.js'

/** Result of resolving every slide's comments into a deck-wide author registry + per-comment numbering. */
export interface ResolvedComments {
	/** Authors in first-appearance order, ready to serialize to `commentAuthors.xml`. */
	authors: ResolvedCommentAuthor[]
	/** Per-comment `authorId`/`idx` keyed by the stored comment object. */
	meta: Map<SlideComment, { authorId: number; idx: number }>
}

/**
 * Walk every slide's comments and build the deck-wide author registry + per-comment numbering.
 *
 * Legacy comments number each comment with a *per-author* 1-based `idx`; an author's `lastIdx` is the
 * highest `idx` it used. Authors are keyed by `name`+`initials` and assigned ids in first-appearance
 * order (slide order, then insertion order). `clrIdx` mirrors `id` (one colour slot per author).
 * @param {PresSlideInternal[]} slides - all presentation slides
 * @return {ResolvedComments} the author list and per-comment metadata
 */
export function resolveCommentAuthors(slides: PresSlideInternal[]): ResolvedComments {
	const byKey = new Map<string, ResolvedCommentAuthor>()
	const authors: ResolvedCommentAuthor[] = []
	const perAuthorCount = new Map<number, number>()
	const meta = new Map<SlideComment, { authorId: number; idx: number }>()

	;(slides || []).forEach((slide) => {
		;(slide._comments || []).forEach((comment) => {
			const key = comment.author + '\0' + comment.initials
			let author = byKey.get(key)
			if (!author) {
				author = {
					id: authors.length,
					name: comment.author,
					initials: comment.initials,
					lastIdx: 0,
					clrIdx: authors.length,
				}
				byKey.set(key, author)
				authors.push(author)
				perAuthorCount.set(author.id, 0)
			}
			const idx = (perAuthorCount.get(author.id) ?? 0) + 1
			perAuthorCount.set(author.id, idx)
			author.lastIdx = idx
			meta.set(comment, { authorId: author.id, idx })
		})
	})

	return { authors, meta }
}

/**
 * Creates the presentation-level `ppt/commentAuthors.xml` part (`<p:cmAuthorLst>`).
 * @param {ResolvedCommentAuthor[]} authors - resolved deck-wide author registry
 * @return {string} XML
 */
export function makeXmlCommentAuthors(authors: ResolvedCommentAuthor[]): string {
	const cmAuthors = authors
		.map((author) =>
			voidEl('p:cmAuthor', {
				id: author.id,
				name: author.name,
				initials: author.initials,
				lastIdx: author.lastIdx,
				clrIdx: author.clrIdx,
			})
		)
		.join('')
	return (
		XML_DECL +
		CRLF +
		el(
			'p:cmAuthorLst',
			{
				'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
				'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
				'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
			},
			raw(cmAuthors)
		)
	)
}

/**
 * Creates a per-slide comments part `ppt/comments/comment{N}.xml` (`<p:cmLst>`).
 * @param {PresSlideInternal} slide - the slide whose comments are serialized
 * @param {Map} meta - per-comment `authorId`/`idx` from {@link resolveCommentAuthors}
 * @return {string} XML
 */
export function makeXmlComments(
	slide: PresSlideInternal,
	meta: Map<SlideComment, { authorId: number; idx: number }>
): string {
	const cms = (slide._comments || [])
		.map((comment) => {
			const m = meta.get(comment)
			if (!m) return '' // defensive: comment must have been seen by resolveCommentAuthors
			// Child order is fixed by CT_Comment: <p:pos> then <p:text>. pos x/y are ST_Coordinate (EMU).
			return el('p:cm', { authorId: m.authorId, dt: comment.date || null, idx: m.idx }, [
				raw(voidEl('p:pos', { x: Math.round(inch2Emu(comment.x)), y: Math.round(inch2Emu(comment.y)) })),
				raw(el('p:text', null, comment.text)),
			])
		})
		.join('')
	return (
		XML_DECL +
		CRLF +
		el(
			'p:cmLst',
			{
				'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
				'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
				'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
			},
			raw(cms)
		)
	)
}
