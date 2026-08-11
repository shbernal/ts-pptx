// Modern (2018) comments: the p188 schema PowerPoint writes today — deck-wide
// GUID-keyed authors in ppt/authors.xml, and per-slide threaded comments in
// ppt/comments/modernComment_*.xml. Read-only (the writer emits the legacy
// schema). The fixture modern-comments.pptx was authored by desktop PowerPoint via
// the Comments.Add2 / Replies.Add2 COM API: slide 2 carries one comment by Ada
// Lovelace with one reply by Grace Hopper; slide 1 has none.

import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, test } from 'vitest'

import { assert, assertEqual } from '../helpers.js'
import { authorRead } from './authored.js'
import { openFixture } from './corpus.js'

const __dirname = path.dirname(fileURLToPath(import.meta.url))

async function load(name) {
	return openFixture(name)
}

describe('modern (2018) comments', () => {
	test('commentSchema reports "modern" for a modernComment deck', async () => {
		const pres = await load('modern-comments.pptx')
		assertEqual(pres.commentSchema, 'modern', 'a modernComment_* part → modern')
		// The modern schema does not populate the legacy accessors.
		assertEqual(pres.commentAuthors.length, 0, 'no legacy commentAuthors')
		assertEqual(pres.slides[1].comments.length, 0, 'no legacy per-slide comments')
	})

	test('modernCommentAuthors decode with GUID ids, userId and providerId', async () => {
		const authors = (await load('modern-comments.pptx')).modernCommentAuthors
		assertEqual(authors.length, 2, 'two authors in authors.xml')
		assertEqual(authors[0].id, '{E8A64ABA-B822-D93A-1829-32C11809D46F}', 'author id is the GUID string')
		assertEqual(authors[0].name, 'Ada Lovelace', 'first author name')
		assertEqual(authors[0].initials, 'AL', 'first author initials')
		assertEqual(authors[0].userId, 'ada@example.com', 'userId decoded')
		assertEqual(authors[0].providerId, 'Windows Live', 'providerId decoded')
		assertEqual(authors[1].name, 'Grace Hopper', 'second author name')
	})

	test('a slide with no modern comments part reads []', async () => {
		const pres = await load('modern-comments.pptx')
		assertEqual(pres.slides[0].modernComments.length, 0, 'slide 1 has no modern comments')
	})

	test('a comment resolves its author, position, timestamp and body text', async () => {
		const [comment] = (await load('modern-comments.pptx')).slides[1].modernComments
		assert(comment, 'slide 2 has one modern comment')
		assertEqual(comment.id, '{0B4BE8B0-B578-4C2F-B9E5-A882080696D3}', 'comment GUID id')
		assertEqual(comment.author, 'Ada Lovelace', 'author resolved via @authorId GUID')
		assertEqual(comment.authorInitials, 'AL', 'author initials resolved')
		assertEqual(comment.authorId, '{E8A64ABA-B822-D93A-1829-32C11809D46F}', 'raw authorId GUID kept')
		assertEqual(comment.text, 'Tighten this headline', 'body text read from p188:txBody')
		assertEqual(comment.created, '2026-07-24T08:10:07.838', 'created timestamp kept as written')
		assertEqual(comment.x, 1524000, 'marker x in EMU (p188:pos)')
		assertEqual(comment.y, 1016000, 'marker y in EMU (p188:pos)')
	})

	test('the reply is nested under the comment and resolves its own author', async () => {
		const [comment] = (await load('modern-comments.pptx')).slides[1].modernComments
		assertEqual(comment.replies.length, 1, 'one reply, nested not flattened')
		const reply = comment.replies[0]
		assertEqual(reply.id, '{95F4A554-CD5C-4107-8752-201F04D46DCB}', 'reply GUID id')
		assertEqual(reply.author, 'Grace Hopper', 'reply resolves the second author')
		assertEqual(reply.authorInitials, 'GH', 'reply author initials')
		assertEqual(reply.text, 'Agreed, will do', 'reply body text')
		assertEqual(reply.created, '2026-07-24T08:10:07.868', 'reply created timestamp')
		assertEqual(reply.x, null, 'a reply carries no position')
		assertEqual(reply.y, null, 'a reply carries no position')
		assertEqual(reply.replies.length, 0, 'a reply has no further replies')
	})
})

describe('commentSchema discriminator on non-modern decks', () => {
	test('a legacy-comment deck reports "legacy"', async () => {
		// Author a legacy deck with the writer and load it back through the read model.
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addComment({ author: 'Ada Lovelace', text: 'legacy note', x: 1, y: 1 })
		})
		assertEqual(presentation.commentSchema, 'legacy', 'a commentN.xml part → legacy')
		assertEqual(presentation.modernCommentAuthors.length, 0, 'no modern authors on a legacy deck')
	})

	test('a deck with no comments reports "none"', async () => {
		const pres = await load('empty.pptx')
		assertEqual(pres.commentSchema, 'none', 'no comment parts → none')
		assertEqual(pres.slides[0]?.modernComments.length ?? 0, 0, 'no modern comments')
	})
})
