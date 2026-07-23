// Read accessor for legacy review comments: Slide.comments decodes the slide's
// p:cmLst/p:cm and resolves each comment's author through the deck-wide
// Presentation.commentAuthors registry (p:cmAuthorLst). Authorable write→read: the
// writer emits these via slide.addComment(...), so this is a genuine round-trip.

import { describe, test } from 'vitest'
import { assertEqual } from '../helpers.js'
import { authorRead, schemaErrors, validatorInstalled } from './authored.js'

const EMU_PER_INCH = 914400
const inchToEmu = (inches) => Math.round(inches * EMU_PER_INCH)

describe('Slide.comments + Presentation.commentAuthors', () => {
	test('reads each comment and resolves its author through the registry', async () => {
		const { presentation } = await authorRead((pres) => {
			const s1 = pres.addSlide()
			s1.addComment({ author: 'Ada Lovelace', initials: 'AL', text: 'Tighten this headline', x: 1, y: 0.5 })
			s1.addComment({ author: 'Ada Lovelace', initials: 'AL', text: 'Second note', x: 2, y: 1 })
			const s2 = pres.addSlide()
			s2.addComment({ author: 'Grace Hopper', initials: 'GH', text: 'Nice chart', x: 3, y: 2 })
		})

		// Deck-wide author registry: two authors in first-appearance order.
		const authors = presentation.commentAuthors
		assertEqual(authors.length, 2, 'two distinct authors')
		assertEqual(authors[0].name, 'Ada Lovelace', 'first author name')
		assertEqual(authors[0].id, 0, 'first author id is 0')
		assertEqual(authors[0].lastIdx, 2, 'Ada used two per-author indices')
		assertEqual(authors[1].name, 'Grace Hopper', 'second author name')
		assertEqual(authors[1].id, 1, 'second author id is 1')

		// Slide 1: two comments, both resolved to Ada.
		const c1 = presentation.slides[0].comments
		assertEqual(c1.length, 2, 'slide 1 has two comments')
		assertEqual(c1[0].author, 'Ada Lovelace', 'author resolved via authorId')
		assertEqual(c1[0].authorInitials, 'AL', 'author initials resolved')
		assertEqual(c1[0].authorId, 0, 'raw authorId kept')
		assertEqual(c1[0].idx, 1, 'first per-author index')
		assertEqual(c1[0].text, 'Tighten this headline', 'body text read from p:text')
		assertEqual(c1[0].x, inchToEmu(1), 'marker x in EMU (1in)')
		assertEqual(c1[0].y, inchToEmu(0.5), 'marker y in EMU (0.5in)')
		assertEqual(c1[1].idx, 2, 'second comment gets per-author idx 2')

		// Slide 2: one comment, resolved to Grace.
		const c2 = presentation.slides[1].comments
		assertEqual(c2.length, 1, 'slide 2 has one comment')
		assertEqual(c2[0].author, 'Grace Hopper', 'slide 2 author resolves')
		assertEqual(c2[0].idx, 1, "Grace's per-author index restarts at 1")
	})

	test('a comment carries its authored date when set', async () => {
		const iso = '2026-07-23T12:00:00.000Z'
		const { presentation } = await authorRead((pres) => {
			pres.addSlide().addComment({ author: 'Ada Lovelace', text: 'dated', date: iso })
		})
		assertEqual(presentation.slides[0].comments[0].date, iso, 'the @dt timestamp round-trips')
	})

	test('a deck with no comments reads [] for authors and per-slide comments', async () => {
		const { presentation } = await authorRead((pres) => {
			pres.addSlide()
		})
		assertEqual(presentation.commentAuthors.length, 0, 'no commentAuthors part → []')
		assertEqual(presentation.slides[0].comments.length, 0, 'no comments part → []')
	})

	test.skipIf(!validatorInstalled)('a deck with comments stays schema-valid', async () => {
		const { buf } = await authorRead((pres) => {
			pres.addSlide().addComment({ author: 'Ada Lovelace', text: 'valid', x: 1, y: 1 })
		})
		assertEqual((await schemaErrors(buf)).length, 0, 'no schema violations')
	})
})
