import {
	setDiagnosticHandler,
	build,
	readEntry,
	listEntries,
	defineRegressionSuite,
	assert,
	assertEqual,
	contentTypeOverrideParts,
	selfClosingTags,
	xmlAttributes,
	xmlOpeningTags,
} from '../../helpers.js'

// The *definition* side of `slide.addComment()` (`gen/define/comment.ts`), as distinct from
// `comments-xml.test.mjs`, which byte-pins the emitters given already-normalized `SlideComment`
// records. Everything here goes through the public builder, because that is the only way to reach
// the normalization the definer does: trimming and validating the author, deriving initials from a
// display name, and refusing — with a warning, never an exception — a comment that has no author
// or no body.
//
// The refusals were the whole gap. Every existing caller (schema-cases.js, the read round-trips)
// passes a valid author and text, so neither guard had ever executed against `dist/`, and nothing
// would have noticed the day one of them started emitting a degenerate `<p:cm>` instead.
//
// Three branches are deliberately left red, all "unreachable by construction" in the sense of
// docs/testing.md:
//   - `if (!first)` and the `?? first` fallback in `deriveAuthorInitials`. The function is only
//     ever called after `author` has been trimmed and checked non-empty, so splitting it on
//     whitespace and dropping the blanks always yields at least one word, and the last index is
//     always in range. Both guards keep the function total on its own terms; neither has an input.
//   - `if (!target._comments) target._comments = []`. `SlideBuilder` declares `_comments` as
//     `SlideComment[] = []`, and it is the definer's only caller, so the array is never missing.
//     Live if this is ever pointed at a bare internal slide; dead through the public door.
//
// A related note for whoever measures next: `gen/slide/comments.ts` shows five red branches too,
// and they are the same kind of defensive fallback (`slides || []`, `_comments || []`, `?? 0`, and
// the `!m` skip). `comments-xml.test.mjs` does exercise them, but it imports from `src/`, so it
// contributes nothing to the reported numbers — see the header there.

/** Build, capturing library warnings (`log.ts` routes every one through `console.warn`). */
async function buildCapturingWarnings(buildFn) {
	const warnings = []
	setDiagnosticHandler((d) => warnings.push(d.message))
	try {
		const result = await build(buildFn)
		return { ...result, warnings }
	} finally {
		setDiagnosticHandler(null)
	}
}

/** @param {string} tag @returns {Record<string, string>} */
const attrs = (tag) => /** @type {Record<string, string>} */ (xmlAttributes(tag))

/** Assert the package carries no trace of a comment: no parts, no rels target, no content type. */
function assertNoComments(zip, contentTypes) {
	const entries = listEntries(zip).filter((name) => /^ppt\/(comments\/|commentAuthors\.xml)/.test(name))
	assertEqual(entries.length, 0, `expected no comment parts; got ${JSON.stringify(entries)}`)
	const overrides = contentTypeOverrideParts(contentTypes).filter((part) => /comment/i.test(part))
	assertEqual(overrides.length, 0, `expected no comment content-type Overrides; got ${JSON.stringify(overrides)}`)
}

/** Every `<p:cmAuthor>` in the deck-wide registry, as an attribute map. */
async function commentAuthors(zip) {
	return selfClosingTags(await readEntry(zip, 'ppt/commentAuthors.xml'), 'p:cmAuthor').map(attrs)
}

/** Every `<p:cm>` on slide 1, as an attribute map (the element wraps `p:pos`/`p:text`). */
async function commentsOnSlide1(zip) {
	return xmlOpeningTags(await readEntry(zip, 'ppt/comments/comment1.xml'), 'p:cm').map(attrs)
}

defineRegressionSuite('Comment definition', [
	{
		// Three shapes of "no author", one per branch: options omitted entirely (reachable from
		// untyped JS, which is why the definer optional-chains), `author` absent so the typeof guard
		// falls through, and an author that is nothing but whitespace, which survives typeof and dies
		// on the trim. All three warn and drop the comment rather than emitting an authorless <p:cm>.
		name: 'a comment with no usable author warns and is dropped',
		fn: async () => {
			const { zip, warnings } = await buildCapturingWarnings((p) => {
				const s = p.addSlide()
				s.addComment(/** @type {any} */ (undefined))
				s.addComment(/** @type {any} */ ({ text: 'no author at all' }))
				s.addComment({ author: '   ', text: 'whitespace author' })
			})
			assertEqual(
				warnings.filter((message) => /requires a non-empty `author`/.test(message)).length,
				3,
				`expected one author warning per rejected comment; got ${JSON.stringify(warnings)}`
			)
			assertNoComments(zip, await readEntry(zip, '[Content_Types].xml'))
		},
	},
	{
		// Same policy on the body: `text` missing (typeof falls through) and `text: ''` (typeof holds,
		// the emptiness check catches it). Both reach the second guard, which only fires once the
		// author has already been accepted.
		name: 'a comment with no body warns and is dropped',
		fn: async () => {
			const { zip, warnings } = await buildCapturingWarnings((p) => {
				const s = p.addSlide()
				s.addComment(/** @type {any} */ ({ author: 'Ada Lovelace' }))
				s.addComment({ author: 'Ada Lovelace', text: '' })
			})
			assertEqual(
				warnings.filter((message) => /requires non-empty `text`/.test(message)).length,
				2,
				`expected one text warning per rejected comment; got ${JSON.stringify(warnings)}`
			)
			assertNoComments(zip, await readEntry(zip, '[Content_Types].xml'))
		},
	},
	{
		// A mononym has no last name to take a second letter from, so the derivation stops at one
		// character — and uppercases it, since the marker shows initials regardless of how the
		// display name was typed.
		name: 'a single-word author derives a single uppercased initial',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addComment({ author: 'plato', text: 'the cave slide needs work' })
			})
			const authors = await commentAuthors(zip)
			assertEqual(authors.length, 1, 'expected one author entry')
			assertEqual(authors[0].name, 'plato', 'the display name is kept as typed')
			assertEqual(authors[0].initials, 'P', 'the derived initial')
		},
	},
	{
		// `initials` is honoured when it says something, but a blank string is not something: it would
		// put an empty marker on the slide, so it falls through to the same derivation as an omitted
		// value. The author here is lower-case to show the derived pair is uppercased, not copied.
		name: 'blank initials fall through to the derived pair',
		fn: async () => {
			const { zip } = await build((p) => {
				p.addSlide().addComment({ author: 'ada lovelace', initials: '   ', text: 'derive these' })
			})
			const authors = await commentAuthors(zip)
			assertEqual(authors.length, 1, 'expected one author entry')
			assertEqual(authors[0].initials, 'AL', 'first and last initial, uppercased')
		},
	},
	{
		// `date` takes a `Date` as well as a string; the definer normalizes to ISO-8601 because `@dt`
		// is xsd:dateTime. An empty string is the one string that does not qualify — it would emit
		// `dt=""`, which is not a valid dateTime — so it is treated as absent and the attribute is
		// omitted, the same as never passing `date` at all.
		name: 'a Date is normalized to an ISO dt, and an empty date string is omitted',
		fn: async () => {
			const stamped = new Date(Date.UTC(2026, 6, 28, 12, 0, 0))
			const { zip } = await build((p) => {
				p.addSlide()
					.addComment({ author: 'Ada Lovelace', text: 'dated by Date', date: stamped })
					.addComment({ author: 'Ada Lovelace', text: 'dated by nothing', date: '' })
			})
			const cms = await commentsOnSlide1(zip)
			assertEqual(cms.length, 2, 'expected both comments to survive')
			assertEqual(cms[0].dt, '2026-07-28T12:00:00.000Z', 'the Date, as ISO-8601')
			assert(!('dt' in cms[1]), `expected no dt attribute on the second comment; got ${JSON.stringify(cms[1])}`)
			// Same author twice: one registry entry, numbered 1 then 2.
			assertEqual((await commentAuthors(zip)).length, 1, 'expected a single author entry')
			assertEqual([cms[0].idx, cms[1].idx].join(','), '1,2', 'per-author idx numbering')
		},
	},
])
