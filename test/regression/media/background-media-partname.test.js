import {
	PNG_1X1_DATA_URI,
	build,
	listEntries,
	readEntry,
	defineRegressionSuite,
	assert,
	assertEqual,
} from '../../helpers.js'

// A background image rel's `Target` is used twice: written into the `.rels` part, and (with
// `..` swapped for `ppt`) used verbatim as the ZIP entry name. The part used to be named after
// the layout, from `defineSlideMaster({ title })`, so a caller could put arbitrary characters
// into an OPC part name, and XML-escaping does not help: it is undone before the target is
// resolved, and the ZIP entry name is never escaped at all. Sanitizing the title closed that, but
// different titles sanitize alike, so two layouts could write one part. The name is now keyed by
// the layout, like every other media part, and the title never reaches it; the hostile titles
// below are the guard that it does not come back.
//
// The byte-identity harness cannot see any of this — no demo deck sets an image background — so
// these assertions are the only evidence.

/** Undo the XML escaping a consumer would undo before resolving the target. */
function unescapeXml(str) {
	return str
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&amp;/g, '&')
}

/**
 * Resolve a relationship Target the way a conformant OPC consumer does: as an RFC 3986 URI
 * reference against the owning part's base. Using the platform URL parser rather than string
 * surgery is the point — `?`/`#`/`%` only misbehave under real URI rules.
 */
function resolveTarget(relsPartName, target) {
	const base = 'file:///' + relsPartName.replace(/_rels\/[^/]*$/, '')
	return decodeURIComponent(new URL(unescapeXml(target), base).pathname).replace(/^\//, '')
}

async function buildWithMasterTitle(title) {
	const { zip } = await build((pres) => {
		pres.defineSlideMaster({ title, background: { data: PNG_1X1_DATA_URI } })
		pres.addSlide({ masterTitle: title })
	})
	const entries = listEntries(zip)
	// The background rel hangs off whichever part owns the background (the slide master here),
	// so find it by content rather than assuming a part name.
	for (const relsPartName of entries.filter((name) => name.endsWith('.rels'))) {
		const match = (await readEntry(zip, relsPartName)).match(/Target="([^"]*media[^"]*)"/)
		if (match) return { entries, relsPartName, target: match[1] }
	}
	throw new Error('no media relationship found in any .rels part; entries: ' + entries.join(', '))
}

// Each of these broke a different way before the fix: `%` made the target undecodable,
// `?`/`#` truncated the resolved path at the query/fragment, `/` pushed the media into a
// subdirectory. The rest are shapes worth pinning even though they always worked.
const HOSTILE_TITLES = [
	['percent (invalid escape)', '100%done'],
	['question mark (query)', 'what?now'],
	['hash (fragment)', 'tag#1'],
	['forward slash (path segment)', 'a/b'],
	['backslash', 'a\\b'],
	['ampersand', 'R&D Deck'],
	['angle brackets', 'a<b>c'],
	['double quotes', 'say "hi"'],
	['dot segments', '..'],
	['whitespace only', '   '],
	['non-ASCII', '設計マスター'],
]

defineRegressionSuite(
	'background media part name is derived safely from a caller-supplied master title',
	HOSTILE_TITLES.map(([label, title]) => ({
		name: `${label}: rel target resolves to a real package part`,
		fn: async () => {
			const { entries, relsPartName, target } = await buildWithMasterTitle(title)
			const resolved = resolveTarget(relsPartName, target)
			assert(
				entries.includes(resolved),
				`rel Target ${JSON.stringify(target)} resolved to ${JSON.stringify(resolved)}, ` +
					`which is not in the package. media entries: ${entries.filter((e) => e.includes('media/')).join(', ')}`
			)
		},
	}))
)

defineRegressionSuite('background media part name character set', [
	{
		name: 'part name is restricted to URI-unreserved characters',
		fn: async () => {
			for (const [, title] of HOSTILE_TITLES) {
				const { entries } = await buildWithMasterTitle(title)
				const media = entries.filter((name) => name.includes('media/'))
				assertEqual(media.length, 1, `expected exactly one media part for title ${JSON.stringify(title)}`)
				const leaf = media[0].slice('ppt/media/'.length)
				assert(
					/^[A-Za-z0-9._-]+$/.test(leaf),
					`media part name ${JSON.stringify(leaf)} (from title ${JSON.stringify(title)}) ` +
						'contains characters outside [A-Za-z0-9._-]'
				)
			}
		},
	},
	{
		name: 'the part name is keyed by the layout and carries none of its title',
		fn: async () => {
			// The title used to be sanitized into the name, so "My Master Slide" read
			// `My-Master-Slide-image-1.png`, and "A B" and "A-B" both read `A-B-image-1.png`: one part
			// for two layouts. Naming it the way every other media part is named leaves nothing of the
			// title to collide on.
			for (const title of ['My Master Slide', '   ', 'what?now']) {
				const { entries } = await buildWithMasterTitle(title)
				const media = entries.find((name) => name.includes('media/'))
				assert(/^ppt\/media\/image-sl-\d+-1\.png$/.test(media ?? ''), `layout-keyed part name; got ${media}`)
				for (const word of title.split(/[^A-Za-z0-9]+/).filter(Boolean)) {
					assert(
						!(media ?? '').includes(word),
						`the part name ${media} carries none of the title ${JSON.stringify(title)}`
					)
				}
			}
		},
	},
])
