/**
 * charcheck: keep em dashes out of what a reader sees, on the site, in the docs and in
 * the README.
 *
 * The banned characters are built from their code points rather than written literally,
 * so this file does not report itself.
 *
 * Two rules cover the same prose at two severities, because this gate arrived on a repo
 * that already had 695 em dashes in `docs/`. Erroring on all of them would have meant
 * either rewriting the author's prose wholesale or turning the hook off, and a hook that
 * blames you for text you did not write is a hook you disable within a week. So the
 * surfaces that are clean today error, the docs backlog warns, and `--max-warnings` in
 * `lint:chars` freezes the backlog at its current size: it can shrink, never grow.
 *
 * To retire a docs page from the backlog, clean it, add it to `DOCS_CLEAN`, and lower the
 * `--max-warnings` number in the `lint:chars` script by what you removed.
 */

import { strategies } from 'charcheck/config'

/** @param {number} codePoint */
const cp = (codePoint) => String.fromCodePoint(codePoint)

/** Em dash and horizontal bar. */
const DASHES = [cp(0x2014), cp(0x2015)]

/**
 * Matches the dash together with the space around it, which is what `clauseSeparator`
 * needs in order to put a line break back rather than swallow it. Written as regex
 * escapes so the characters never appear literally in this file.
 */
const DASH_PATTERN = '\\s*[\\u2014\\u2015]\\s*'

const MESSAGE = 'Use a comma, a colon, parentheses, or reword.'

/**
 * Docs pages cleaned of em dashes and promoted out of the backlog below. Every entry
 * here is a page that now errors rather than warns; the list only grows.
 *
 * @type {string[]}
 */
const DOCS_CLEAN = []

/**
 * Written by a generator, so a finding has no author to tell. `docs/reference/api/` is
 * TypeDoc output, `doc-index.md` and `llms*.txt` are built by the docs scripts, and all
 * three are gitignored; the pattern is here so a scan of a built tree agrees with a scan
 * of a clean one.
 */
const GENERATED = [
	'dist/**',
	'coverage/**',
	'docs/.vitepress/cache/**',
	'docs/.vitepress/dist/**',
	'docs/reference/api/**',
	'docs/doc-index.md',
	'docs/public/llms*.txt',
	'pnpm-lock.yaml',
]

export default {
	rules: [
		{
			// Prose only: fenced and inline code are exempt, because a dash inside a code
			// sample is part of the sample and not something to reword.
			id: 'no-em-dash-in-prose',
			pattern: DASH_PATTERN,
			scope: 'markdown',
			fix: strategies.clauseSeparator,
			message: MESSAGE,
			include: ['README.md', 'www/**/*.md', ...DOCS_CLEAN],
		},
		{
			// The `docs/` backlog. Same rule, same fix, reported rather than enforced, and
			// held to its current size by `--max-warnings`.
			id: 'no-em-dash-in-prose-backlog',
			pattern: DASH_PATTERN,
			scope: 'markdown',
			fix: strategies.clauseSeparator,
			severity: 'warn',
			message: MESSAGE,
			include: ['docs/**/*.md'],
			exclude: DOCS_CLEAN,
		},
		{
			// The site's one Vue component: template text and allowlisted attributes. Its
			// comments and its stylesheet are exempt, which is the point of `markup` over
			// `raw` on a file that is mostly neither prose nor code.
			id: 'no-em-dash-on-the-page',
			chars: DASHES,
			scope: 'markup',
			message: MESSAGE,
			include: ['www/**/*.vue'],
		},
		{
			// Site config and theme code: string literals only. That is where the nav labels,
			// the tagline and the meta description live, and it is the only text in these
			// files a visitor can read. Comments keep their dashes, as they do repo-wide.
			id: 'no-em-dash-in-site-text',
			chars: DASHES,
			scope: 'strings',
			message: MESSAGE,
			include: ['docs/.vitepress/config.mts', 'docs/.vitepress/theme/**/*.ts', 'www/**/*.ts'],
		},
	],
	ignore: GENERATED,
}
