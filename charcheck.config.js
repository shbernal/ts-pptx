/**
 * charcheck: keep em dashes out of what a reader sees, on the site, in the docs and in
 * the README.
 *
 * The banned characters are built from their code points rather than written literally,
 * so this file does not report itself.
 *
 * Every surface errors. The gate arrived on a repo that already had 684 em dashes in
 * `docs/`, which warned under a frozen `--max-warnings` while they were worked off; that
 * backlog is now empty, so the second severity and its `DOCS_CLEAN` allowlist are gone
 * and there is one rule per surface. `lint:chars` keeps `--max-warnings 0` so that a rule
 * added at `warn` in future fails the gate rather than scrolling past in a hook.
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
 *
 * The whitespace was `[ \t]` rather than `\s` from the gate's arrival until charcheck
 * 0.2.3, working around a silent miss: `\s` matches a newline, so a trailing `\s*` ran
 * the match off the end of a hard-wrapped line and into the next node, and when that
 * node was an inline code span the match belonged to no region and the finding was
 * thrown away as a clean scan (charcheck#16). It hid 11 real dashes here. 0.2.3 reports
 * the longest match that fits inside the region instead of discarding it, so `\s` is
 * back and those cases report. The pin is exact for this reason: a downgrade below
 * 0.2.3 makes this pattern under-report again, and it does so silently.
 */
const DASH_PATTERN = '\\s*[\\u2014\\u2015]\\s*'

const MESSAGE = 'Use a comma, a colon, parentheses, or reword.'

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
			include: ['README.md', 'www/**/*.md', 'docs/**/*.md'],
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
