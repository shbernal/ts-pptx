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
 * The whitespace is `[ \t]`, not `\s`, and that is load-bearing rather than pedantic.
 * `\s` matches a newline, so a trailing `\s*` runs the match off the end of a
 * hard-wrapped line and into the next node; when that node is an inline code span,
 * charcheck drops the finding silently and reports a clean scan (charcheck#16). That is
 * not a hypothetical: it hid 11 real dashes here, every one of them a dash ending a
 * wrapped line whose continuation began with a code span. Horizontal-only whitespace
 * keeps the match inside the line, which reports all of them and leaves the fixer's
 * output unchanged on the cases both forms could already see. Restore `\s` after the
 * upstream fix ships, re-run, and expect those findings to keep being reported.
 */
const DASH_PATTERN = '[ \\t]*[\\u2014\\u2015][ \\t]*'

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
