// The two transforms that make the scoped alias differ from the canonical package.
//
// Everything else about the alias is a byte copy, so these two functions are the entire
// surface where the second publish can be wrong — and both fail quietly rather than
// loudly. A manifest that kept the canonical name would publish the canonical package a
// second time under its own name (a no-op that reports success), and a README whose
// banner landed in the wrong place still renders.

import { describe, expect, test } from 'vitest'
import { ALIAS_NAME, aliasManifest, aliasReadme } from '../../scripts/alias-package.mjs'

describe('aliasManifest', () => {
	test('substitutes the alias name', () => {
		expect(aliasManifest({ name: 'pptx-ts', version: '3.6.0' }).name).toBe(ALIAS_NAME)
	})

	// The version must travel with the canonical package, because the publish workflow
	// stages the alias from the same tree it just published and the two would otherwise
	// drift apart silently.
	test('keeps the canonical version when none is given', () => {
		expect(aliasManifest({ name: 'pptx-ts', version: '3.6.0' }).version).toBe('3.6.0')
	})

	// Bootstrap only: npm needs the package to exist before a trusted publisher can be
	// configured for it, and that first publish is deliberately not a release version.
	test('takes a version override', () => {
		expect(aliasManifest({ name: 'pptx-ts', version: '3.6.0' }, { version: '0.0.1' }).version).toBe('0.0.1')
	})

	// The one field that must not travel. `prepack` runs tsdown, which finds the
	// repository's config by searching upward and whose `clean: true` then deletes the
	// staged `dist/`: a publish that forgets `--ignore-scripts` would gut the directory,
	// and the retry with the flag would push a package with no code in it.
	test('drops the scripts block', () => {
		const canonical = { name: 'pptx-ts', version: '3.6.0', scripts: { prepack: 'pnpm run build' } }
		expect(aliasManifest(canonical).scripts).toBeUndefined()
	})

	test('leaves every other field alone', () => {
		const canonical = {
			name: 'pptx-ts',
			version: '3.6.0',
			exports: { '.': './dist/index.js' },
			files: ['dist'],
			devDependencies: { vitest: '^4.1.11' },
		}
		const aliased = aliasManifest(canonical)
		expect(aliased.exports).toEqual(canonical.exports)
		expect(aliased.files).toEqual(canonical.files)
		expect(aliased.devDependencies).toEqual(canonical.devDependencies)
	})

	// A one-line diff against the canonical manifest is the claim the whole approach
	// rests on: reordered keys would make that diff unreadable and hide a real change.
	test('keeps key order', () => {
		const canonical = { name: 'pptx-ts', version: '3.6.0', license: 'MIT' }
		expect(Object.keys(aliasManifest(canonical))).toEqual(['name', 'version', 'license'])
	})

	// The deletion is on the copy, not on the caller's object. Getting that wrong would
	// strip `scripts` out of the repository's own manifest in memory, which is a very
	// quiet way to break whatever ran next.
	test('does not mutate the manifest it was given', () => {
		const canonical = { name: 'pptx-ts', version: '3.6.0', scripts: { build: 'tsdown' } }
		aliasManifest(canonical, { version: '0.0.1' })
		expect(canonical).toEqual({ name: 'pptx-ts', version: '3.6.0', scripts: { build: 'tsdown' } })
	})
})

describe('aliasReadme', () => {
	const canonicalName = 'pptx-ts'

	test('puts the banner under the title, not above it', () => {
		const readme = ['# ts-pptx', '', '![badge](x)', '', 'Body.'].join('\n')
		const lines = aliasReadme(readme, { canonicalName }).split('\n')
		expect(lines[0]).toBe('# ts-pptx')
		expect(lines.find((line) => line.startsWith('>'))).toContain('is an alias')
	})

	test('names both packages', () => {
		const out = aliasReadme('# ts-pptx\n\nBody.\n', { canonicalName })
		expect(out).toContain(`\`${ALIAS_NAME}\``)
		expect(out).toContain(`https://www.npmjs.com/package/${canonicalName}`)
	})

	test('keeps the whole canonical body', () => {
		const readme = '# ts-pptx\n\n## Install\n\n```bash\npnpm add pptx-ts\n```\n'
		expect(aliasReadme(readme, { canonicalName })).toContain('```bash\npnpm add pptx-ts\n```')
	})

	// Being unplaced is worse than being ugly: a README with no `# ` heading still has to
	// carry the banner somewhere a reader sees it.
	test('prepends when there is no title', () => {
		const out = aliasReadme('Body only.\n', { canonicalName })
		expect(out.startsWith('>')).toBe(true)
		expect(out).toContain('Body only.')
	})

	// The first `# ` wins, not the last: a `# ` further down (inside a fenced block, say)
	// must not pull the banner into the middle of the document.
	test('uses the first title', () => {
		const readme = ['# ts-pptx', '', 'Body.', '', '# Appendix'].join('\n')
		const lines = aliasReadme(readme, { canonicalName }).split('\n')
		expect(lines.findIndex((line) => line.startsWith('>'))).toBe(2)
	})
})
