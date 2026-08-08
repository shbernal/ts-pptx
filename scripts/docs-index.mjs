#!/usr/bin/env node
// Generate `docs/doc-index.md`: a read_when discovery index.
//
// Every narrative doc declares a `read_when:` list in its frontmatter, but nothing aggregates
// them, so there is no single place to answer "which doc covers this task?". This walks `docs/`,
// collects each page carrying a non-empty `read_when`, and emits a generated `docs/doc-index.md`
// grouping every page under its scenario hints.
//
// Generated pages without `read_when` (e.g. the typedoc `reference/api/` tree) are skipped by
// construction. The output is a generated artifact (gitignored, like `reference/api/` and
// `public/llms*.txt`); regenerate it with `pnpm run docs:index`. It is validated by
// `scripts/docs-check.mjs` (frontmatter + links), which runs after generation in `docs:build`.

import { writeFileSync } from 'node:fs'
import path from 'node:path'

import { compactStrings, parseFrontmatter, requireDocsDir, walkDocs } from './docs-frontmatter.mjs'

const OUTPUT_NAME = 'doc-index.md'

const HEADER = [
	'---',
	'doc-schema-version: 1',
	'title: "Documentation Index"',
	'summary: "Generated read_when discovery index: every guide and the scenarios that should send you to it."',
	'read_when:',
	'  - Looking for which doc covers a task or scenario',
	'  - Discovering documentation by when to read it',
	'doc_type: "reference"',
	'---',
	'',
	'<!-- GENERATED FILE. Do not edit by hand.',
	'     Regenerate with `pnpm run docs:index` (runs in `docs:prepare`).',
	'     Source: the `read_when:` frontmatter across docs/. -->',
	'',
	'# Documentation Index',
	'',
	'Each doc below declares, in its frontmatter, the situations in which you',
	'should read it. This page aggregates those `read_when` hints so you can find',
	'the right doc by task. When a task matches a hint, read that doc before',
	'coding and update it when behavior changes.',
	'',
]

const docsDir = requireDocsDir('docs:index')

const entries = []
for (const rel of walkDocs(docsDir, OUTPUT_NAME)) {
	const { data } = parseFrontmatter(path.join(docsDir, rel))
	const hints = compactStrings(data.read_when)
	if (hints.length === 0) continue // generated/api pages and any page without hints
	entries.push({
		rel,
		title: String(data.title ?? '').trim() || rel,
		summary: String(data.summary ?? '').trim(),
		readWhen: hints,
	})
}

const lines = [...HEADER]
for (const entry of entries) {
	lines.push(`## [${entry.title}](${entry.rel})`, '')
	if (entry.summary) lines.push(entry.summary, '')
	lines.push('Read when:', '')
	for (const hint of entry.readWhen) lines.push(`- ${hint}`)
	lines.push('')
}

const outputPath = path.join(docsDir, OUTPUT_NAME)
writeFileSync(outputPath, lines.join('\n').replace(/\s+$/, '') + '\n', 'utf8')
console.log(
	`docs:index: wrote ${path.relative(process.cwd(), outputPath).split(path.sep).join('/')} (${entries.length} doc(s))`
)
