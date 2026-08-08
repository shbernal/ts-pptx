#!/usr/bin/env node
// Print every docs page with its `summary` and `read_when` hints.
//
// The discovery entry point for an agent or a contributor: one screen answering "what is
// documented here, and when should I read it?".

import path from 'node:path'

import { compactStrings, parseFrontmatter, requireDocsDir, walkDocs } from './docs-frontmatter.mjs'

const docsDir = requireDocsDir('docs:list')

const out = ['Listing all markdown files in docs folder:']
for (const rel of walkDocs(docsDir)) {
	const { data, error } = parseFrontmatter(path.join(docsDir, rel))
	const summary = String(data.summary ?? '').trim()
	if (summary) {
		out.push(`${rel} - ${summary}`)
		const readWhen = compactStrings(data.read_when)
		if (readWhen.length > 0) out.push(`  Read when: ${readWhen.join('; ')}`)
	} else {
		out.push(`${rel} - [${error ?? 'summary key missing'}]`)
	}
}
out.push(
	'\nReminder: when a task matches any "Read when" hint above, read that doc before coding and update docs when behavior changes.'
)
console.log(out.join('\n'))
