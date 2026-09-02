#!/usr/bin/env node
// Create a docs page with the standard frontmatter, optionally wiring it into docs.json nav.
//
//   pnpm run docs:new -- guides/setup --title "Setup" --read-when "Setting the repo up"

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'
import { ALLOWED_DOC_TYPES } from './docs-frontmatter.mjs'
import { ROOT, parseCliOrExit } from './script-utils.mjs'

/**
 * `guides/my-page.md`, `/docs/guides/my-page`, … → `guides/my-page`.
 * @param {string} slug
 * @returns {string}
 */
function pageKey(slug) {
	let key = slug.trim().replace(/^\/+|\/+$/g, '')
	if (key.startsWith('docs/')) key = key.slice('docs/'.length)
	return key.replace(/\.md$/, '')
}

/**
 * `my-new-page` → `My New Page`.
 * @param {string} slug
 * @returns {string}
 */
function titleFromSlug(slug) {
	return path
		.basename(slug)
		.split(/[-_]+/)
		.filter(Boolean)
		.map((part) => part.charAt(0).toUpperCase() + part.slice(1))
		.join(' ')
}

/**
 * Every `{ group, pages }` object anywhere in a navigation tree.
 * @param {unknown} value
 * @returns {NavGroup[]}
 */
function collectGroups(value) {
	/** @type {NavGroup[]} */
	const groups = []
	if (Array.isArray(value)) {
		for (const item of value) groups.push(...collectGroups(item))
	} else if (value && typeof value === 'object') {
		const record = /** @type {Record<string, unknown>} */ (value)
		if (typeof record.group === 'string' && Array.isArray(record.pages)) groups.push(/** @type {NavGroup} */ (record))
		for (const item of Object.values(record)) groups.push(...collectGroups(item))
	}
	return groups
}

/**
 * One navigation group in `docs/docs.json`.
 * @typedef {{group: string, pages: string[]}} NavGroup
 */

/**
 * @param {string} docsJsonPath
 * @param {string} key the page key to add
 * @param {string} groupName the group to add it under, created if absent
 * @returns {void}
 */
function addToNav(docsJsonPath, key, groupName) {
	if (!existsSync(docsJsonPath)) return
	const config = JSON.parse(readFileSync(docsJsonPath, 'utf8'))
	config.navigation ??= []
	if (!Array.isArray(config.navigation)) throw new Error('docs/docs.json navigation must be a list')

	let target = collectGroups(config.navigation).find((group) => group.group === groupName)
	if (!target) {
		target = { group: groupName, pages: [] }
		config.navigation.push(target)
	}
	if (!Array.isArray(target.pages)) throw new Error(`docs/docs.json group \`${groupName}\` pages must be a list`)
	if (!target.pages.includes(key)) target.pages.push(key)

	writeFileSync(docsJsonPath, JSON.stringify(config, null, 2) + '\n', 'utf8')
}

const USAGE = `Create a docs page with the standard frontmatter, optionally wiring it into docs.json nav.

  pnpm run docs:new -- guides/setup --title "Setup" --read-when "Setting the repo up"

Options:
  --title <text>       page title (default: derived from the slug)
  --summary <text>     one-line summary for the frontmatter
  --type <type>        one of ${ALLOWED_DOC_TYPES.join(', ')} (default: guide)
  --read-when <hint>   a read_when hint; repeat for more than one
  --nav-group <name>   docs.json nav group to add the page to
  --force              overwrite an existing page
  -h, --help           show this message`

const { values, positionals } = parseCliOrExit(process.argv.slice(2), {
	usage: USAGE,
	allowPositionals: true,
	options: {
		title: { type: 'string' },
		summary: { type: 'string' },
		type: { type: 'string', default: 'guide' },
		'read-when': { type: 'string', multiple: true, default: [] },
		'nav-group': { type: 'string' },
		force: { type: 'boolean', default: false },
	},
})

const slug = positionals[0]
if (!slug) {
	console.error('docs:new: a page slug is required\n')
	console.error(USAGE)
	process.exit(2)
}
if (!ALLOWED_DOC_TYPES.includes(values.type)) {
	console.error(`docs:new: --type must be one of: ${ALLOWED_DOC_TYPES.join(', ')}`)
	process.exit(2)
}

// Resolved from the repo root, not the caller's cwd: `pnpm run` happens to set the cwd to the
// package root, but `node scripts/docs-new.mjs` from a subdirectory would otherwise scaffold a
// second `docs/` tree there.
const docsDir = path.resolve(ROOT, 'docs')
mkdirSync(docsDir, { recursive: true })

const key = pageKey(slug)
const title = values.title || titleFromSlug(key)
const summary = values.summary || `Documentation page for ${title}.`
// `parseCli` types its `values` loosely (see its doc), so the `multiple: true` flag arrives
// without its element type; the annotation is what the template literal below needs.
/** @type {string[]} */
const readWhen = values['read-when'].length > 0 ? values['read-when'] : [`Working on ${title}`]

const target = path.resolve(docsDir, `${key}.md`)
if (!target.startsWith(docsDir + path.sep)) {
	console.error('docs:new: page slug must stay under docs/')
	process.exit(1)
}
if (existsSync(target) && !values.force) {
	console.error(`docs:new: ${target} already exists; pass --force to overwrite`)
	process.exit(1)
}

mkdirSync(path.dirname(target), { recursive: true })
writeFileSync(
	target,
	`---
doc-schema-version: 1
title: "${title}"
summary: "${summary}"
read_when:
${readWhen.map((hint) => `  - ${hint}`).join('\n')}
doc_type: "${values.type}"
---

# ${title}

Add the useful project-specific content here.
`,
	'utf8'
)

if (values['nav-group']) addToNav(path.join(docsDir, 'docs.json'), key, values['nav-group'])

console.log(`docs:new: wrote ${target}`)
