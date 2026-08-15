#!/usr/bin/env node
// Create a docs page with the standard frontmatter, optionally wiring it into docs.json nav.
//
//   pnpm run docs:new -- guides/setup --title "Setup" --read-when "Setting the repo up"

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'
import { parseArgs } from 'node:util'

import { ALLOWED_DOC_TYPES } from './docs-frontmatter.mjs'

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

const { values, positionals } = parseArgs({
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
	console.error(
		'usage: node scripts/docs-new.mjs <slug> [--title T] [--summary S] [--type T] [--read-when H]... [--nav-group G] [--force]'
	)
	process.exit(2)
}
if (!ALLOWED_DOC_TYPES.includes(values.type)) {
	console.error(`docs:new: --type must be one of: ${ALLOWED_DOC_TYPES.join(', ')}`)
	process.exit(2)
}

const docsDir = path.resolve('docs')
mkdirSync(docsDir, { recursive: true })

const key = pageKey(slug)
const title = values.title || titleFromSlug(key)
const summary = values.summary || `Documentation page for ${title}.`
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
