import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import path from 'node:path'

const root = process.cwd()
const docsDir = path.join(root, 'docs')
const publicDir = path.join(docsDir, 'public')
const docsConfig = JSON.parse(readFileSync(path.join(docsDir, 'docs.json'), 'utf8'))
const pkg = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8'))
const baseUrl = (process.env.DOCS_BASE_URL ?? 'https://shbernal.github.io/PptxGenJS/').replace(/\/?$/, '/')
const excludedDirs = new Set(['.vitepress', 'archive', 'public', 'research'])

function walkMarkdown(dir) {
	const out = []
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		if (entry.isDirectory()) {
			if (!excludedDirs.has(entry.name)) out.push(...walkMarkdown(path.join(dir, entry.name)))
			continue
		}
		if (entry.isFile() && entry.name.endsWith('.md')) {
			if (entry.name === 'README.md' && existsSync(path.join(dir, 'index.md'))) continue
			out.push(path.join(dir, entry.name))
		}
	}
	return out.sort()
}

function parseFrontmatter(markdown) {
	if (!markdown.startsWith('---\n')) return [{}, markdown]
	const end = markdown.indexOf('\n---\n', 4)
	if (end === -1) return [{}, markdown]

	const frontmatter = {}
	for (const line of markdown.slice(4, end).split('\n')) {
		const match = line.match(/^([A-Za-z0-9_-]+):\s*(.*)$/)
		if (!match) continue
		frontmatter[match[1]] = match[2].replace(/^"|"$/g, '')
	}
	return [frontmatter, markdown.slice(end + 5).trimStart()]
}

function titleFromBody(body, fallback) {
	const heading = body.match(/^#\s+(.+)$/m)
	return heading ? heading[1].trim() : fallback
}

function routeFor(filePath) {
	const rel = path.relative(docsDir, filePath).split(path.sep).join('/')
	const withoutExt = rel.replace(/\.md$/, '')
	if (withoutExt === 'index') return baseUrl
	const route =
		withoutExt.endsWith('/index') || withoutExt.endsWith('/README')
			? withoutExt.slice(0, withoutExt.lastIndexOf('/'))
			: withoutExt
	return new URL(route ? `${route}/` : '', baseUrl).toString()
}

function pageRecord(filePath) {
	const raw = readFileSync(filePath, 'utf8')
	const [frontmatter, body] = parseFrontmatter(raw)
	const rel = path.relative(docsDir, filePath).split(path.sep).join('/')
	const fallbackTitle = rel.replace(/\.md$/, '')
	return {
		body,
		rel,
		summary: frontmatter.summary ?? '',
		title: frontmatter.title ?? titleFromBody(body, fallbackTitle),
		url: routeFor(filePath),
	}
}

function docsNavigationOrder(records) {
	const byRouteKey = new Map(records.map((record) => [record.rel.replace(/\.md$/, ''), record]))
	const ordered = []
	for (const group of docsConfig.navigation ?? []) {
		for (const page of group.pages ?? []) {
			const record = byRouteKey.get(page)
			if (record) ordered.push(record)
		}
	}
	for (const record of records) {
		if (!ordered.includes(record)) ordered.push(record)
	}
	return ordered
}

const records = docsNavigationOrder(walkMarkdown(docsDir).map(pageRecord))

const llms = [
	`# ${docsConfig.name}`,
	'',
	`> ${docsConfig.description}`,
	'',
	`Package: ${pkg.name}`,
	`Version: ${pkg.version}`,
	'',
	'## Docs',
	'',
	...records.map((record) => `- [${record.title}](${record.url})${record.summary ? `: ${record.summary}` : ''}`),
	'',
].join('\n')

const llmsFull = [
	`# ${docsConfig.name}`,
	'',
	`> ${docsConfig.description}`,
	'',
	...records.flatMap((record) => [
		`## ${record.title}`,
		'',
		`URL: ${record.url}`,
		'',
		record.summary ? `${record.summary}\n` : '',
		record.body.trim(),
		'',
	]),
].join('\n')

mkdirSync(publicDir, { recursive: true })
writeFileSync(path.join(publicDir, 'llms.txt'), llms, 'utf8')
writeFileSync(path.join(publicDir, 'llms-full.txt'), llmsFull, 'utf8')

console.log(`generated ${path.relative(root, path.join(publicDir, 'llms.txt'))}`)
console.log(`generated ${path.relative(root, path.join(publicDir, 'llms-full.txt'))}`)
