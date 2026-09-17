// The site sidebar for the generated API reference: one collapsed group per entry point, and
// inside it one collapsed group per kind of member. The kinds and their members come from each
// module's `README.md`, which is TypeDoc's own grouping, so this file holds no table of kinds.
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import path from 'node:path'

import { routeForPage } from './docs-frontmatter.mjs'

/**
 * @typedef {{ text: string, link?: string, collapsed?: boolean, items?: SidebarItem[] }} SidebarItem
 * The subset of VitePress's `DefaultTheme.SidebarItem` this file writes.
 */

/**
 * VitePress renders sidebar text as HTML, so a label is escaped, and a deprecated member, which
 * TypeDoc strikes through as `~~Name~~`, keeps its strike.
 * @param {string} label - the markdown text of a list link
 * @returns {string}
 */
function sidebarText(label) {
	const struck = /^~~(.*)~~$/.exec(label)
	const plain = (struck?.[1] ?? label)
		.replaceAll(/\\([^A-Za-z0-9])/g, '$1')
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
	return struck ? `<s>${plain}</s>` : plain
}

/**
 * Every `.md` file under `dir`, as paths relative to it, `/`-separated.
 * @param {string} dir
 * @param {string} [prefix]
 * @returns {string[]}
 */
function markdownUnder(dir, prefix = '') {
	return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
		const rel = `${prefix}${entry.name}`
		if (entry.isDirectory()) return markdownUnder(path.join(dir, entry.name), `${rel}/`)
		return entry.isFile() && entry.name.endsWith('.md') ? [rel] : []
	})
}

/**
 * The API reference sidebar. Throws, listing every problem, when a module `README.md` links a
 * page that does not exist, or when a generated page is linked by neither the sidebar nor its
 * module's `README.md`.
 * @param {string} outDir - the generated reference, `docs/reference/api`
 * @param {{ name: string, specifier: string }[]} modules - in the order the landing page lists them
 * @returns {SidebarItem[]}
 */
export function apiSidebar(outDir, modules) {
	const docsDir = path.dirname(path.dirname(outDir))
	/** @param {string} file - absolute path of a generated page */
	const route = (file) => `/${routeForPage(path.relative(docsDir, file).split(path.sep).join('/'))}`
	/** @type {string[]} */
	const problems = []
	/** @type {SidebarItem[]} */
	const sidebar = [{ text: 'API reference', link: route(path.join(outDir, 'index.md')) }]

	for (const { name, specifier } of modules) {
		const moduleDir = path.join(outDir, name)
		const readme = path.join(moduleDir, 'README.md')
		/** @type {Set<string>} */
		const reached = new Set(['README.md'])
		/** @type {SidebarItem[]} */
		const kinds = []
		/** @type {SidebarItem | undefined} */
		let kind
		for (const line of readFileSync(readme, 'utf8').split('\n')) {
			const heading = /^## (.+)$/.exec(line)
			if (heading?.[1]) {
				kind = { text: heading[1].trim(), collapsed: true, items: [] }
				kinds.push(kind)
				continue
			}
			const link = /^- \[(.+)\]\(([^)\s#]+\.md)\)$/.exec(line)
			if (!kind?.items || !link?.[1] || !link[2]) continue
			const rel = path.posix.normalize(link[2])
			const file = path.join(moduleDir, ...rel.split('/'))
			reached.add(rel)
			if (!existsSync(file)) {
				problems.push(`${name}/README.md links ${link[2]}, which was not generated`)
				continue
			}
			kind.items.push({ text: sidebarText(link[1]), link: route(file) })
		}
		for (const rel of markdownUnder(moduleDir)) {
			if (!reached.has(rel)) problems.push(`${name}/${rel} is linked by neither the sidebar nor ${name}/README.md`)
		}
		sidebar.push({
			text: specifier,
			link: route(readme),
			collapsed: true,
			// A section with no member list (`## Example`, `## References`) has nothing to navigate to.
			items: kinds.filter((group) => group.items && group.items.length > 0),
		})
	}

	if (problems.length > 0) throw new Error(problems.join('\n'))
	return sidebar
}
