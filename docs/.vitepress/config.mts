import { readFileSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import type { DefaultTheme } from 'vitepress'
import { defineConfig } from 'vitepress'
import { mermaidFences } from '../../www/diagrams/fence'

const configDir = path.dirname(fileURLToPath(import.meta.url))
const docsDir = path.resolve(configDir, '..')
const docsConfig = JSON.parse(readFileSync(path.join(docsDir, 'docs.json'), 'utf8')) as {
	description: string
	name: string
	navigation: Array<{ group: string; pages: string[] }>
	/** Docs-relative directories read on GitHub and never built into the site. */
	repoOnly: string[]
}

function pageTitle(page: string): string {
	const filePath = path.join(docsDir, `${page}.md`)
	try {
		const markdown = readFileSync(filePath, 'utf8')
		const frontmatterTitle = markdown.match(/^title:\s+"?(.+?)"?$/m)
		if (frontmatterTitle) return frontmatterTitle[1]
		const heading = markdown.match(/^#\s+(.+)$/m)
		if (heading) return heading[1]
	} catch {
		// Generated pages may not exist until docs:api has run.
	}

	return page
		.slice(page.lastIndexOf('/') + 1)
		.replaceAll('-', ' ')
		.replace(/\b\w/g, (char) => char.toUpperCase())
}

function routeFor(page: string): string {
	return page.endsWith('/index') ? `/${page.slice(0, -'/index'.length)}/` : `/${page}`
}

const sidebar = docsConfig.navigation.map((group) => ({
	text: group.group,
	items: group.pages.map((page) => ({
		text: pageTitle(page),
		link: routeFor(page),
	})),
}))

// `docs:api` writes the API reference's own sidebar next to the pages it generates. Without it the
// API pages fall back to the guide sidebar rather than an empty one.
function readApiSidebar(): DefaultTheme.SidebarItem[] | undefined {
	try {
		return JSON.parse(
			readFileSync(path.join(docsDir, 'reference', 'api', 'sidebar.json'), 'utf8')
		) as DefaultTheme.SidebarItem[]
	} catch {
		return undefined
	}
}

const apiSidebar = readApiSidebar()

export default defineConfig({
	base: process.env.VITEPRESS_BASE ?? '/ts-pptx/',
	cleanUrls: true,
	description: docsConfig.description,
	lang: 'en-US',
	markdown: {
		// `mermaid` fences render as diagrams, drawn in the browser by `www/diagrams/`.
		config: mermaidFences,
	},
	srcExclude: docsConfig.repoOnly.map((dir) => `${dir}/**`),
	title: docsConfig.name,
	themeConfig: {
		nav: [
			{ text: 'Guide', link: '/getting-started/introduction' },
			{ text: 'Demos', link: '/demos' },
			{ text: 'API', link: '/reference/api/' },
			{ text: 'GitHub', link: 'https://github.com/shbernal/ts-pptx' },
		],
		search: {
			provider: 'local',
		},
		// VitePress picks the longest key that prefixes the route, so the API key wins under it.
		sidebar: apiSidebar ? { '/reference/api/': apiSidebar, '/': sidebar } : { '/': sidebar },
	},
	vite: {
		build: {
			chunkSizeWarningLimit: 5000,
		},
		resolve: {
			// This repo declares `vue` itself (for `www/theme`), and VitePress carries its own
			// copy. pnpm resolves those to two versions, and two Vue runtimes in one page means
			// `inject`/`provide` and the app instance stop matching across the boundary — which
			// shows up as a component that mounts but sees none of the theme's context. Naming
			// it here collapses both specifiers onto one copy.
			dedupe: ['vue'],
		},
		esbuild: {
			target: 'es2022',
			tsconfigRaw: {
				compilerOptions: {
					target: 'es2022',
				},
			},
		},
	},
})
