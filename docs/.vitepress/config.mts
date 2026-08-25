import { readFileSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { defineConfig } from 'vitepress'

const configDir = path.dirname(fileURLToPath(import.meta.url))
const docsDir = path.resolve(configDir, '..')
const docsConfig = JSON.parse(readFileSync(path.join(docsDir, 'docs.json'), 'utf8')) as {
	description: string
	name: string
	navigation: Array<{ group: string; pages: string[] }>
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

export default defineConfig({
	base: process.env.VITEPRESS_BASE ?? '/ts-pptx/',
	cleanUrls: true,
	description: docsConfig.description,
	lang: 'en-US',
	title: docsConfig.name,
	themeConfig: {
		nav: [
			{ text: 'Guide', link: '/project-target' },
			{ text: 'Demos', link: '/demos' },
			{ text: 'API', link: '/reference/api/' },
			{ text: 'GitHub', link: 'https://github.com/shbernal/ts-pptx' },
		],
		search: {
			provider: 'local',
		},
		sidebar,
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
