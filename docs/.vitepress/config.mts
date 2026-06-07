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
		.split('/')
		.at(-1)!
		.replaceAll('-', ' ')
		.replace(/\b\w/g, char => char.toUpperCase())
}

function routeFor(page: string): string {
	return page.endsWith('/index') ? `/${page.slice(0, -'/index'.length)}/` : `/${page}`
}

const sidebar = docsConfig.navigation.map(group => ({
	text: group.group,
	items: group.pages.map(page => ({
		text: pageTitle(page),
		link: routeFor(page),
	})),
}))

export default defineConfig({
	base: process.env.VITEPRESS_BASE ?? '/PptxGenJS/',
	cleanUrls: true,
	description: docsConfig.description,
	lang: 'en-US',
	title: docsConfig.name,
	themeConfig: {
		nav: [
			{ text: 'Guide', link: '/' },
			{ text: 'API', link: '/reference/api/' },
			{ text: 'GitHub', link: 'https://github.com/shbernal/PptxGenJS' },
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
