// The API reference sidebar is built from each module's `README.md`, in the order the landing
// page lists the entry points. It has to fail rather than return a partial sidebar, because a
// page it leaves out is reachable only by a link, and no gate would notice.

import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, test } from 'vitest'
import { apiSidebar } from '../../scripts/docs-api-sidebar.mjs'

/** A throwaway `docs/reference/api` tree. */
let root = ''
let outDir = ''

/** @param {string} rel @param {string} [content] */
function write(rel, content = '# page\n') {
	const file = path.join(outDir, ...rel.split('/'))
	mkdirSync(path.dirname(file), { recursive: true })
	writeFileSync(file, content)
}

const modules = [
	{ name: 'index', specifier: 'pptx-ts' },
	{ name: 'read', specifier: 'pptx-ts/read' },
]

beforeEach(() => {
	root = mkdtempSync(path.join(os.tmpdir(), 'docs-api-sidebar-'))
	outDir = path.join(root, 'docs', 'reference', 'api')
	write('index.md')
	write(
		'index/README.md',
		[
			'# pptx-ts',
			'',
			'## Interfaces',
			'',
			'- [SlideProps](interfaces/SlideProps.md)',
			'',
			'## Classes',
			'',
			'- [TsPptx](classes/TsPptx.md)',
			'- [~~Old\\_Name~~](classes/Old_Name.md)',
			'',
		].join('\n')
	)
	write('index/interfaces/SlideProps.md')
	write('index/classes/TsPptx.md')
	write('index/classes/Old_Name.md')
	write(
		'read/README.md',
		['# pptx-ts/read', '', '## Functions', '', '- [readPptx](functions/readPptx.md)', '', '## References', ''].join(
			'\n'
		)
	)
	write('read/functions/readPptx.md')
})

afterEach(() => rmSync(root, { force: true, recursive: true }))

describe('apiSidebar', () => {
	test('opens with the landing page, then one group per module in the order given', () => {
		// `read` sorts after `index` on disk; the order given must win.
		const sidebar = apiSidebar(outDir, [...modules].reverse())
		expect(sidebar.map((item) => item.text)).toEqual(['API reference', 'pptx-ts/read', 'pptx-ts'])
		expect(sidebar[0]).toEqual({ text: 'API reference', link: '/reference/api/' })
	})

	test('groups each module by kind, in README order, with routes and collapsed groups', () => {
		const [, index, read] = apiSidebar(outDir, modules)
		expect(index).toEqual({
			text: 'pptx-ts',
			link: '/reference/api/index/README',
			collapsed: true,
			items: [
				{
					text: 'Interfaces',
					collapsed: true,
					items: [{ text: 'SlideProps', link: '/reference/api/index/interfaces/SlideProps' }],
				},
				{
					text: 'Classes',
					collapsed: true,
					items: [
						{ text: 'TsPptx', link: '/reference/api/index/classes/TsPptx' },
						{ text: '<s>Old_Name</s>', link: '/reference/api/index/classes/Old_Name' },
					],
				},
			],
		})
		// A section with no member list is left out.
		expect(read?.items?.map((item) => item.text)).toEqual(['Functions'])
	})

	test('throws when a README links a page that was not generated', () => {
		rmSync(path.join(outDir, 'index', 'classes', 'TsPptx.md'))
		expect(() => apiSidebar(outDir, modules)).toThrow(/index\/README\.md links classes\/TsPptx\.md/)
	})

	test('throws when a generated page is reached by neither the sidebar nor the README', () => {
		write('read/functions/orphan.md')
		expect(() => apiSidebar(outDir, modules)).toThrow(/read\/functions\/orphan\.md is linked by neither/)
	})
})
