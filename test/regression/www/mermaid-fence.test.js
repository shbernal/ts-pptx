import { fileURLToPath } from 'node:url'
import { createMarkdownRenderer } from 'vitepress'
import { beforeAll, describe, expect, it } from 'vitest'
import siteConfig from '../../../docs/.vitepress/config.mts'
import { decodeGraph, MERMAID_COMPONENT } from '../../../www/diagrams/fence.ts'

/**
 * `mermaid` fences, through the site's own markdown configuration.
 *
 * The renderer is VitePress's, built from the `markdown` options in
 * `docs/.vitepress/config.mts`, so the wiring is asserted along with the rule: a rule that is
 * never installed, or that VitePress's own fence handling wraps, fails here. Drawing the
 * diagram needs a browser and is not covered.
 */

const docsDir = fileURLToPath(new URL('../../../docs/', import.meta.url))

/** @type {import('vitepress').MarkdownRenderer} */
let md

beforeAll(async () => {
	md = await createMarkdownRenderer(docsDir, siteConfig.markdown, '/')
})

/**
 * The encoded `graph` attribute of the diagram component in `html`.
 * @param {string} html
 * @returns {string}
 */
function graphAttribute(html) {
	const match = html.match(new RegExp(`<${MERMAID_COMPONENT} graph="([^"]*)" />`))
	if (!match) throw new Error(`no <${MERMAID_COMPONENT}> in: ${html}`)
	return match[1]
}

describe('mermaid fences', () => {
	it('become the diagram component, carrying the graph unchanged', () => {
		const graph = 'flowchart LR\n  A["a {{ b }} <c>"] --> B[élan]\n'
		const html = md.render(`\`\`\`mermaid\n${graph}\`\`\`\n`, {})

		const encoded = graphAttribute(html)
		expect(decodeGraph(encoded)).toBe(graph)
		// Markdown output is compiled as a Vue template, where a `{{` or a `<` left in the
		// attribute would be read as template syntax.
		expect(encoded).toMatch(/^[\w%.!~*'()-]*$/)
		expect(html).not.toContain('language-mermaid')
	})

	it('leave every other fence to VitePress', () => {
		const html = md.render('```ts\nconst answer = 42\n```\n', {})

		expect(html).toContain('language-ts')
		expect(html).not.toContain(MERMAID_COMPONENT)
	})
})
