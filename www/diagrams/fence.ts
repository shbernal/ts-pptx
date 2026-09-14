/**
 * `mermaid` fences, on the markdown side.
 *
 * A fenced block whose language is `mermaid` renders as `<MermaidDiagram>` instead of a
 * highlighted code block. The graph travels as a URI-encoded attribute: markdown output is
 * compiled as a Vue template, so a `{{` or a `<` left in the graph would be read as template
 * syntax. Encoded, the attribute holds nothing but letters, digits and `%-_.!~*'()`.
 *
 * This file runs in the site config under Node and in the component in the browser, so it
 * imports nothing at runtime.
 */
import type { MarkdownRenderer } from 'vitepress'

/** The component a `mermaid` fence becomes, and the name `www/theme` registers it under. */
export const MERMAID_COMPONENT = 'MermaidDiagram'

export function encodeGraph(graph: string): string {
	return encodeURIComponent(graph)
}

export function decodeGraph(encoded: string): string {
	return decodeURIComponent(encoded)
}

/** A markdown-it plugin: route `mermaid` fences to the component, every other fence onward. */
export function mermaidFences(md: MarkdownRenderer): void {
	const fence = md.renderer.rules.fence
	md.renderer.rules.fence = (tokens, index, options, env, self) => {
		const token = tokens[index]
		if (token.info.trim().split(/\s+/, 1)[0] === 'mermaid') {
			return `<${MERMAID_COMPONENT} graph="${encodeGraph(token.content)}" />\n`
		}
		return fence ? fence(tokens, index, options, env, self) : self.renderToken(tokens, index, options)
	}
}
