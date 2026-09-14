/**
 * Draws a mermaid graph to SVG, in the browser.
 *
 * `mermaid` is imported on the first call rather than at the top of the file, so it is a
 * chunk of its own that only a page with a diagram fetches. Whether it stays off every other
 * page also depends on where this file is imported from: see `www/theme/index.ts`.
 */

let pending: Promise<unknown> = Promise.resolve()
let renders = 0

/**
 * The SVG markup for one graph, in mermaid's dark or default theme.
 *
 * Calls run one at a time. `initialize` sets mermaid's configuration for the whole page, so
 * two diagrams drawn concurrently across a colour mode flip could each take the other's theme.
 */
export function renderMermaid(graph: string, dark: boolean): Promise<string> {
	const drawn = pending.then(async () => {
		const { default: mermaid } = await import('mermaid')
		mermaid.initialize({
			startOnLoad: false,
			securityLevel: 'strict',
			// A graph that does not parse throws, and the component shows the message, rather
			// than mermaid also inserting its own syntax-error diagram into the page.
			suppressErrorRendering: true,
			theme: dark ? 'dark' : 'default',
		})
		// A fresh id per call: mermaid renders into a temporary element under that id, and
		// a redraw must not find the previous one.
		renders += 1
		const { svg } = await mermaid.render(`mermaid-diagram-${renders}`, graph)
		return svg
	})
	pending = drawn.catch(() => undefined)
	return drawn
}
