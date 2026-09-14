/**
 * The site's VitePress theme.
 *
 * It extends the default theme rather than replacing it: the docs are the bulk of the
 * site, and the default layout is what makes them navigable. What is added here is a
 * palette (`style.css`) and two components.
 *
 * `<DeckPreview />` is registered **asynchronously** on purpose. It pulls in two copies of
 * the library and the whole of `pptx-html`; a synchronous import would put all of that in
 * the chunk every page of the site loads, to serve one page.
 *
 * `<MermaidDiagram />` is what a `mermaid` fence renders as. The component itself is small
 * and registered synchronously; mermaid is imported inside it on first render
 * (`www/diagrams/render.ts`), so it stays out of pages that draw no diagram. That holds only
 * because the import lives in the theme's chunk: VitePress preloads every dynamic import of
 * its *app* chunk on every page, which is how a plugin that registers its component there
 * put all of mermaid on every page.
 */
import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'
import { defineAsyncComponent } from 'vue'
import { MERMAID_COMPONENT } from '../diagrams/fence'
import MermaidDiagram from '../diagrams/MermaidDiagram.vue'
import './style.css'

export default {
	extends: DefaultTheme,
	enhanceApp({ app }) {
		app.component(
			'DeckPreview',
			defineAsyncComponent(() => import('../demos/DeckPreview.vue'))
		)
		app.component(MERMAID_COMPONENT, MermaidDiagram)
	},
} satisfies Theme
