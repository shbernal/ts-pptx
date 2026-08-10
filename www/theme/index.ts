/**
 * The site's VitePress theme.
 *
 * It extends the default theme rather than replacing it: the docs are the bulk of the
 * site, and the default layout is what makes them navigable. What is added here is a
 * palette (`style.css`) and one component.
 *
 * `<DeckPreview />` is registered **asynchronously** on purpose. It pulls in two copies of
 * the library and the whole of `pptx-html`; a synchronous import would put all of that in
 * the chunk every page of the site loads, to serve one page.
 */
import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'
import { defineAsyncComponent } from 'vue'
import './style.css'

export default {
	extends: DefaultTheme,
	enhanceApp({ app }) {
		app.component(
			'DeckPreview',
			defineAsyncComponent(() => import('../demos/DeckPreview.vue'))
		)
	},
} satisfies Theme
