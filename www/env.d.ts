/**
 * Module shapes the bundler understands and `tsc` does not.
 *
 * Only what this directory actually imports. A blanket `declare module '*'` would make
 * every typo in a specifier resolve to `any`, which is the opposite of what typechecking
 * `www/` is for.
 */
declare module '*.css'

declare module '*.vue' {
	import type { DefineComponent } from 'vue'

	const component: DefineComponent
	export default component
}
