import eslint from '@eslint/js'
import prettier from 'eslint-config-prettier/flat'
import tseslint from 'typescript-eslint'

/** @type {Record<string, 'readonly'>} */
const nodeGlobals = {
	Blob: 'readonly',
	Buffer: 'readonly',
	clearInterval: 'readonly',
	clearTimeout: 'readonly',
	console: 'readonly',
	process: 'readonly',
	setInterval: 'readonly',
	setTimeout: 'readonly',
	structuredClone: 'readonly',
	TextDecoder: 'readonly',
	TextEncoder: 'readonly',
	URL: 'readonly',
	URLSearchParams: 'readonly',
}

export default tseslint.config(
	{
		ignores: [
			'coverage/**',
			'demos/**',
			'dist/**',
			'node_modules/**',
			'src/bld/**',
			'tools/ooxml-validator/bin/**',
			'types/**',
		],
	},
	{
		files: ['src/**/*.ts'],
		extends: [eslint.configs.recommended, tseslint.configs.recommendedTypeChecked],
		languageOptions: {
			parserOptions: {
				projectService: true,
				tsconfigRootDir: import.meta.dirname,
			},
		},
		rules: {
			// Gate both compile-time escape hatches from the null-safety work: a bare `!`
			// and a provably-redundant `as` cast. no-unnecessary-type-assertion already ships
			// on in recommendedTypeChecked; it is pinned here so the `!`/`as` symmetry is
			// explicit and survives any upstream preset change.
			'@typescript-eslint/no-non-null-assertion': 'error',
			'@typescript-eslint/no-unnecessary-type-assertion': 'error',
			// --- type-aware rules intentionally relaxed for this codebase ---
			// Several async methods exist only to satisfy a uniform Promise-returning contract
			// (runtime adapters, zip/opc save, excel worksheet) even when a given impl has no await.
			'@typescript-eslint/require-await': 'off',
			// Output paths deliberately coerce `unknown`/union values with String()/.toString() while
			// assembling OOXML strings; the object-stringification guard is noise here.
			'@typescript-eslint/no-base-to-string': 'off',
			// Public color types are `literal-union | string` on purpose: the literals drive editor
			// autocomplete while `string` keeps an escape hatch for arbitrary hex values.
			'@typescript-eslint/no-redundant-type-constituents': 'off',
			'no-lone-blocks': 0,
			// Library output goes through `warn()` / `warnOnce()` so a consumer can route or silence
			// it (see diagnostics.ts), and failures go through the error classes. A direct `console.*`
			// is neither: it cannot be captured, muted, or branched on. Two files are exempt below;
			// everything else that reaches for one wants a diagnostic code or a thrown error instead.
			'no-console': 'error',
		},
	},
	{
		// The two exemptions to `no-console`, both deliberate.
		//
		// `diagnostics.ts` owns the default handler -- the one `console.warn` every diagnostic
		// funnels into when the consumer has installed nothing else.
		//
		// The `verbose: true` table tracers are a different kind of output from a diagnostic: a
		// DEV-ONLY flag (see `TableProps.verbose`) that prints a multi-line trace of the auto-paging
		// arithmetic. Routing it through the diagnostic handler would mean inventing a code per
		// formatted line and flooding a consumer's handler with output that reports no condition.
		// A consumer silences it by not passing the flag, so it is opt-in, not unroutable.
		files: ['src/diagnostics.ts', 'src/gen/table/autopage.ts', 'src/gen/table/html-dom.ts'],
		rules: { 'no-console': 'off' },
	},
	{
		files: ['scripts/**/*.mjs', 'test/**/*.mjs', 'test/**/*.js'],
		extends: [eslint.configs.recommended],
		languageOptions: {
			ecmaVersion: 2024,
			globals: nodeGlobals,
			sourceType: 'module',
		},
	},
	{
		// The browser lane is the one place under `test/` whose code runs in a page, and the
		// two halves of it need different amounts of DOM.
		//
		// A **spec** is ordinary Node except for its `page.evaluate` callbacks, which are
		// serialized and evaluated in the page. Those reach `window` and nothing else here
		// grants more, deliberately: a reference to `document` or `fetch` outside an evaluate
		// callback would be a mistake worth catching.
		files: ['test/browser/**/*.mjs'],
		languageOptions: { globals: { ...nodeGlobals, window: 'readonly' } },
	},
	{
		// A **harness module** is page code outright — loaded by its own document over a
		// `<script type="module">`, never imported by Node. Reading the DOM is its whole job
		// (`table.mjs` renders a fixture and measures it), so withholding `document` there
		// would only mean writing the same access in a form the linter cannot see.
		//
		// Kept to this directory rather than folded into the block above so the distinction
		// survives: the specs still cannot touch the DOM outside an evaluate callback.
		files: ['test/browser/harness/**/*.mjs'],
		languageOptions: {
			globals: {
				...nodeGlobals,
				document: 'readonly',
				getComputedStyle: 'readonly',
				window: 'readonly',
			},
		},
	},
	{
		// The root build configs matched no `files` block at all, so zero rules
		// applied to them. Untyped `recommended` only: these are a handful of
		// declarative config objects, and turning on type-checked rules would mean
		// paying `projectService` for them.
		files: ['eslint.config.mjs'],
		extends: [eslint.configs.recommended],
		languageOptions: {
			ecmaVersion: 2024,
			globals: nodeGlobals,
			sourceType: 'module',
		},
	},
	{
		// Same intent as the block above, but these are TypeScript and so need the
		// TS parser; espree cannot parse `as`/`type` syntax. Still untyped rules.
		files: ['vitest.config.ts', 'tsdown.config.ts', 'tsdown.dev.config.ts', 'playwright.config.ts'],
		extends: [tseslint.configs.recommended],
		languageOptions: {
			ecmaVersion: 2024,
			globals: nodeGlobals,
			sourceType: 'module',
		},
	},
	// Must be last: turns off any ESLint rules that would conflict with Prettier,
	// which is now the sole formatter of record (see .prettierrc.json).
	prettier
)
