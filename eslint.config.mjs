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
		},
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
		files: ['vitest.config.ts', 'tsdown.config.ts', 'tsdown.dev.config.ts'],
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
