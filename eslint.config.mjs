import eslint from '@eslint/js'
import stylistic from '@stylistic/eslint-plugin'
import tseslint from 'typescript-eslint'

const nodeGlobals = {
	Blob: 'readonly',
	Buffer: 'readonly',
	clearInterval: 'readonly',
	clearTimeout: 'readonly',
	console: 'readonly',
	process: 'readonly',
	setInterval: 'readonly',
	setTimeout: 'readonly',
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
		plugins: {
			'@stylistic': stylistic,
		},
		languageOptions: {
			parserOptions: {
				projectService: true,
				tsconfigRootDir: import.meta.dirname,
			},
		},
		rules: {
			'@stylistic/comma-dangle': ['error', 'only-multiline'],
			'@stylistic/indent': ['error', 'tab', { SwitchCase: 1, ImportDeclaration: 1 }],
			'@stylistic/no-tabs': ['error', { allowIndentationTabs: true }],
			'@stylistic/quotes': ['error', 'single'],
			'@stylistic/semi': ['error', 'never'],
			'@typescript-eslint/no-non-null-assertion': 'error',
			// --- type-aware rules intentionally relaxed for this codebase ---
			// CHART_NAME (string union) and CHART_TYPE (enum) are parallel definitions with
			// identical string values, and scheme-color checks compare runtime strings to the
			// SCHEME_COLORS enum. Those comparisons are value-safe; unifying the type/enum pairs
			// is a public-API refactor (see STATIC-CHECK-HARDENING.md, Gap 4).
			'@typescript-eslint/no-unsafe-enum-comparison': 'off',
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
		files: ['rollup.config.mjs', 'scripts/**/*.mjs', 'test/**/*.mjs', 'test/**/*.js'],
		extends: [eslint.configs.recommended],
		languageOptions: {
			ecmaVersion: 2024,
			globals: nodeGlobals,
			sourceType: 'module',
		},
	}
)
