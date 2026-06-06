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
		extends: [eslint.configs.recommended, tseslint.configs.recommended],
		plugins: {
			'@stylistic': stylistic,
		},
		rules: {
			'@stylistic/comma-dangle': ['error', 'only-multiline'],
			'@stylistic/indent': ['error', 'tab', { SwitchCase: 1, ImportDeclaration: 1 }],
			'@stylistic/no-tabs': ['error', { allowIndentationTabs: true }],
			'@stylistic/quotes': ['error', 'single'],
			'@stylistic/semi': ['error', 'never'],
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
