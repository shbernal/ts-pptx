// The bundle-size ratchet's parsing half, which is where it can fail silently.
//
// A ratchet that over-counts fails loudly and gets looked at. One that UNDER-counts
// reports a smaller bundle and passes — indistinguishable from a real win. Every case
// below is about that direction: a specifier form the closure must not miss.

import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { afterAll, describe, expect, test } from 'vitest'
import { closureOf, relativeImportsOf } from '../../scripts/bundle-size-ratchet.mjs'

describe('relativeImportsOf', () => {
	test('finds the three specifier forms tsdown emits', () => {
		const text = [
			`import { a } from './a.js'`,
			`export * from './b.js'`,
			`const c = await import('./c.js')`,
			`import './side-effect.js'`,
		].join('\n')
		expect(relativeImportsOf(text)).toEqual(['./a.js', './b.js', './c.js', './side-effect.js'])
	})

	// The regression this file exists for. A side-effect import names a file that ships
	// and that pulls in its own subtree; missing it silently shrinks the measurement.
	test('finds a side-effect import with no binding', () => {
		expect(relativeImportsOf(`import './polyfill.js'`)).toEqual(['./polyfill.js'])
		expect(relativeImportsOf(`import"./tight.js"`)).toEqual(['./tight.js'])
	})

	test('ignores bare specifiers — those are the consumer to resolve, not bytes we ship', () => {
		const text = [`import { unzip } from 'fflate'`, `import 'core-js/stable'`, `await import('opentype.js')`].join('\n')
		expect(relativeImportsOf(text)).toEqual([])
	})

	test('handles both quote styles and parent-relative paths', () => {
		expect(relativeImportsOf(`import x from "../shared/y.js"`)).toEqual(['../shared/y.js'])
	})

	// The over-counting direction, and the one that turned into a hard failure rather than a
	// wrong number: `dist/shapes-*.js` carries `{@link import('./notes.js')}` in a doc comment,
	// and the closure walked into it demanding a `dist/notes.js` no build emits. It stayed
	// invisible while only the browser entry was budgeted, because that entry never reaches
	// that chunk.
	test('a specifier that is only written ABOUT, in a comment, is not an import', () => {
		const text = [
			`/** See {@link import('./notes.js').NotesPlaceholder} for the shape. */`,
			`import { real } from './real.js'`,
			`// import { gone } from './removed.js'`,
		].join('\n')
		expect(relativeImportsOf(text)).toEqual(['./real.js'])
	})

	// Blanking a comment must not join the lines around it, or two statements become one.
	test('stripping a multi-line comment leaves the statements on either side of it', () => {
		const text = [`import { a } from './a.js'`, `/* a`, `   multi-line`, `   comment */`, `import './b.js'`].join('\n')
		expect(relativeImportsOf(text)).toEqual(['./a.js', './b.js'])
	})

	// Deliberately narrow: a trailing `//` is left alone, because stripping it would also eat
	// the `//` inside any URL on that line -- and a parser that quietly removes code is exactly
	// the under-counting failure this file exists for.
	test('a URL in a string literal is not mistaken for a comment', () => {
		expect(relativeImportsOf(`const u = 'https://example.com/x'; import './after.js'`)).toEqual(['./after.js'])
	})
})

describe('closureOf', () => {
	const dirs = []
	const emit = (files) => {
		const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ts-pptx-closure-'))
		dirs.push(dir)
		for (const [name, text] of Object.entries(files)) {
			fs.mkdirSync(path.join(dir, path.dirname(name)), { recursive: true })
			fs.writeFileSync(path.join(dir, name), text)
		}
		return dir
	}
	afterAll(() => {
		for (const dir of dirs) fs.rmSync(dir, { recursive: true, force: true })
	})

	test('walks transitively and includes the entry itself', () => {
		const dir = emit({
			'entry.js': `import './a.js'\nexport * from './nested/b.js'`,
			'a.js': `import 'fflate'`,
			'nested/b.js': `import '../a.js'`,
		})
		expect(closureOf('entry.js', dir)).toEqual(['a.js', 'entry.js', 'nested/b.js'])
	})

	// Directly the bug: before the fix, `entry.js` reached nothing and the whole
	// subtree below the side-effect import vanished from the measurement.
	test('a side-effect import keeps its subtree in the closure', () => {
		const dir = emit({
			'entry.js': `import './heavy.js'`,
			'heavy.js': `import './heavier.js'`,
			'heavier.js': `export const x = 1`,
		})
		expect(closureOf('entry.js', dir)).toEqual(['entry.js', 'heavier.js', 'heavy.js'])
	})

	test('survives an import cycle rather than looping forever', () => {
		const dir = emit({ 'a.js': `import './b.js'`, 'b.js': `import './a.js'` })
		expect(closureOf('a.js', dir)).toEqual(['a.js', 'b.js'])
	})

	test('a missing file is a hard error, never a silently smaller closure', () => {
		const dir = emit({ 'entry.js': `import './gone.js'` })
		expect(() => closureOf('entry.js', dir)).toThrow(/gone\.js is missing/)
	})
})
