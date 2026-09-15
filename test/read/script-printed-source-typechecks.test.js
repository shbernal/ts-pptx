// A printed script type-checks against the package it imports.
//
// `script-print.test.js` runs printed scripts, and running one says nothing about its types: Node
// strips them before it executes anything. A consumer who opens a printed script in an editor, or
// runs `tsc` over it, sees every option the write API's declarations reject. Two such defects went
// unseen that way: a table's source style GUID printed as a string `tableStyle` refused, and table
// cells and their runs carrying run and paragraph options, such as `strike` and
// `lineSpacingMultiple`, which the emitter writes and `TableCellProps` did not declare.
//
// So every corpus deck is printed by both tiers, and the whole set is checked in one program against
// the built declarations that the printed `import ... from 'pptx-ts'` resolves to. The scripts are
// never written: a compiler host serves them from memory, at paths inside the repository so the
// package resolves itself by name through its own `exports`.

import path from 'node:path'
import { describe, test } from 'vitest'
import ts from 'typescript-6'
import { printScript, printStandaloneScript } from '../../dist/script.js'
import { assert, assertEqual } from '../helpers.js'
import { SCRATCH, fixtureNames, irFor } from './corpus.js'

/** Where the in-memory scripts claim to live: inside the package, so `pptx-ts` resolves to itself. */
const ROOT = path.join(SCRATCH, 'printed-typecheck').replaceAll('\\', '/')

/** The settings a consumer running a printed script under Node would type-check it with. */
const OPTIONS = {
	strict: true,
	noEmit: true,
	target: ts.ScriptTarget.ES2022,
	module: ts.ModuleKind.NodeNext,
	moduleResolution: ts.ModuleResolutionKind.NodeNext,
	skipLibCheck: true,
	types: ['node'],
}

/**
 * Every diagnostic one program over `files` reports, as `TS<code> <file>: <message>`.
 * @param {Map<string, string>} files - source text by path, every path under {@link ROOT}
 * @returns {string[]}
 */
function typeErrors(files) {
	const host = ts.createCompilerHost(OPTIONS)
	const { getSourceFile, fileExists, readFile } = host
	host.getSourceFile = (fileName, languageVersion, onError, shouldCreate) => {
		const code = files.get(fileName)
		return code === undefined
			? getSourceFile.call(host, fileName, languageVersion, onError, shouldCreate)
			: ts.createSourceFile(fileName, code, languageVersion, true)
	}
	host.fileExists = (fileName) => files.has(fileName) || fileExists.call(host, fileName)
	host.readFile = (fileName) => files.get(fileName) ?? readFile.call(host, fileName)
	const program = ts.createProgram([...files.keys()], OPTIONS, host)
	return ts
		.getPreEmitDiagnostics(program)
		.map(
			(diagnostic) =>
				`TS${diagnostic.code} ${path.basename(diagnostic.file?.fileName ?? '(options)')}: ` +
				ts.flattenDiagnosticMessageText(diagnostic.messageText, '\n')
		)
}

describe('a printed script type-checks', () => {
	test('every corpus deck, printed by both tiers', { timeout: 120_000 }, async () => {
		const files = new Map()
		for (const name of fixtureNames) {
			const ir = await irFor(name)
			const stem = name.replace(/[^\w.-]/g, '_')
			files.set(`${ROOT}/${stem}.anchored.ts`, printScript(ir).code)
			files.set(`${ROOT}/${stem}.standalone.ts`, printStandaloneScript(ir).code)
		}
		assertEqual(typeErrors(files).join('\n'), '', 'no printed script has a type error')
	})

	// A clean run above is evidence only if this setup can fail. A module that did not resolve would
	// be TS2307 rather than TS2322, so the one error asserted here also proves `pptx-ts` resolved.
	test('the same check rejects a wrong option', { timeout: 60_000 }, () => {
		const code = [
			"import type { TableProps } from 'pptx-ts'",
			"export const props: TableProps = { colW: 'wide' }",
		].join('\n')
		const errors = typeErrors(new Map([[`${ROOT}/wrong-option.ts`, code]]))
		assertEqual(errors.length, 1, `exactly one error; got:\n${errors.join('\n')}`)
		assert(errors[0]?.startsWith('TS2322 '), `the error is the rejected option; got: ${errors[0]}`)
	})
})
