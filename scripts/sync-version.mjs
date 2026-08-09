#!/usr/bin/env node
// Rewrites the `VERSION` constant in `src/presentation.ts` from `package.json`.
//
// `package.json` is the version of record: `publish.yml` refuses to publish unless the
// git tag matches it, and the constant exists only so `pres.version` can report it back
// to a consumer. The constant is therefore derived — but it used to be derived by hand,
// so cutting a release meant editing two files and remembering the second one.
//
// It could not ship wrong. `test/regression/api/public-accessors.test.js` pins
// `pres.version` to the manifest, and that test runs in `verify`, which the publish
// workflow runs as its `CI gate` job. A forgotten bump cost a red CI round trip, not a
// mis-reported library. This script removes the round trip; the test stays as the
// backstop, because a script nobody is forced to run guarantees nothing on its own.
//
// Wired as the `version` lifecycle script, so `pnpm version minor` bumps the manifest,
// runs this, and puts the rewritten source in the same version commit. `--stage` is what
// puts it there: pnpm commits `package.json` itself but stages nothing else.
//
// Usage:
//   node scripts/sync-version.mjs           rewrite the constant if it is stale
//   node scripts/sync-version.mjs --stage   ...and `git add` it (the `version` hook)
//   node scripts/sync-version.mjs --check   report drift, write nothing, exit 1

import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { ROOT, isMain, parseCliOrExit, runCli } from './script-utils.mjs'

const SOURCE = 'src/presentation.ts'

// Anchored to a whole line so it cannot match a mention of the constant inside a comment
// or a string. Exactly one match is required — see `replaceVersion`.
const PATTERN = /^const VERSION = '[^']*'$/gm
const CAPTURE = /^const VERSION = '([^']*)'$/m

const USAGE = `Usage: node scripts/sync-version.mjs [--check] [--stage]

Rewrites the VERSION constant in ${SOURCE} from package.json.

  --check   report drift and exit 1; write nothing
  --stage   git add the file after rewriting it (used by the \`version\` lifecycle script)`

/**
 * Point the `VERSION` constant at `version`, and report what it said before.
 *
 * Requires exactly one match rather than replacing whatever it finds: zero means the
 * constant was renamed or moved and this script would otherwise silently do nothing,
 * which is the failure mode the whole exercise is meant to remove.
 * @param {string} source contents of `src/presentation.ts`
 * @param {string} version the version `package.json` declares
 * @returns {{text: string, previous: string}}
 */
export function replaceVersion(source, version) {
	const count = source.match(PATTERN)?.length ?? 0
	const previous = count === 1 ? CAPTURE.exec(source)?.[1] : undefined
	if (previous === undefined) {
		throw new Error(
			`expected exactly one \`const VERSION = '…'\` line in ${SOURCE}, found ${count}. ` +
				'If the constant was renamed or moved, update PATTERN in scripts/sync-version.mjs.'
		)
	}
	return { text: source.replace(PATTERN, `const VERSION = '${version}'`), previous }
}

/** @returns {number} process exit code */
function main() {
	const { values } = parseCliOrExit(process.argv.slice(2), {
		options: {
			check: { type: 'boolean', default: false },
			stage: { type: 'boolean', default: false },
		},
		usage: USAGE,
	})

	const version = String(JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version)
	const sourcePath = path.join(ROOT, SOURCE)
	const { text, previous } = replaceVersion(fs.readFileSync(sourcePath, 'utf8'), version)

	if (previous === version) {
		console.log(`${SOURCE} already reports ${version}.`)
		return 0
	}

	if (values.check) {
		console.error(
			`${SOURCE} reports ${previous}, package.json says ${version}.\n` +
				'Run `pnpm run version:sync`, or bump with `pnpm version <major|minor|patch>` so both move together.'
		)
		return 1
	}

	fs.writeFileSync(sourcePath, text)
	console.log(`${SOURCE}: ${previous} -> ${version}`)
	// Not `run()` from script-utils: that appends `.cmd` to any bare command on Windows,
	// which is right for the pnpm/npm shims and wrong for git, a real executable.
	if (values.stage) execFileSync('git', ['add', '--', SOURCE], { cwd: ROOT, stdio: 'inherit' })
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
