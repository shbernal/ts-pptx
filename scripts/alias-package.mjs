#!/usr/bin/env node
// Stage a publishable copy of this package under its scoped alias name, `@shbernal/ts-pptx`.
//
// npm has one package per name and no concept of a redirect, so an "alias" on the
// registry is not a pointer: it is a second publish of the same content under a second
// name. This script produces that second copy.
//
// It produces it by *copying* rather than by editing the working tree. The obvious
// alternative — rewrite `package.json#name`, publish, put it back — leaves a repo whose
// manifest says `@shbernal/ts-pptx` if anything between the two steps fails, and the thing
// between them is a network call. A staged directory has no such half state, and it can be
// read before it is published, which the bootstrap path below actually wants.
//
// Three things differ from the canonical package, and nothing else does:
//   - `name`, which is the whole point.
//   - `README.md`, which gains a banner naming the unscoped package as canonical. npm
//     renders the README on the package page, and an unannotated copy of this one opens
//     by telling the reader to install a package they did not search for.
//   - `scripts`, which is dropped. Not for tidiness: every script in it names `src/`,
//     `test/` or `scripts/`, none of which the staged copy contains, so all of them are
//     dead here — and one is worse than dead. `prepack` runs `tsdown`, which searches
//     *upward* for its config and finds the repository's, whose `clean: true` then
//     deletes the staged `dist/` before failing on entries it cannot resolve. An
//     `npm publish` here that forgets `--ignore-scripts` therefore does not merely fail:
//     it guts the directory, and the obvious retry *with* the flag publishes a package
//     containing no code. Removing the field means no lifecycle script can fire under
//     any command anyone types, and `--ignore-scripts` in the documented commands stops
//     being load-bearing. `devDependencies` stays: npm never acts on it from a tarball,
//     and it names no path.
// `version`, `exports`, `files`, `dependencies` and the `dist/` and `skills/` payloads
// are the same bytes, so the alias ships what the repo's gates already proved rather than
// a near-copy of it.
//
// Two callers:
//   - `.github/workflows/publish.yml`, which stages and publishes the alias *after* the
//     canonical publish has succeeded, so a failure here cannot cost the real release.
//   - the one-time bootstrap of a name nobody has published yet. npm cannot configure a
//     trusted publisher for a package that does not exist, because the setting lives on
//     the package's settings page (npm/cli#8544), so such a name needs one manual
//     `npm publish` to bring it into being before OIDC can take over. Both names this
//     repo publishes are long past that point; `docs/RELEASING.md` has the runbook.
//
// Usage:
//   node scripts/alias-package.mjs                    stage into .tmp/alias-package
//   node scripts/alias-package.mjs --out <dir>        stage somewhere else
//   node scripts/alias-package.mjs --version 0.0.1    override the version (bootstrap only)
//   node scripts/alias-package.mjs --print-name       print the alias name and exit

import fs from 'node:fs/promises'
import path from 'node:path'
import { ROOT, isMain, parseCliOrExit, run, runCli } from './script-utils.mjs'

/**
 * The scoped name, which the project published under until the unscoped one took over as
 * canonical. It keeps being published so existing installs keep resolving.
 *
 * Single source of truth: the publish workflow reads it back with `--print-name` rather
 * than repeating the literal, so the "is this version already published?" guard and the
 * thing it guards can never name different packages.
 */
export const ALIAS_NAME = '@shbernal/ts-pptx'

const DEFAULT_OUT = path.join('.tmp', 'alias-package')

/**
 * Copied verbatim. `package.json#files` is `["dist", "skills"]`, and npm always adds
 * `package.json`, `README.md` and the licence on top of it — those three are the whole
 * published surface, and the first two are written rather than copied.
 */
const PAYLOAD = ['dist', 'skills', 'LICENSE']

/**
 * The canonical manifest with the alias name substituted in place and `scripts` removed.
 *
 * Assigning over an existing key leaves it where it was, so apart from that deletion the
 * alias manifest differs from the canonical one on one line rather than in key order. See
 * the header for why `scripts` cannot come along.
 * @param {Record<string, any>} manifest the parsed canonical `package.json`
 * Both overrides admit an explicit `undefined` — they are read with `??` and truthiness, and the
 * caller assembles them from CLI flags it may not have been given, so an absent key and a
 * `undefined` one say the same thing here.
 * @param {{name?: string | undefined, version?: string | undefined}} [overrides]
 * @returns {Record<string, any>}
 */
export function aliasManifest(manifest, overrides = {}) {
	// Annotated because spreading a `Record<string, any>` and then naming one key narrows
	// the result to that key alone, which loses `version` on the next line.
	/** @type {Record<string, any>} */
	const aliased = { ...manifest, name: overrides.name ?? ALIAS_NAME }
	if (overrides.version) aliased.version = overrides.version
	delete aliased.scripts
	return aliased
}

/**
 * The canonical README with a banner that says which package a reader is looking at.
 *
 * Placed after the `# ` title rather than above it, so the npm page still opens with the
 * project name; a document whose first element is a disclaimer reads as a warning about
 * the software instead of a note about the name. If there is no title the banner goes
 * first, because being unplaced is worse than being ugly.
 * @param {string} readme the canonical `README.md`
 * @param {{aliasName?: string, canonicalName: string}} names
 * @returns {string}
 */
export function aliasReadme(readme, { aliasName = ALIAS_NAME, canonicalName }) {
	const banner =
		`> **\`${aliasName}\` is an alias.** It is the same package as\n` +
		`> [\`${canonicalName}\`](https://www.npmjs.com/package/${canonicalName}), published from the\n` +
		'> same commit at the same version with the same contents. Install one or the other,\n' +
		'> never both: two copies of this library in one dependency tree are two module\n' +
		'> registries, and state such as the diagnostic handler is per-copy.\n' +
		`> \`${canonicalName}\` is the canonical name, and the issue tracker, the changelog and\n` +
		'> every example below use it.\n'

	const lines = readme.split('\n')
	const title = lines.findIndex((line) => line.startsWith('# '))
	if (title < 0) return banner + '\n' + readme
	const head = lines.slice(0, title + 1).join('\n')
	const tail = lines.slice(title + 1).join('\n')
	return head + '\n\n' + banner + tail
}

/**
 * Write the staged package.
 * @param {{out?: string, version?: string}} [options]
 * @returns {Promise<{dir: string, name: string, version: string}>}
 */
export async function stageAlias({ out = DEFAULT_OUT, version } = {}) {
	// The payload is `dist/`, so the freshness guard runs here rather than being sequenced
	// by the caller. It is a no-op when `dist/` is current, which in the publish workflow
	// it always is.
	await run(process.execPath, [path.join(ROOT, 'scripts', 'ensure-dist.mjs')])

	const dir = path.resolve(ROOT, out)
	if (dir === ROOT) throw new Error('refusing to stage the alias over the repository root')
	await fs.rm(dir, { recursive: true, force: true })
	await fs.mkdir(dir, { recursive: true })

	for (const entry of PAYLOAD) {
		await fs.cp(path.join(ROOT, entry), path.join(dir, entry), { recursive: true })
	}

	const manifest = aliasManifest(JSON.parse(await fs.readFile(path.join(ROOT, 'package.json'), 'utf8')), { version })
	// Tab-indented with a trailing newline, matching the repo's own manifest and oxfmt.
	await fs.writeFile(path.join(dir, 'package.json'), JSON.stringify(manifest, null, '\t') + '\n')

	const readme = await fs.readFile(path.join(ROOT, 'README.md'), 'utf8')
	const canonical = JSON.parse(await fs.readFile(path.join(ROOT, 'package.json'), 'utf8')).name
	await fs.writeFile(path.join(dir, 'README.md'), aliasReadme(readme, { canonicalName: canonical }))

	return { dir, name: manifest.name, version: manifest.version }
}

const USAGE = `Usage: node scripts/alias-package.mjs [--out <dir>] [--version <version>] [--print-name]

Stages a publishable copy of this package under the scoped alias name.

  --out <dir>          where to stage (default ${DEFAULT_OUT})
  --version <version>  publish version override; bootstrap only
  --print-name         print the alias name and exit`

/** @returns {Promise<number>} */
async function main() {
	const { values } = parseCliOrExit(process.argv.slice(2), {
		options: {
			out: { type: 'string' },
			version: { type: 'string' },
			'print-name': { type: 'boolean', default: false },
		},
		usage: USAGE,
	})

	if (values['print-name']) {
		console.log(ALIAS_NAME)
		return 0
	}

	const staged = await stageAlias({ out: values.out, version: values.version })
	console.log(`Staged ${staged.name}@${staged.version} in ${path.relative(ROOT, staged.dir) || staged.dir}`)
	console.log(`Publish it with: npm publish ${path.relative(ROOT, staged.dir)} --access public --ignore-scripts`)
	return 0
}

if (isMain(import.meta.url)) await runCli(main)
